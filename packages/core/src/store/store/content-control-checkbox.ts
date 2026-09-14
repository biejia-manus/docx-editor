import { isInlineRunContainer, WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import { mintCheckboxRun } from './content-control-run.ts';
import { runPropertiesNodeOf } from './tree-op-nodes.ts';

/** The symbol a checkbox state writes, plus every glyph the control's two states declare. */
export interface CheckboxSymbol {
  readonly hex: string;
  readonly font: string;
  /** Hex code points of both states, so the display run is found by what it shows. */
  readonly states: readonly string[];
}

/** Both state glyphs a checkbox declares, with Word's defaults for the ones it leaves out. */
export function checkboxStateHexes(checkbox: {
  readonly checkedState?: { readonly value?: string };
  readonly uncheckedState?: { readonly value?: string };
}): readonly string[] {
  return [checkbox.checkedState?.value ?? '2612', checkbox.uncheckedState?.value ?? '2610'];
}

const BLOCK_KINDS: ReadonlySet<OoxmlNode['kind']> = new Set([
  'paragraph',
  'table',
  'tableRow',
  'tableCell',
]);

/** Zero-length closers that a new run must precede, or a range covering the paragraph loses it. */
const TRAILING_MARKER_KINDS: ReadonlySet<OoxmlNode['kind']> = new Set([
  'bookmarkEnd',
  'commentRangeEnd',
  'moveFromRangeEnd',
  'moveToRangeEnd',
]);

/** Paragraph-mark run properties that describe a revision, not formatting a new run inherits. */
const MARK_REVISION_NAMES: ReadonlySet<string> = new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'rPrChange',
]);

function isWml(node: OoxmlNode, localName: string): boolean {
  return (
    node.kind !== 'textValue' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    node.localName === localName
  );
}

function paragraphOf(nextId: () => string, children: readonly OoxmlNode[]): OoxmlNode {
  return {
    id: nextId(),
    kind: 'paragraph',
    namespaceUri: WML_NAMESPACE_URI,
    localName: 'p',
    prefix: 'w',
    namespaceBindings: [],
    attributes: [],
    children,
  } as unknown as OoxmlNode;
}

function cloneWithFreshIds(node: OoxmlNode, nextId: () => string): OoxmlNode {
  if (node.kind === 'textValue') return { id: nextId(), kind: 'textValue', value: node.value };
  return {
    ...node,
    id: nextId(),
    children: node.children.map((child) => cloneWithFreshIds(child, nextId)),
  } as OoxmlNode;
}

/** The paragraph mark's `w:rPr`, minus revision records, as the formatting a new run inherits. */
function paragraphMarkProperties(
  paragraph: OoxmlElement,
  nextId: () => string
): OoxmlNode | undefined {
  const pPr = paragraph.children.find((child) => isWml(child, 'pPr'));
  if (!pPr || pPr.kind === 'textValue') return undefined;
  const rPr = pPr.children.find((child) => isWml(child, 'rPr'));
  if (!rPr || rPr.kind === 'textValue') return undefined;
  const formatting = rPr.children.filter(
    (child) => child.kind !== 'textValue' && !MARK_REVISION_NAMES.has(child.localName)
  );
  if (formatting.length === 0) return undefined;
  return cloneWithFreshIds({ ...rPr, children: formatting } as OoxmlNode, nextId);
}

function normalizedHex(value: string | undefined): string | null {
  if (value === undefined || !/^[0-9A-Fa-f]{1,6}$/.test(value)) return null;
  return value.toUpperCase().padStart(4, '0');
}

function glyphOf(hex: string): string | null {
  const code = Number.parseInt(hex, 16);
  if (!Number.isInteger(code) || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) {
    return null;
  }
  return String.fromCodePoint(code);
}

/** Whether a run shows one of the control's declared state glyphs and nothing else. */
function displaysState(run: OoxmlElement, hexes: ReadonlySet<string>, glyphs: ReadonlySet<string>) {
  const content = run.children.filter((child) => child.kind !== 'runProperties');
  if (content.length !== 1) return false;
  const only = content[0]!;
  if (only.kind === 'textValue') return false;
  if (isWml(only, 'sym')) {
    const char = normalizedHex(only.attributes.find((a) => a.localName === 'char')?.value);
    return char !== null && hexes.has(char);
  }
  if (only.kind === 'text') {
    const value = only.children
      .map((child) => (child.kind === 'textValue' ? child.value : ''))
      .join('');
    return glyphs.has(value);
  }
  return false;
}

/**
 * Update a checkbox's display run without replacing its structural content. A cell-level
 * SDT owns w:tc, including its width and borders; flattening that to a run deletes a cell
 * from the table. Row/block controls likewise keep their containers and sibling content.
 *
 * The display run is the one showing a declared state glyph; only when no run does is the
 * first live run taken. Deleted and moved-from runs are history, never the display, and a
 * paragraph with no live run gets one. `inline` says where the control sits: a block-level
 * control with no content yet gets its run inside a new paragraph, because w:sdtContent at
 * block level cannot hold a bare w:r.
 */
export function checkboxContent(
  content: OoxmlElement | undefined,
  symbol: CheckboxSymbol,
  text: string,
  nextId: () => string,
  inline: boolean
): readonly OoxmlNode[] | null {
  const mint = (properties?: OoxmlNode): OoxmlNode =>
    mintCheckboxRun(nextId, symbol.hex, symbol.font, properties, text);
  if (!content || content.children.length === 0) {
    return inline ? [mint()] : [paragraphOf(nextId, [mint()])];
  }

  const hexes = new Set<string>();
  const glyphs = new Set<string>([text]);
  for (const state of [symbol.hex, ...symbol.states]) {
    const hex = normalizedHex(state);
    if (hex === null) continue;
    hexes.add(hex);
    const glyph = glyphOf(hex);
    if (glyph !== null) glyphs.add(glyph);
  }
  const isDisplayRun = (run: OoxmlElement): boolean => displaysState(run, hexes, glyphs);

  // Follow content containers only: never mistake a historical run in rPrChange or a
  // nested control's value for this checkbox's display. Bound recursion on imported XML.
  const isContainer = (node: OoxmlElement): boolean =>
    node.kind === 'contentControlContent' ||
    BLOCK_KINDS.has(node.kind) ||
    isInlineRunContainer(node) ||
    isWml(node, 'customXml');

  const rewrite = (
    node: OoxmlNode,
    depth: number,
    accept: (run: OoxmlElement) => boolean
  ): OoxmlNode | null => {
    if (node.kind === 'textValue' || depth > 32) return null;
    if (node.kind === 'run') {
      if (!accept(node)) return null;
      const run = mint(runPropertiesNodeOf(node));
      if (run.kind === 'textValue') return null;
      return { ...node, children: run.children } as OoxmlNode;
    }
    if (node.kind === 'revisionDelete' || node.kind === 'revisionMoveFrom') return null;
    if (!isContainer(node)) return null;
    for (let i = 0; i < node.children.length; i++) {
      const child = rewrite(node.children[i]!, depth + 1, accept);
      if (child) {
        const children = [...node.children];
        children[i] = child;
        return { ...node, children } as OoxmlNode;
      }
    }
    return null;
  };

  // No live run anywhere: the first paragraph gets one, placed before its closing markers so
  // a bookmark or comment that covered the empty paragraph covers the glyph too.
  const appendRun = (node: OoxmlNode, depth: number): OoxmlNode | null => {
    if (node.kind === 'textValue' || depth > 32) return null;
    if (node.kind === 'paragraph') {
      let at = node.children.length;
      while (at > 0 && TRAILING_MARKER_KINDS.has(node.children[at - 1]!.kind)) at--;
      const children: OoxmlNode[] = [...node.children];
      children.splice(at, 0, mint(paragraphMarkProperties(node, nextId)));
      return { ...node, children } as OoxmlNode;
    }
    if (!isContainer(node) || isInlineRunContainer(node)) return null;
    for (let i = 0; i < node.children.length; i++) {
      const child = appendRun(node.children[i]!, depth + 1);
      if (child) {
        const children = [...node.children];
        children[i] = child;
        return { ...node, children } as OoxmlNode;
      }
    }
    return null;
  };

  const updated =
    rewrite(content, 0, isDisplayRun) ?? rewrite(content, 0, () => true) ?? appendRun(content, 0);
  // Unsupported content must not be flattened as a fallback: that is the data-loss bug.
  return updated && updated.kind !== 'textValue' ? updated.children : null;
}
