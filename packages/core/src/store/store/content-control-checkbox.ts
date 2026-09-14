import { isInlineRunContainer } from '../package/ooxml-shared.ts';
import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import { mintCheckboxRun } from './content-control-run.ts';
import { runPropertiesNodeOf } from './tree-op-nodes.ts';

/**
 * Update a checkbox's display run without replacing its structural content. A cell-level
 * SDT owns w:tc, including its width and borders; flattening that to a run deletes a cell
 * from the table. Row/block controls likewise keep their containers and sibling content.
 */
export function checkboxContent(
  content: OoxmlElement | undefined,
  symbol: { readonly hex: string; readonly font: string },
  text: string,
  nextId: () => string
): readonly OoxmlNode[] | null {
  const mint = (properties?: OoxmlNode): OoxmlNode =>
    mintCheckboxRun(nextId, symbol.hex, symbol.font, properties, text);
  if (!content || content.children.length === 0) return [mint()];

  // Follow content containers only: never mistake a historical run in rPrChange or a
  // nested control's value for this checkbox's display. Bound recursion on imported XML.
  const update = (node: OoxmlNode, depth: number): OoxmlNode | null => {
    if (node.kind === 'textValue' || depth > 32) return null;
    if (node.kind === 'run') {
      const run = mint(runPropertiesNodeOf(node));
      if (run.kind === 'textValue') return null;
      return { ...node, children: run.children } as OoxmlNode;
    }
    if (
      node.kind !== 'contentControlContent' &&
      node.kind !== 'table' &&
      node.kind !== 'tableRow' &&
      node.kind !== 'tableCell' &&
      node.kind !== 'paragraph' &&
      !isInlineRunContainer(node)
    )
      return null;
    for (let i = 0; i < node.children.length; i++) {
      const child = update(node.children[i]!, depth + 1);
      if (child) {
        const children = [...node.children];
        children[i] = child;
        return { ...node, children } as OoxmlNode;
      }
    }
    if (node.kind === 'paragraph') {
      return { ...node, children: [...node.children, mint()] } as OoxmlNode;
    }
    return null;
  };
  const updated = update(content, 0);
  // Unsupported content must not be flattened as a fallback: that is the data-loss bug.
  return updated && updated.kind !== 'textValue' ? updated.children : null;
}
