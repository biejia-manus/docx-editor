import { expect, test } from 'bun:test';
import {
  contentControlsIn,
  contentControlPropertiesOf,
  findNode,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlPart,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const meta = { name: '/word/document.xml', contentType: 'application/xml' };
const run =
  '<w:r w:rsidRPr="00D55315"><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/><w:b/></w:rPr><w:t>☐</w:t></w:r>';
const paragraph = (content: string) =>
  '<w:p w14:paraId="4E8CAE0A" w14:textId="0E84969E" w:rsidR="00FA4ED4"><w:pPr><w:pStyle w:val="BOX"/></w:pPr><w:bookmarkStart w:id="0" w:name="Check"/>' +
  content +
  '<w:bookmarkEnd w:id="0"/></w:p>';
const cell = (content: string) =>
  '<w:tc><w:tcPr><w:tcW w:w="434" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/></w:tcBorders></w:tcPr>' +
  content +
  '</w:tc>';
const otherCell =
  '<w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Keep this label</w:t></w:r></w:p></w:tc>';
const row = (content: string) =>
  '<w:tr w:rsidR="00000001"><w:trPr><w:cantSplit/></w:trPr>' + content + otherCell + '</w:tr>';
const table = (content: string) =>
  '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="434"/><w:gridCol w:w="8000"/></w:tblGrid>' +
  content +
  '</w:tbl>';
const control = (content: string) =>
  '<w:sdt><w:sdtPr><w:id w:val="780542881"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtEndPr/><w:sdtContent>' +
  content +
  '</w:sdtContent></w:sdt>';

function load(body: string): OoxmlPart {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`,
    meta
  );
  if (!parsed.ok) throw new Error(parsed.reason);
  return parsed.part;
}

for (const [placement, body] of [
  ['inline', paragraph(control(run))],
  ['paragraph', control(paragraph(run))],
  ['cell', table(row(control(cell(paragraph(run)))))],
  ['row', table(control(row(cell(paragraph(run)))))],
  ['table', control(table(row(cell(paragraph(run)))))],
] as const) {
  for (const input of ['string', 'typed'] as const) {
    test(`checkbox ${input} toggles retain the complete ${placement} structure across serialization`, () => {
      const original = load(body);
      let part = original;
      let expected = serializeOoxmlPart(original);
      for (const checked of [true, false, true]) {
        const node = contentControlsIn(part.root)[0]!.node;
        const result = applyTreeOp(part, {
          op: 'setContentControlValue',
          controlId: node.id,
          value: input === 'string' ? String(checked) : { kind: 'checkbox', checked },
        });
        expect(result.ok).toBe(true);
        if (!result.ok) throw new Error(result.reason);
        expect(findNode(result.part, node.id)).toBeDefined();
        const xml = serializeOoxmlPart(result.part);
        expected = expected
          .replace(
            /<w14:checked w14:val="[01]"\/>/,
            `<w14:checked w14:val="${checked ? '1' : '0'}"/>`
          )
          .replace(
            /<w:t>☐<\/w:t>|<w:sym w:char="261[02]" w:font="MS Gothic"\/>/,
            `<w:sym w:char="${checked ? '2612' : '2610'}" w:font="MS Gothic"/>`
          );
        // Exact XML comparison catches lost cell/row properties, bookmarks, run attributes,
        // paragraph IDs, sibling cells and content—not just the final checked flag.
        expect(xml).toBe(expected);
        const reopened = readOoxmlPart(xml, meta);
        if (!reopened.ok) throw new Error(reopened.reason);
        part = reopened.part;
        expect(
          contentControlPropertiesOf(contentControlsIn(part.root)[0]!.node).checkbox?.checked
        ).toBe(checked);
      }
      expect(serializeOoxmlPart(original)).toBe(serializeOoxmlPart(load(body)));
    });
  }
}

test('an empty checkbox paragraph keeps its cell and properties when its first run is inserted', () => {
  const part = load(table(row(control(cell(paragraph(''))))));
  const node = contentControlsIn(part.root)[0]!.node;
  const result = applyTreeOp(part, {
    op: 'setContentControlValue',
    controlId: node.id,
    value: 'true',
  });
  if (!result.ok) throw new Error(result.reason);
  const before = serializeOoxmlPart(part);
  const after = serializeOoxmlPart(result.part);
  expect(after.match(/<w:tc>/g)?.length).toBe(before.match(/<w:tc>/g)?.length);
  expect(after).toContain('<w:pStyle w:val="BOX"/>');
  expect(after).toContain('<w:bookmarkStart w:id="0" w:name="Check"/>');
  expect(after).toContain('<w:sym w:char="2612" w:font="MS Gothic"/>');
  expect(after).toContain('Keep this label');
});

for (const input of ['string', 'typed'] as const) {
  test(`checkbox ${input} writes refuse opaque content without flattening it`, () => {
    const part = load(
      control('<x:content xmlns:x="urn:opaque"><x:value>keep</x:value></x:content>')
    );
    const before = serializeOoxmlPart(part);
    const node = contentControlsIn(part.root)[0]!.node;
    const result = applyTreeOp(part, {
      op: 'setContentControlValue',
      controlId: node.id,
      value: input === 'string' ? 'true' : { kind: 'checkbox', checked: true },
    });
    expect(result).toEqual({ ok: false, reason: 'unsupported' });
    expect(serializeOoxmlPart(part)).toBe(before);
  });
}

const mintedRun =
  '<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr><w:sym w:char="2612" w:font="MS Gothic"/></w:r>';
const deletedRun = (content: string) =>
  '<w:del w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">' + content + '</w:del>';

for (const input of ['string', 'typed'] as const) {
  const value = input === 'string' ? 'true' : ({ kind: 'checkbox', checked: true } as const);
  const toggle = (part: OoxmlPart): string => {
    const node = contentControlsIn(part.root)[0]!.node;
    const result = applyTreeOp(part, { op: 'setContentControlValue', controlId: node.id, value });
    if (!result.ok) throw new Error(result.reason);
    return serializeOoxmlPart(result.part);
  };

  test(`an empty block-level checkbox ${input} write lands its run inside a paragraph`, () => {
    expect(toggle(load(control('')))).toContain(
      `<w:sdtContent><w:p>${mintedRun}</w:p></w:sdtContent>`
    );
  });

  test(`an empty inline checkbox ${input} write lands a bare run`, () => {
    expect(toggle(load(paragraph(control(''))))).toContain(
      `<w:sdtContent>${mintedRun}</w:sdtContent>`
    );
  });

  test(`checkbox ${input} writes skip a tracked deletion and update the live run`, () => {
    const deleted = deletedRun('<w:r><w:rPr><w:b/></w:rPr><w:delText>☒</w:delText></w:r>');
    const xml = toggle(load(paragraph(control(deleted + run))));
    expect(xml).toContain('<w:delText>☒</w:delText>');
    expect(xml).toContain(
      '</w:del><w:r w:rsidRPr="00D55315"><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/><w:b/></w:rPr><w:sym w:char="2612" w:font="MS Gothic"/></w:r>'
    );
    expect(xml.match(/<w:sym /g)).toHaveLength(1);
  });

  test(`checkbox ${input} writes add a live run when the paragraph holds only a tracked deletion`, () => {
    const deleted = deletedRun('<w:r><w:delText>☒</w:delText></w:r>');
    const xml = toggle(load(control(paragraph(deleted))));
    expect(xml).toContain('<w:delText>☒</w:delText>');
    // The new run lands inside the bookmark that covered the empty paragraph.
    expect(xml).toContain(`</w:del>${mintedRun}<w:bookmarkEnd w:id="0"/></w:p>`);
    expect(xml.match(/<w:sym /g)).toHaveLength(1);
  });

  test(`checkbox ${input} writes update the glyph run, not a label that precedes it`, () => {
    const label = '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Option A </w:t></w:r>';
    const xml = toggle(load(table(row(control(cell(paragraph(label + run)))))));
    expect(xml).toContain('<w:t xml:space="preserve">Option A </w:t>');
    expect(xml).toContain(
      'Option A </w:t></w:r><w:r w:rsidRPr="00D55315"><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/><w:b/></w:rPr><w:sym w:char="2612" w:font="MS Gothic"/></w:r>'
    );
    expect(xml.match(/<w:sym /g)).toHaveLength(1);
    expect(xml).not.toContain('<w:t>☐</w:t>');
  });

  test(`checkbox ${input} writes recognise a previously written w:sym as the display run`, () => {
    const label = '<w:r><w:t xml:space="preserve">Option A </w:t></w:r>';
    const symbolRun = '<w:r><w:sym w:font="MS Gothic" w:char="2612"/></w:r>';
    const xml = toggle(load(paragraph(control(label + symbolRun))));
    expect(xml).toContain('<w:t xml:space="preserve">Option A </w:t>');
    expect(xml.match(/<w:sym /g)).toHaveLength(1);
  });

  test(`checkbox ${input} writes reach a glyph run inside block-level w:customXml`, () => {
    const wrapped =
      '<w:customXml w:uri="urn:x" w:element="section">' + paragraph(run) + '</w:customXml>';
    const xml = toggle(load(control(wrapped)));
    expect(xml).toContain('<w:customXml w:element="section" w:uri="urn:x">');
    expect(xml).toContain('<w:sym w:char="2612" w:font="MS Gothic"/>');
    expect(xml).not.toContain('<w:t>☐</w:t>');
  });

  test(`checkbox ${input} writes into an empty paragraph inherit the paragraph mark formatting`, () => {
    const marked =
      '<w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="a" w:date="2024-01-01T00:00:00Z"/><w:b/></w:rPr></w:pPr></w:p>';
    const xml = toggle(load(control(marked)));
    expect(xml).toContain(
      '</w:pPr><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/><w:b/></w:rPr><w:sym w:char="2612" w:font="MS Gothic"/></w:r></w:p>'
    );
  });

  test(`checkbox ${input} writes refuse a nested control rather than writing into it`, () => {
    const nested =
      '<w:sdt><w:sdtPr><w:id w:val="5"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>';
    const part = load(paragraph(control(nested)));
    const before = serializeOoxmlPart(part);
    const node = contentControlsIn(part.root)[0]!.node;
    const result = applyTreeOp(part, { op: 'setContentControlValue', controlId: node.id, value });
    expect(result).toEqual({ ok: false, reason: 'unsupported' });
    expect(serializeOoxmlPart(part)).toBe(before);
  });
}
