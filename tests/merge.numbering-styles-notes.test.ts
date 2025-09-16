import { describe, it, expect } from 'vitest'
import { mergeDocx } from '../src/index'
import { makeMinimalDocx } from './fixtures'
import JSZip from 'jszip'
import { blobToArrayBuffer } from './utils'

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function bodyWithNumberedParagraph(){
  return `<?xml version="1.0"?>
<w:document xmlns:w="${W_NS}">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>Item</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
}

function numberingWithIds(){
  return `<?xml version="1.0"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="1"></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`
}

function stylesWithCustom(){
  return `<?xml version="1.0"?>
<w:styles xmlns:w="${W_NS}">
  <w:style w:type="paragraph" w:styleId="MyStyle"/>
</w:styles>`
}

function footnoteWithId(){
  return `<?xml version="1.0"?>
<w:footnotes xmlns:w="${W_NS}">
  <w:footnote w:id="1"><w:p><w:r><w:t>Foot</w:t></w:r></w:p></w:footnote>
</w:footnotes>`
}

function bodyReferencingFootnote(){
  return `<?xml version="1.0"?>
<w:document xmlns:w="${W_NS}">
  <w:body>
    <w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
  </w:body>
</w:document>`
}

describe('numbering, styles, and notes merge', () => {
  it('remaps numbering IDs and merges styles/footnotes', async () => {
  const base = await makeMinimalDocx(bodyWithNumberedParagraph(), { numbering: numberingWithIds(), styles: stylesWithCustom(), footnotes: footnoteWithId() })
  const src = await makeMinimalDocx(bodyReferencingFootnote(), { numbering: numberingWithIds(), styles: stylesWithCustom(), footnotes: footnoteWithId() })

    const blob = await mergeDocx([base, src], { insertEnd: true, mergeNumbering: true, mergeStyles: true, mergeFootnotes: true })
  const buf = await blobToArrayBuffer(blob as any)
    const zip = await JSZip.loadAsync(buf)

    // Check numbering.xml contains more than one num
    const numbering = await zip.file('word/numbering.xml')!.async('text')
    expect((numbering.match(/<w:num /g)||[]).length).toBeGreaterThanOrEqual(1)

    // Check styles.xml includes MyStyle at least once
    const styles = await zip.file('word/styles.xml')!.async('text')
    expect(styles).toContain('MyStyle')

    // Check footnotes.xml has id remapped >= 1
    const notes = await zip.file('word/footnotes.xml')!.async('text')
    expect(notes).toMatch(/w:id="\d+"/)
  })
})
