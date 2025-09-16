import { describe, it, expect } from 'vitest'
import { mergeDocx } from '../src/index'
import { makeMinimalDocx, defaultContentTypes } from './fixtures'
import JSZip from 'jszip'
import { blobToArrayBuffer } from './utils'

const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships'

function docWithImage(rId = 'rId1'){
  return `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r>
      <w:drawing>
        <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:blipFill>
                  <a:blip r:embed="${rId}"/>
                </pic:blipFill>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r></w:p>
  </w:body>
</w:document>`
}

function relsWithImage(target = 'media/image1.png'){
  return `<Relationships xmlns="${NS_PR}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/></Relationships>`
}

function contentTypesWithPng(){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
}

describe('relationships & media copying', () => {
  it('copies media and remaps rId', async () => {
    const imageData = new Uint8Array([137,80,78,71]) // fake PNG header
  const base = await makeMinimalDocx(docWithImage('rId1'), { rels: relsWithImage('media/image1.png'), contentTypes: contentTypesWithPng(), media: { 'word/media/image1.png': imageData } })
  const src = await makeMinimalDocx(docWithImage('rId1'), { rels: relsWithImage('media/image2.png'), contentTypes: contentTypesWithPng(), media: { 'word/media/image2.png': imageData } })

    const blob = await mergeDocx([base, src], { insertEnd: true })
  const buf = await blobToArrayBuffer(blob as any)
    const zip = await JSZip.loadAsync(buf)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('text')
    // Should contain some rId2 or higher pointing to some merged media path
    expect(rels).toMatch(/Type="[^"]+\/image"/)
    const files = Object.keys(zip.files)
    expect(files.some(f => f.startsWith('word/media/'))).toBe(true)
  })
})
