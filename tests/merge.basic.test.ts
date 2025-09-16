import { describe, it, expect } from 'vitest'
import { mergeDocx } from '../src/index'
import { makeMinimalDocx } from './fixtures'
import { blobToArrayBuffer } from './utils'

function p(text: string){
  return `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`
}

describe('mergeDocx basic', () => {
  it('throws if less than two sources', async () => {
    await expect(mergeDocx([await makeMinimalDocx(p('A'))], { insertEnd: true })).rejects.toThrow()
  })

  it('merges at end', async () => {
    const blob = await mergeDocx([
      await makeMinimalDocx(p('Base')),
      await makeMinimalDocx(p('Second')),
      await makeMinimalDocx(p('Third')),
    ], { insertEnd: true, pageBreaks: false })

  const buf = await blobToArrayBuffer(blob as any)
    // quick sanity: ensure document.xml exists and includes expected text
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(buf)
    const doc = await zip.file('word/document.xml')!.async('text')
    expect(doc).toContain('Base')
    expect(doc).toContain('Second')
    expect(doc).toContain('Third')
  })

  it('inserts before pattern', async () => {
    const blob = await mergeDocx([
      await makeMinimalDocx(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>ANCHOR</w:t></w:r></w:p></w:body></w:document>`),
      await makeMinimalDocx(p('Insert1')),
      await makeMinimalDocx(p('Insert2')),
    ], { pattern: 'ANCHOR', pageBreaks: false })

  const buf = await blobToArrayBuffer(blob as any)
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(buf)
    const doc = await zip.file('word/document.xml')!.async('text')
    // Both inserts should be present
    expect(doc).toContain('Insert1')
    expect(doc).toContain('Insert2')
  })
})
