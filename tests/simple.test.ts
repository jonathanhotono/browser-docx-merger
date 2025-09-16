import { describe, it, expect } from 'vitest'
import { mergeDocx } from '../src/index'

describe('library smoke tests', () => {
  it('throws error for insufficient inputs', async () => {
    const mockBlob = new Blob(['test'], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    await expect(mergeDocx([mockBlob], { insertEnd: true })).rejects.toThrow('Need at least two DOCX sources')
  })

  it('throws error for no insertion options', async () => {
    const mockBlob1 = new Blob(['test1'], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    const mockBlob2 = new Blob(['test2'], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    await expect(mergeDocx([mockBlob1, mockBlob2], {})).rejects.toThrow('Provide a pattern or set insertStart/insertEnd')
  })
})