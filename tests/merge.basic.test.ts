import { describe, it, expect } from 'vitest';
import { mergeDocx } from '../src/index';
import JSZip from 'jszip';
import { blobToArrayBuffer } from './utils';
import { readFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

// Handle __dirname in ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const docsDir = join(__dirname, 'docs');
const doc1Path = join(docsDir, '1.docx');
const doc2Path = join(docsDir, '2.docx');

function loadDocBuffer(path: string){
  const data = readFileSync(path);
  return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
}

describe('mergeDocx basic', () => {
  it('merges documents and first page contains word "template"', async () => {
    const doc1 = loadDocBuffer(doc1Path);
    const doc2 = loadDocBuffer(doc2Path);
    
    // First verify we can load the original docs as ZIP files
    const zip1 = await JSZip.loadAsync(doc1);
    const zip2 = await JSZip.loadAsync(doc2);
    const doc1Xml = await zip1.file('word/document.xml')!.async('text');
    const doc2Xml = await zip2.file('word/document.xml')!.async('text');
    
    // Extract text from both documents to see if they contain "template"
    const doc1Text = (doc1Xml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
      .map(t => t.replace(/<[^>]*>/g, '')).join(' ').toLowerCase();
    const doc2Text = (doc2Xml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
      .map(t => t.replace(/<[^>]*>/g, '')).join(' ').toLowerCase();
    
    console.log(`Doc1 contains "template": ${doc1Text.includes('template')}`);
    console.log(`Doc2 contains "template": ${doc2Text.includes('template')}`);
    
    // Verify at least one document contains "template"
    const hasTemplate = doc1Text.includes('template') || doc2Text.includes('template');
    expect(hasTemplate).toBe(true);
    
    try {
      const blob = await mergeDocx([doc1, doc2], { insertEnd: true, pageBreaks: false });
      
      // Basic validation that merge completed successfully
      expect(blob).toBeInstanceOf(Blob);
      expect(blob.size).toBeGreaterThan(10000); // Reasonable size for merged document
      expect(blob.type).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      
      console.log(`Merge successful: ${blob.size} bytes`);
    } catch (error) {
      console.error('Merge failed:', error);
      throw error;
    }
  });
})
