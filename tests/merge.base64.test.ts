import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { mergeDocx } from '../src/index';
import { readFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

// Handle __dirname in ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const docsDir = join(__dirname, 'docs');
const doc1Path = join(docsDir, '1.docx');
const doc2Path = join(docsDir, '2.docx');

function loadDocBuffer(path: string): Uint8Array {
  const data = readFileSync(path);
  return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
}

function bufferToBase64(buffer: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(buffer).toString('base64');
  }
  // Fallback for browser environments
  let binary = '';
  for (let i = 0; i < buffer.length; i++) {
    binary += String.fromCharCode(buffer[i]);
  }
  return btoa(binary);
}

describe('mergeDocx base64 integration', () => {
  it('merges two base64 DOCX sources (append at end)', async () => {
    // Load real DOCX files and convert to base64
    const doc1Buffer = loadDocBuffer(doc1Path);
    const doc2Buffer = loadDocBuffer(doc2Path);
    const doc1Base64 = bufferToBase64(doc1Buffer);
    const doc2Base64 = bufferToBase64(doc2Buffer);
    
    console.log(`Base64 lengths: doc1=${doc1Base64.length}, doc2=${doc2Base64.length}`);
    
    try {
      const blob = await mergeDocx([doc1Base64, doc2Base64], {
        insertEnd: true,
        mergeNumbering: true,
        mergeStyles: true,
        mergeFootnotes: true,
      });

      expect(blob).toBeInstanceOf(Blob);
      expect(blob.type).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      expect(blob.size).toBeGreaterThan(10000); // Reasonable size for merged document
      
      console.log(`Base64 merge successful: ${blob.size} bytes`);
    } catch (error) {
      console.error('Base64 merge failed:', error);
      throw error;
    }
  });

  it('accepts mixed input types (base64 + Uint8Array) and still merges', async () => {
    const doc1Buffer = loadDocBuffer(doc1Path);
    const doc2Buffer = loadDocBuffer(doc2Path);
    const doc1Base64 = bufferToBase64(doc1Buffer);
    
    try {
      const blob = await mergeDocx([doc1Base64, doc2Buffer], { insertEnd: true });
      expect(blob).toBeInstanceOf(Blob);
      expect(blob.size).toBeGreaterThan(10000);
      
      console.log(`Mixed input merge successful: ${blob.size} bytes`);
    } catch (error) {
      console.error('Mixed input merge failed:', error);
      throw error;
    }
  });
});
