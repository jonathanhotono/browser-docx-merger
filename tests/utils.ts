export async function blobToArrayBuffer(blob: any): Promise<ArrayBuffer> {
  if (!blob) throw new Error('No blob provided');
  
  // In Node.js, the Blob might not have arrayBuffer method
  if (typeof blob.arrayBuffer === 'function') {
    try {
      return await blob.arrayBuffer();
    } catch (error) {
      console.warn('blob.arrayBuffer() failed, trying stream method:', error);
    }
  }
  
  // Try using stream method for Node.js environments
  if (typeof blob.stream === 'function') {
    try {
      const stream = blob.stream();
      const reader = stream.getReader();
      const chunks: Uint8Array[] = [];
      
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
      }
      
      const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const combined = new Uint8Array(totalLength);
      let offset = 0;
      for (const chunk of chunks) {
        combined.set(chunk, offset);
        offset += chunk.length;
      }
      
      return combined.buffer.slice(0, totalLength);
    } catch (error) {
      console.warn('blob.stream() failed, trying Response fallback:', error);
    }
  }
  
  // Final fallback using Response
  try {
    return await new Response(blob).arrayBuffer();
  } catch (error) {
    throw new Error(`Failed to convert blob to ArrayBuffer: ${error}`);
  }
}
