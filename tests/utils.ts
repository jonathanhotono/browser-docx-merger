export async function blobToArrayBuffer(blob: any): Promise<ArrayBuffer> {
  if (!blob) throw new Error('No blob provided');
  if (typeof blob.arrayBuffer === 'function') return blob.arrayBuffer();
  return await new Response(blob).arrayBuffer();
}
