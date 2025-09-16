import JSZip from 'jszip';

export async function makeMinimalDocx(bodyXml: string, extras?: { numbering?: string, styles?: string, footnotes?: string, endnotes?: string, rels?: string, contentTypes?: string, media?: Record<string, Uint8Array> }){
  const zip = new JSZip();
  
  // Add required .docx structure with proper XML headers
  zip.file('[Content_Types].xml', extras?.contentTypes ?? defaultContentTypes());
  zip.file('_rels/.rels', defaultMainRels());
  zip.file('word/document.xml', ensureXmlHeader(bodyXml));
  zip.file('word/_rels/document.xml.rels', extras?.rels ?? defaultDocRels());
  
  // Optional parts
  if(extras?.numbering) zip.file('word/numbering.xml', ensureXmlHeader(extras.numbering));
  if(extras?.styles) zip.file('word/styles.xml', ensureXmlHeader(extras.styles));  
  if(extras?.footnotes) zip.file('word/footnotes.xml', ensureXmlHeader(extras.footnotes));
  if(extras?.endnotes) zip.file('word/endnotes.xml', ensureXmlHeader(extras.endnotes));

  if(extras?.media){
    for(const [path, data] of Object.entries(extras.media)){
      zip.file(path, data);
    }
  }

  const buf = await zip.generateAsync({type:'arraybuffer', compression: 'DEFLATE'});
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

export function defaultContentTypes(){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
}

function defaultMainRels(){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
}

function defaultDocRels(){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
}

function ensureXmlHeader(xml: string): string {
  if(xml.trim().startsWith('<?xml')) return xml;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
}
