/*
  browser-doc-merger (framework-style)
  Public API:
    - mergeDocx(files: (ArrayBuffer|Uint8Array|Blob)[], options: MergeOptions): Promise<Blob>
    - mergeDocxFromFiles(files: File[], options: MergeOptions): Promise<Blob>
  UMD global: DocxMerger
*/

import JSZip from 'jszip';

export type MergeOptions = {
  pattern?: string | null;
  insertStart?: boolean;
  insertEnd?: boolean;
  pageBreaks?: boolean;
  mergeNumbering?: boolean;
  mergeStyles?: boolean;
  mergeFootnotes?: boolean;
  onLog?: (msg: string, level?: 'info'|'ok'|'warn'|'err') => void;
};

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships'
} as const;

const P = {
  document: 'word/document.xml',
  docRels: 'word/_rels/document.xml.rels',
  numbering: 'word/numbering.xml',
  styles: 'word/styles.xml',
  footnotes: 'word/footnotes.xml',
  endnotes: 'word/endnotes.xml',
  fontTable: 'word/fontTable.xml',
  theme: 'word/theme/theme1.xml',
  webSettings: 'word/webSettings.xml',
  settings: 'word/settings.xml',
  ct: '[Content_Types].xml'
} as const;

export async function mergeDocx(inputBuffers: (ArrayBuffer|Uint8Array|Blob)[], options: MergeOptions = {}): Promise<Blob>{
  const opt = withDefaults(options);
  const log = (m: string, lvl: 'info'|'ok'|'warn'|'err'='info')=>opt.onLog?.(m,lvl);

  if(inputBuffers.length < 2) throw new Error('Need at least two DOCX sources');
  if(!opt.pattern && !opt.insertStart && !opt.insertEnd){
    throw new Error('Provide a pattern or set insertStart/insertEnd');
  }

  log?.(`Reading ${inputBuffers.length} DOCX files ...`);
  const buffers = await Promise.all(inputBuffers.map(src => toBytes(src)));
  const zips = await Promise.all(buffers.map((buf: Uint8Array) => JSZip.loadAsync(buf)));

  const baseZip = zips[0];
  const baseDoc = await readXml(baseZip, P.document);
  const baseRels = await readXmlOrCreate(baseZip, P.docRels, relsSkeleton());
  const baseCT = await readXml(baseZip, P.ct);

  const allocRid = makeRidAllocator(baseRels);
  const ensureCT = makeContentTypesEnsurer(baseCT);

  const baseBody = bodyEl(baseDoc);
  let baseSectPr = lastSectPr(baseBody);
  if(baseSectPr) baseBody.removeChild(baseSectPr);

  let totalReplacements = 0 + (opt.insertStart?1:0) + (opt.insertEnd?1:0);
  if(opt.insertStart){
    for(let i=1;i<zips.length;i++){
      await appendDocInto(baseZip, zips[i], baseDoc, baseRels, allocRid, ensureCT, opt, /*atEnd*/ false);
    }
  }

  if(opt.pattern){
    const matches = findParagraphsContainingText(baseDoc, opt.pattern);
    if(matches.length===0){
      log?.('Pattern not found in base document', 'warn');
    } else {
      const anchor = matches[0];
      for(let i=1;i<zips.length;i++){
        await insertDocBefore(baseZip, zips[i], baseDoc, baseRels, allocRid, ensureCT, anchor, opt);
      }
      totalReplacements += 1;
    }
  }

  if(opt.insertEnd){
    for(let i=1;i<zips.length;i++){
      await appendDocInto(baseZip, zips[i], baseDoc, baseRels, allocRid, ensureCT, opt, /*atEnd*/ true);
    }
  }

  if(totalReplacements===0) throw new Error('No insertion was performed (nothing to do).');

  if(baseSectPr) baseBody.appendChild(baseSectPr);

  await writeXml(baseZip, P.document, baseDoc);
  await writeXml(baseZip, P.docRels, baseRels);
  await writeXml(baseZip, P.ct, baseCT);

  log?.('Packaging merged .docx ...');
  const bytes = await baseZip.generateAsync({type:'arraybuffer'});
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  log?.('Done.');
  return blob;
}

export async function mergeDocxFromFiles(files: File[], options: MergeOptions = {}): Promise<Blob>{
  return mergeDocx(files, options);
}

// ====== High-level helpers (ported) ======
async function insertDocBefore(baseZip: JSZip, srcZip: JSZip, baseDoc: Document, baseRels: Document, allocRid: ()=>string, ensureCT: (ext:string)=>void, anchorP: Element, opt: RequiredMergeOptions){
  const { importedNodes } = await importBodyNodes(baseZip, srcZip, baseDoc, baseRels, allocRid, ensureCT, opt);
  const parent = anchorP.parentNode as Element; // w:body
  for(const n of importedNodes){ parent.insertBefore(n, anchorP); }
}

async function appendDocInto(baseZip: JSZip, srcZip: JSZip, baseDoc: Document, baseRels: Document, allocRid: ()=>string, ensureCT: (ext:string)=>void, opt: RequiredMergeOptions, atEnd: boolean){
  const baseBody = bodyEl(baseDoc);
  const { importedNodes } = await importBodyNodes(baseZip, srcZip, baseDoc, baseRels, allocRid, ensureCT, opt);
  if(opt.pageBreaks && baseBody.lastChild){ baseBody.appendChild(makePageBreak(baseDoc)); }
  for(const n of importedNodes){ baseBody.appendChild(n); }
}

async function importBodyNodes(baseZip: JSZip, srcZip: JSZip, baseDoc: Document, baseRels: Document, allocRid: ()=>string, ensureCT: (ext:string)=>void, opt: RequiredMergeOptions){
  const srcDoc = await readXml(srcZip, P.document);
  const srcBody = bodyEl(srcDoc);
  const srcSect = lastSectPr(srcBody); if(srcSect) srcBody.removeChild(srcSect);
  const srcRels = await readXmlOrCreate(srcZip, P.docRels, relsSkeleton());

  let numIdMap = new Map<number, number>();
  let absIdMap = new Map<number, number>();
  if(opt.mergeNumbering){
    ({ numIdMap, absIdMap } = await mergeNumberingParts(baseZip, srcZip, baseDoc));
    remapNumberingInBody(srcBody, numIdMap);
  }

  if(opt.mergeStyles){ 
    await mergeStylesPart(baseZip, srcZip); 
    await ensureStyleRelationships(baseRels, allocRid);
  }

  if(opt.mergeFootnotes){
    await mergeNotes(baseZip, srcZip, 'footnotes');
    await mergeNotes(baseZip, srcZip, 'endnotes');
    remapNoteRefsInBody(srcBody, await getNoteIdRemap(baseZip, srcZip, 'footnotes'));
    remapNoteRefsInBody(srcBody, await getNoteIdRemap(baseZip, srcZip, 'endnotes'));
  }

  const importedNodes: Node[] = [];
  for(const node of Array.from(srcBody.childNodes)){
    if(node.nodeType!==1) continue;
    const imported = baseDoc.importNode(node, true) as Element;
    const all = [imported, ...descendants(imported)];
    for(const el of all){
      if(el.nodeType!==1) continue;
      for(const attr of Array.from((el as Element).attributes||[])){
        const needs = (attr.localName==='id' && attr.namespaceURI===NS.r) || attr.name==='r:id';
        if(!needs) continue;
        const oldRid = attr.value;
        const newRid = await remapRelationship(oldRid, srcRels, baseRels, srcZip, baseZip, allocRid, ensureCT);
        if(newRid){ (el as Element).setAttributeNS(NS.r,'r:id',newRid); } else { (el as Element).removeAttributeNS(NS.r,'id'); }
      }
    }
    importedNodes.push(imported);
  }

  return { importedNodes };
}

// ====== Numbering ======
async function mergeNumberingParts(baseZip: JSZip, srcZip: JSZip, baseDoc: Document){
  const baseNum = await readXmlOrCreate(baseZip, P.numbering, numberingSkeleton());
  const srcNum = await readXmlOrCreate(srcZip, P.numbering, numberingSkeleton());

  const baseRoot = baseNum.documentElement; // w:numbering
  const srcRoot = srcNum.documentElement;

  const baseAbsIds = new Set(Array.from(baseRoot.getElementsByTagNameNS(NS.w,'abstractNum')).map(x=>+(x.getAttributeNS(NS.w,'abstractNumId')||x.getAttribute('w:abstractNumId')||x.getAttribute('w:abstractNumId')||'0')));
  const baseNumIds = new Set(Array.from(baseRoot.getElementsByTagNameNS(NS.w,'num')).map(x=>+(x.getAttributeNS(NS.w,'numId')||x.getAttribute('w:numId')||x.getAttribute('w:numId')||'0')));
  let _absMax = Math.max(0, ...Array.from(baseAbsIds.values()));
  let _numMax = Math.max(0, ...Array.from(baseNumIds.values()));
  const nextAbsId = () => (++_absMax);
  const nextNumId = () => (++_numMax);

  const absIdMap = new Map<number, number>();
  const numIdMap = new Map<number, number>();

  for(const abs of Array.from(srcRoot.getElementsByTagNameNS(NS.w,'abstractNum'))){
    const oldId = +(abs.getAttributeNS(NS.w,'abstractNumId') || abs.getAttribute('w:abstractNumId') || abs.getAttribute('abstractNumId') || '0');
    const newId = nextAbsId();
    absIdMap.set(oldId, newId);
    abs.setAttributeNS(NS.w,'w:abstractNumId', String(newId));
    baseRoot.appendChild(baseNum.importNode(abs, true));
  }

  for(const num of Array.from(srcRoot.getElementsByTagNameNS(NS.w,'num'))){
    const oldNumId = +(num.getAttributeNS(NS.w,'numId') || num.getAttribute('w:numId') || num.getAttribute('numId') || '0');
    const newNumId = nextNumId();
    numIdMap.set(oldNumId, newNumId);
    num.setAttributeNS(NS.w,'w:numId', String(newNumId));
    const a = num.getElementsByTagNameNS(NS.w,'abstractNumId')[0] as Element | undefined;
    if(a){ const oldAbs = +(a.getAttributeNS(NS.w,'val') || a.getAttribute('w:val') || a.getAttribute('val') || '0'); const newAbs = absIdMap.get(oldAbs) ?? oldAbs; a.setAttributeNS(NS.w,'w:val', String(newAbs)); }
    baseRoot.appendChild(baseNum.importNode(num, true));
  }

  await writeXml(baseZip, P.numbering, baseNum);
  return { numIdMap, absIdMap };
}

function remapNumberingInBody(body: Element, numIdMap: Map<number,number>){
  const ps = body.getElementsByTagNameNS(NS.w,'p');
  for(const p of Array.from(ps)){
    const numPr = p.getElementsByTagNameNS(NS.w,'numPr')[0] as Element | undefined;
    if(!numPr) continue;
    const numId = numPr.getElementsByTagNameNS(NS.w,'numId')[0] as Element | undefined;
    if(!numId) continue;
    const v = +(numId.getAttributeNS(NS.w,'val') || numId.getAttribute('w:val') || numId.getAttribute('val') || 'NaN');
    if(Number.isFinite(v) && numIdMap.has(v)){
      numId.setAttributeNS(NS.w,'w:val', String(numIdMap.get(v)));
    }
  }
}

// ====== Styles ======
async function mergeStylesPart(baseZip: JSZip, srcZip: JSZip){
  // Merge main styles.xml
  const baseStyles = await readXmlOrCreate(baseZip, P.styles, stylesSkeleton());
  const srcStyles = await readXmlOrCreate(srcZip, P.styles, stylesSkeleton());
  const existing = new Set(Array.from(baseStyles.getElementsByTagNameNS(NS.w,'style')).map(s=>s.getAttributeNS(NS.w,'styleId')||s.getAttribute('w:styleId')||s.getAttribute('styleId')));
  for(const s of Array.from(srcStyles.getElementsByTagNameNS(NS.w,'style'))){
    const id = s.getAttributeNS(NS.w,'styleId')||s.getAttribute('w:styleId')||s.getAttribute('styleId');
    if(id && !existing.has(id)){
      baseStyles.documentElement.appendChild(baseStyles.importNode(s,true));
      existing.add(id);
    }
  }
  await writeXml(baseZip, P.styles, baseStyles);

  // Merge fontTable.xml if it exists in source
  await mergeOptionalStylePart(baseZip, srcZip, P.fontTable, fontTableSkeleton(), 'font');
  
  // Merge theme1.xml if it exists in source  
  await mergeOptionalStylePart(baseZip, srcZip, P.theme, themeSkeleton(), null);
  
  // Merge webSettings.xml if it exists in source
  await mergeOptionalStylePart(baseZip, srcZip, P.webSettings, webSettingsSkeleton(), null);
  
  // Merge settings.xml if it exists in source (only style-related parts)
  await mergeSettingsPart(baseZip, srcZip);
}

async function mergeOptionalStylePart(baseZip: JSZip, srcZip: JSZip, partPath: string, skeleton: string, mergeElementName: string | null) {
  const srcFile = srcZip.file(partPath);
  if (!srcFile) return; // Source doesn't have this part, skip
  
  if (mergeElementName) {
    // Merge specific elements (like fonts)
    const basePart = await readXmlOrCreate(baseZip, partPath, skeleton);
    const srcPart = await readXml(srcZip, partPath);
    const existing = new Set(Array.from(basePart.getElementsByTagNameNS(NS.w, mergeElementName)).map(el => 
      el.getAttributeNS(NS.w, 'name') || el.getAttribute('w:name') || el.getAttribute('name') || ''
    ));
    
    for (const el of Array.from(srcPart.getElementsByTagNameNS(NS.w, mergeElementName))) {
      const name = el.getAttributeNS(NS.w, 'name') || el.getAttribute('w:name') || el.getAttribute('name');
      if (name && !existing.has(name)) {
        basePart.documentElement.appendChild(basePart.importNode(el, true));
        existing.add(name);
      }
    }
    await writeXml(baseZip, partPath, basePart);
  } else {
    // Copy entire file if base doesn't have it
    if (!baseZip.file(partPath)) {
      const content = await srcFile.async('uint8array');
      baseZip.file(partPath, content);
    }
  }
}

async function mergeSettingsPart(baseZip: JSZip, srcZip: JSZip) {
  const srcSettings = srcZip.file(P.settings);
  if (!srcSettings) return;
  
  const baseSettings = await readXmlOrCreate(baseZip, P.settings, settingsSkeleton());
  const srcSettingsDoc = await readXml(srcZip, P.settings);
  
  // Merge specific style-related settings
  const styleElements = ['defaultTabStop', 'characterSpacingControl', 'printTwoOnOne', 'printColorBlackWhite', 'doNotPromptForConvert', 'mwSmallCaps'];
  
  for (const elementName of styleElements) {
    const srcElement = srcSettingsDoc.getElementsByTagNameNS(NS.w, elementName)[0];
    if (srcElement && !baseSettings.getElementsByTagNameNS(NS.w, elementName)[0]) {
      baseSettings.documentElement.appendChild(baseSettings.importNode(srcElement, true));
    }
  }
  
  await writeXml(baseZip, P.settings, baseSettings);
}

async function ensureStyleRelationships(baseRels: Document, allocRid: () => string) {
  const styleParts = [
    { path: P.styles, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' },
    { path: P.fontTable, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable' },
    { path: P.theme, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' },
    { path: P.webSettings, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings' },
    { path: P.settings, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings' }
  ];
  
  for (const part of styleParts) {
    const target = relativeTo(P.document, part.path);
    if (!findRelationshipByTypeAndTarget(baseRels, part.type, target)) {
      addRelationship(baseRels, allocRid(), part.type, target);
    }
  }
}

// ====== Footnotes / Endnotes ======
async function mergeNotes(baseZip: JSZip, srcZip: JSZip, kind: 'footnotes'|'endnotes'){
  const part = kind==='footnotes'? P.footnotes : P.endnotes;
  const relPath = 'word/_rels/'+(kind==='footnotes'?'footnotes.xml.rels':'endnotes.xml.rels');
  const skel = kind==='footnotes'? footnotesSkeleton() : endnotesSkeleton();

  const baseNotes = await readXmlOrCreate(baseZip, part, skel);
  const srcNotes = await readXmlOrCreate(srcZip, part, skel);

  const tag = kind==='footnotes'?'footnote':'endnote';
  const baseIds = Array.from(baseNotes.getElementsByTagNameNS(NS.w, tag)).map(n=>+(n.getAttributeNS(NS.w,'id')||n.getAttribute('w:id')||n.getAttribute('id')||'0'));
  let maxId = Math.max(0, ...baseIds.filter(n=>n>=0));

  const idMap = new Map<number, number>();
  for(const n of Array.from(srcNotes.getElementsByTagNameNS(NS.w, tag))){
    const oldId = +(n.getAttributeNS(NS.w,'id') || n.getAttribute('w:id') || n.getAttribute('id') || '0');
    if(oldId<0){ idMap.set(oldId, oldId); continue; }
    const newId = ++maxId; idMap.set(oldId, newId);
    n.setAttributeNS(NS.w,'w:id', String(newId));
    baseNotes.documentElement.appendChild(baseNotes.importNode(n,true));
  }

  (baseZip as any).__noteMap = (baseZip as any).__noteMap || {}; (baseZip as any).__noteMap[kind] = idMap;

  const baseRel = await readXmlOrCreate(baseZip, relPath, relsSkeleton());
  const srcRel = await readXmlOrCreate(srcZip, relPath, relsSkeleton());
  const allocRid = makeRidAllocator(baseRel);
  for(const el of Array.from(srcRel.getElementsByTagNameNS(NS.pr,'Relationship'))){
    const type = el.getAttribute('Type')!;
    const target = el.getAttribute('Target')!;
    const mode = el.getAttribute('TargetMode')||'';
    const existing = findRelationshipByTypeAndTarget(baseRel, type, target, mode||undefined);
    if(existing) continue;
    const newId = allocRid();
    addRelationship(baseRel, newId, type, target, mode||undefined);
    if(!mode){
      const srcFull = normalizePath(part, target);
      const data = await srcZip.file(srcFull)?.async('uint8array');
      if(data) baseZip.file(uniquePartPath(baseZip, srcFull), data);
    }
  }

  await writeXml(baseZip, part, baseNotes);
  await writeXml(baseZip, relPath, baseRel);
}

async function getNoteIdRemap(baseZip: JSZip, _srcZip: JSZip, kind: 'footnotes'|'endnotes'){
  return ((baseZip as any).__noteMap && (baseZip as any).__noteMap[kind]) ? (baseZip as any).__noteMap[kind] as Map<number,number> : new Map<number,number>();
}

function remapNoteRefsInBody(body: Element, idMap: Map<number,number>){
  if(!idMap || idMap.size===0) return;
  for(const tag of ['footnoteReference','endnoteReference']){
    const refs = body.getElementsByTagNameNS(NS.w, tag);
    for(const r of Array.from(refs)){
      const v = +(r.getAttributeNS(NS.w,'id') || r.getAttribute('w:id') || r.getAttribute('id') || 'NaN');
      if(idMap.has(v)) (r as Element).setAttributeNS(NS.w,'w:id', String(idMap.get(v)));
    }
  }
}

// ====== Relationships ======
async function remapRelationship(oldRid: string, srcRelXml: Document, baseRelXml: Document, srcZip: JSZip, baseZip: JSZip, nextRid: ()=>string, ensureCT: (ext:string)=>void){
  if(!oldRid) return null;
  const srcRel = findRelationshipById(srcRelXml, oldRid);
  if(!srcRel) return null;
  const type = srcRel.getAttribute('Type')!;
  const target = srcRel.getAttribute('Target')!;
  const mode = srcRel.getAttribute('TargetMode')||undefined;

  const existing = findRelationshipByTypeAndTarget(baseRelXml, type, target, mode);
  if(existing) return existing.getAttribute('Id');

  const newId = nextRid();

  if(type.endsWith('/hyperlink')){
    addRelationship(baseRelXml, newId, type, target, mode||'External');
    return newId;
  }
  if(type.endsWith('/image')){
    const srcPath = normalizePath(P.document, target);
    const srcMediaPath = canonicalizeMediaPath(srcPath);
    const file = srcZip.file(srcMediaPath);
    if(file){
      const data = await file.async('uint8array');
      const { name, ext } = splitNameExt(srcMediaPath);
      const dest = uniqueMediaName(baseZip, name, ext);
      ensureCT(ext);
      baseZip.file(dest, data);
      const relTarget = relativeTo(P.document, dest);
      addRelationship(baseRelXml, newId, type, relTarget, mode);
      return newId;
    }
    addRelationship(baseRelXml, newId, type, target, mode);
    return newId;
  }

  const srcFull = normalizePath(P.document, target);
  const srcFile = srcZip.file(srcFull);
  if(srcFile){
    const data = await srcFile.async('uint8array');
    const unique = uniquePartPath(baseZip, srcFull);
    baseZip.file(unique, data);
    const relTarget = relativeTo(P.document, unique);
    addRelationship(baseRelXml, newId, type, relTarget, mode);
    return newId;
  }

  addRelationship(baseRelXml, newId, type, target, mode);
  return newId;
}

function findRelationshipById(relXml: Document, id: string){
  const list = relXml.getElementsByTagNameNS(NS.pr,'Relationship');
  for(const el of Array.from(list)){ if(el.getAttribute('Id')===id) return el; }
  return null;
}
function findRelationshipByTypeAndTarget(relXml: Document, type: string, target: string, mode?: string){
  const list = relXml.getElementsByTagNameNS(NS.pr,'Relationship');
  for(const el of Array.from(list)){
    if(el.getAttribute('Type')===type && el.getAttribute('Target')===target && ((el.getAttribute('TargetMode')||'')===(mode||''))) return el;
  }
  return null;
}
function addRelationship(relXml: Document, id: string, type: string, target: string, mode?: string){
  const root = relXml.documentElement;
  const rel = relXml.createElementNS(NS.pr,'Relationship');
  rel.setAttribute('Id', id);
  rel.setAttribute('Type', type);
  rel.setAttribute('Target', target);
  if(mode) rel.setAttribute('TargetMode', mode);
  root.appendChild(rel);
}

// ====== XML & ZIP helpers ======
async function readXml(zip: JSZip, path: string){
  const f = zip.file(path); if(!f) throw new Error('Missing '+path);
  const txt = await f.async('text');
  return parseXml(txt);
}
async function readXmlOrCreate(zip: JSZip, path: string, fallback: string){
  const f = zip.file(path); if(!f) return parseXml(fallback);
  const txt = await f.async('text'); return parseXml(txt);
}
async function writeXml(zip: JSZip, path: string, xmlDoc: Document){ zip.file(path, serialize(xmlDoc)); }
function parseXml(s: string){
  const doc = new DOMParser().parseFromString(s,'application/xml');
  const err = doc.querySelector('parsererror'); if(err) throw new Error('XML parse error: '+(err.textContent||'unknown'));
  return doc;
}
function serialize(doc: Document){ return new XMLSerializer().serializeToString(doc); }

function bodyEl(doc: Document){ return doc.getElementsByTagNameNS(NS.w,'body')[0] as Element; }
function lastSectPr(body: Element){ const nodes = body?body.getElementsByTagNameNS(NS.w,'sectPr'):[] as any; return nodes && nodes.length? nodes[nodes.length-1] as Element : null; }
function descendants(node: Element){ const out: Element[]=[]; const w=document.createTreeWalker(node, NodeFilter.SHOW_ELEMENT, null); while(w.nextNode()) out.push(w.currentNode as Element); return out; }
function makePageBreak(doc: Document){ const p=doc.createElementNS(NS.w,'w:p'); const r=doc.createElementNS(NS.w,'w:r'); const br=doc.createElementNS(NS.w,'w:br'); br.setAttributeNS(NS.w,'w:type','page'); r.appendChild(br); p.appendChild(r); return p; }

function relsSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.pr}"></Relationships>`; }
function numberingSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${NS.w}"></w:numbering>`; }
function stylesSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${NS.w}"></w:styles>`; }
function footnotesSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${NS.w}"></w:footnotes>`; }
function endnotesSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="${NS.w}"></w:endnotes>`; }
function fontTableSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="${NS.w}"></w:fonts>`; }
function themeSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="E15759"/></a:accent2><a:accent3><a:srgbClr val="70AD47"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="843C0C"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`; }
function webSettingsSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:webSettings xmlns:w="${NS.w}"></w:webSettings>`; }
function settingsSkeleton(){ return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${NS.w}"></w:settings>`; }

function makeRidAllocator(relXml: Document){ return function(){ const list=relXml.getElementsByTagNameNS(NS.pr,'Relationship'); let max=0; for(const el of Array.from(list)){ const m=/rId(\d+)/.exec(el.getAttribute('Id')||''); if(m) max=Math.max(max,+m[1]); } return 'rId'+(max+1); } }

function splitNameExt(path: string){ const i=path.lastIndexOf('.'); const ext=i>=0?path.substring(i+1).toLowerCase():''; return { name:i>=0?path.substring(0,i):path, ext } }
function uniqueMediaName(zip: JSZip, baseName: string, ext: string){ const folder='word/media/'; const base=baseName.substring(baseName.lastIndexOf('/')+1).replace(/[^A-Za-z0-9_\-]/g,'_'); let i=1,out: string; do{ out=`${folder}merged_${base}_${i}.${ext}`; i++; }while(zip.file(out)); return out; }
function uniquePartPath(zip: JSZip, srcFull: string){ const i=srcFull.lastIndexOf('.'); const ext=i>=0?srcFull.substring(i):''; const stem=i>=0?srcFull.substring(0,i):srcFull; let n=1,cand: string; do{ cand=`${stem}_merged_${n}${ext}`; n++; }while(zip.file(cand)); return cand; }
function normalizePath(base: string, target: string){ const baseDir=base.substring(0,base.lastIndexOf('/')+1); const stack=(baseDir+target).split('/'); const out: string[]=[]; for(const seg of stack){ if(seg==='..') out.pop(); else if(seg==='.'||seg==='') continue; else out.push(seg);} return out.join('/'); }
function canonicalizeMediaPath(p: string){ if(p.startsWith('word/media/')) return p; const idx=p.indexOf('word/'); if(idx>=0){ const rest=p.substring(idx+5); if(rest.startsWith('media/')) return 'word/'+rest; } return p; }
function relativeTo(from: string,to: string){ const fromDir=from.substring(0,from.lastIndexOf('/')+1); return to.startsWith(fromDir)? to.substring(fromDir.length) : to; }

function makeContentTypesEnsurer(ctXml: Document){
  return function ensure(ext: string){
    const defaults = ctXml.getElementsByTagName('Default');
    for(const d of Array.from(defaults)){ if((d.getAttribute('Extension')||'').toLowerCase()===ext.toLowerCase()) return; }
    const map: Record<string,string> = { bmp:'image/bmp', gif:'image/gif', jpg:'image/jpeg', jpeg:'image/jpeg', png:'image/png', tif:'image/tiff', tiff:'image/tiff', emf:'image/x-emf', wmf:'image/x-wmf' };
    const d = ctXml.createElement('Default'); d.setAttribute('Extension', ext); d.setAttribute('ContentType', map[ext]||'application/octet-stream'); ctXml.documentElement.appendChild(d);
  }
}

function findParagraphsContainingText(doc: Document, text: string){
  const body = bodyEl(doc); const ps = Array.from(body.getElementsByTagNameNS(NS.w,'p'));
  const matches: Element[]=[];
  for(const p of ps){
    let t='';
    const runs = p.getElementsByTagNameNS(NS.w,'t');
    for(const r of Array.from(runs)){ t += r.textContent || ''; }
    if(t.includes(text)) matches.push(p);
  }
  return matches;
}

export function triggerDownload(url: string, name: string){ const a=document.createElement('a'); a.href=url; a.download=name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1500); }

function withDefaults(o: MergeOptions): RequiredMergeOptions{
  return {
    pattern: o.pattern??null,
    insertStart: !!o.insertStart,
    insertEnd: !!o.insertEnd,
    pageBreaks: o.pageBreaks!==false,
    mergeNumbering: o.mergeNumbering!==false,
    mergeStyles: o.mergeStyles!==false,
    mergeFootnotes: o.mergeFootnotes!==false,
    onLog: o.onLog
  };
}

type RequiredMergeOptions = Required<Omit<MergeOptions,'pattern'|'onLog'>> & { pattern: string|null, onLog?: MergeOptions['onLog'] };

async function toBytes(src: ArrayBuffer|Uint8Array|Blob): Promise<Uint8Array>{
  if(src instanceof Uint8Array) return src;
  if(src instanceof Blob){
    const anySrc: any = src as any;
    const ab: ArrayBuffer = typeof anySrc.arrayBuffer === 'function' ? await anySrc.arrayBuffer() : await new Response(src).arrayBuffer();
    return new Uint8Array(ab);
  }
  // ArrayBuffer
  return new Uint8Array(src);
}
