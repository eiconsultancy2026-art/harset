'use strict';
const { XMLParser, XMLValidator } = require('fast-xml-parser');
const parser = new XMLParser({ ignoreAttributes:false, attributeNamePrefix:'@_', trimValues:true, parseTagValue:false, parseAttributeValue:false });

function decodeXmlEntities(value) {
  return String(value ?? '')
    .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
    .replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#39;/g,"'")
    .replace(/&#13;|&#10;/g,' ');
}
function safeString(value, fallback='') {
  if (value === null || value === undefined) return fallback;
  if (typeof value === 'object') return value['#text'] !== undefined ? safeString(value['#text'], fallback) : fallback;
  return decodeXmlEntities(String(value)).trim();
}
function safeNumber(value, fallback=0) { const n=Number(value); return Number.isFinite(n)?n:fallback; }
function normalizeAmount(value,{debitPositive=true,creditNegative=true}={}) {
  if (value===null||value===undefined||value==='') return 0;
  if (typeof value==='number') return Number.isFinite(value)?Number(value.toFixed(2)):0;
  let raw=decodeXmlEntities(String(value)).trim().replace(/₹/g,'').replace(/Rs\.?/gi,'').replace(/\s+/g,' ');
  if (!raw) return 0;
  const upper=raw.toUpperCase();
  const hasDr=/\bDR\b|\bDEBIT\b/.test(upper), hasCr=/\bCR\b|\bCREDIT\b/.test(upper);
  const parenthesized=/^\s*\(.*\)\s*(DR|CR|DEBIT|CREDIT)?\s*$/i.test(raw);
  raw=raw.replace(/\bDR\b|\bCR\b|\bDEBIT\b|\bCREDIT\b/gi,'').trim();
  if(raw.startsWith('(')&&raw.endsWith(')')) raw=raw.slice(1,-1);
  raw=raw.replace(/,/g,'').replace(/[^0-9.+-]/g,'');
  if(!raw) return 0;
  let n=Number.parseFloat(raw); if(!Number.isFinite(n)) return 0;
  if(parenthesized) n=-Math.abs(n); else if(hasDr&&debitPositive) n=Math.abs(n); else if(hasCr&&creditNegative) n=-Math.abs(n);
  return Number(n.toFixed(2));
}
function getPath(object,pathExpression,fallback=undefined) { if(object==null) return fallback; if(!pathExpression) return object; let current=object; for(const part of String(pathExpression).split('.').filter(Boolean)){ if(current==null||typeof current!=='object'||!(part in current)) return fallback; current=current[part]; } return current===undefined?fallback:current; }
function asArray(value) { if(value==null)return []; return Array.isArray(value)?value:[value]; }
function parseTallyXml(xml) { if(typeof xml!=='string'||!xml.trim()) throw new Error('Tally XML parser received an empty response'); const valid=XMLValidator.validate(xml); if(valid!==true) throw new Error(`Tally XML validation failed: ${valid?.err?.msg||'Invalid XML'}`); try{return parser.parse(xml);}catch(e){throw new Error(`Unable to parse Tally XML: ${e.message}`);} }
function getTallyStatus(parsedXml){ return safeString(getPath(parsedXml,'ENVELOPE.BODY.DESC.STATICVARIABLES.STATUS','')); }
function extractTagValue(xml, tag) { const re=new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`,'i'); const m=String(xml||'').match(re); return m?safeString(m[1]):''; }
function extractTagAmount(xml, tag) { return normalizeAmount(extractTagValue(xml,tag)); }
function extractNamedAmountRows(xml, nameTag='DSPDISPNAME', amountTags=['BSSUBAMT','BSMAINAMT','PLSUBAMT']) {
  const rows=[]; const re=new RegExp(`<${nameTag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${nameTag}>`, 'gi'); let m;
  const source=String(xml||'');
  while((m=re.exec(source))){ const name=safeString(m[1]); const tail=source.slice(re.lastIndex, re.lastIndex+2500); let amount=0, sourceTag=''; for(const tag of amountTags){ const a=extractTagAmount(tail,tag); if(a!==0 || new RegExp(`<${tag}(?:\\s[^>]*)?>\\s*</${tag}>`,'i').test(tail)){ amount=a; sourceTag=tag; break; } } rows.push({name,amount,sourceTag}); }
  return rows.filter(r=>r.name);
}
function normalizeFinancialObject(input={}) { if(input==null||typeof input!=='object')return {}; if(Array.isArray(input))return input.map(normalizeFinancialObject); const out={}; for(const [k,v] of Object.entries(input)){ if(v==null){out[k]='';continue;} if(typeof v==='number'){out[k]=safeNumber(v);continue;} if(typeof v==='string'){out[k]=/[0-9]/.test(v)&&/(,|\.\d+|\bDR\b|\bCR\b|\(|\))/i.test(v)?normalizeAmount(v):safeString(v);continue;} out[k]=Array.isArray(v)?v.map(x=>typeof x==='object'?normalizeFinancialObject(x):x):typeof v==='object'?normalizeFinancialObject(v):v; } return out; }
const amountAt=(obj,p,fallback=0)=>normalizeAmount(getPath(obj,p,fallback));
const stringAt=(obj,p,fallback='')=>safeString(getPath(obj,p,fallback),fallback);
module.exports={parser,safeString,safeNumber,normalizeAmount,parseTallyXml,getPath,asArray,getTallyStatus,extractTagValue,extractTagAmount,extractNamedAmountRows,normalizeFinancialObject,amountAt,stringAt,decodeXmlEntities};
