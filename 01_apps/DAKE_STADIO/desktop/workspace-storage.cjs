'use strict';
const fs=require('node:fs/promises');
const storage=require('./storage.cjs');
function validateWorkspace(data){
 if(!data||data.format!=='dake-workspace'||data.version!==1||!Array.isArray(data.documents)||data.documents.length>16||!data.documents.length)throw new Error('WORKSPACE_INVALID');
 const ids=new Set();for(const item of data.documents){if(!item||typeof item.id!=='string'||item.id.length>160||ids.has(item.id)||!item.meta||typeof item.meta.dirty!=='boolean')throw new Error('WORKSPACE_INVALID');ids.add(item.id);storage.validateProject(item.data);if(item.meta.path!==null&&item.meta.path!==undefined&&(typeof item.meta.path!=='string'||item.meta.path.length>8192))throw new Error('WORKSPACE_INVALID');}
 if(!ids.has(data.activeId))throw new Error('WORKSPACE_INVALID');
 for(const item of data.documents){if(!item.child)continue;const seen=new Set([item.id]);let cursor=item;while(cursor.child){if(typeof cursor.child.layerId!=='string'||seen.has(cursor.child.parentId))throw new Error('COMPONENT_CYCLE');seen.add(cursor.child.parentId);cursor=data.documents.find(d=>d.id===cursor.child.parentId);if(!cursor)throw new Error('WORKSPACE_INVALID');}}
 if(Buffer.byteLength(JSON.stringify(data))>storage.MAX_BYTES)throw new Error('WORKSPACE_TOO_LARGE');return data;
}
async function write(file,data){validateWorkspace(data);await storage.atomicWrite(file,JSON.stringify({timestamp:new Date().toISOString(),data}),{backup:true});return true;}
async function read(file){let found=false;for(const candidate of [file,file+'.bak']){try{const bytes=await storage.readLimited(candidate);found=true;const value=JSON.parse(bytes.toString('utf8'));validateWorkspace(value.data);return value;}catch(e){if(e.code==='ENOENT')continue;found=true;}}if(found)throw new Error('WORKSPACE_INVALID');return null;}
module.exports={validateWorkspace,write,read,discard:storage.discardRecovery};
