'use strict';
const fs=require('node:fs/promises'),path=require('node:path'),crypto=require('node:crypto');
function createRecoveryManager({directory,storage,workspaceStorage}) {
  const pendingDir=path.join(directory,'pending-recovery');
  let initialization;
  async function migrate(){
    await fs.mkdir(pendingDir,{recursive:true});
    for(const [kind,name] of [['workspace','workspace-recovery.json'],['legacy','recovery.json']]){
      const original=path.join(directory,name),base=path.join(pendingDir,crypto.randomUUID()+'-'+kind+'.json');
      for(const suffix of ['','.bak']){
        try{await fs.rename(original+suffix,base+suffix);}
        catch(error){if(error.code!=='ENOENT')throw error;}
      }
    }
  }
  const ready=()=>initialization||(initialization=migrate());
  function location(id){
    if(typeof id!=='string'||!/^[a-f0-9-]{36}-(workspace|legacy)$/.test(id))throw new Error('RECOVERY_INVALID');
    return path.join(pendingDir,id+'.json');
  }
  async function readPending(id){
    await ready();const file=location(id),kind=id.endsWith('-workspace')?'workspace':'legacy';
    const result=kind==='workspace'?await workspaceStorage.read(file):await storage.readRecovery(file);
    if(!result)throw new Error('RECOVERY_MISSING');return {id,kind,...result};
  }
  return {
    async listPending(){
      await ready();const files=await fs.readdir(pendingDir),ids=[...new Set(files.map(f=>f.match(/^([a-f0-9-]{36}-(?:workspace|legacy))\.json(?:\.bak)?$/)?.[1]).filter(Boolean))],result=[];
      for(const id of ids){
        try{const value=await readPending(id),docs=value.kind==='workspace'?value.data.documents:[{data:value.data}];result.push({id,kind:value.kind,timestamp:value.timestamp||'',count:docs.length,name:docs.map(d=>d.data.name).filter(Boolean).slice(0,3).join(' / '),unavailable:false});}
        catch{result.push({id,kind:id.endsWith('-workspace')?'workspace':'legacy',name:'',count:0,timestamp:'',unavailable:true});}
      }
      return result.sort((a,b)=>String(b.timestamp).localeCompare(String(a.timestamp)));
    },
    readPending,
    async discardPending(id){await ready();await storage.discardRecovery(location(id));return true;},
    async writeWorkspace(data){await ready();return workspaceStorage.write(path.join(directory,'workspace-recovery.json'),data);},
    async writeLegacy(data){await ready();return storage.writeRecovery(path.join(directory,'recovery.json'),data);}
  };
}
module.exports={createRecoveryManager};
