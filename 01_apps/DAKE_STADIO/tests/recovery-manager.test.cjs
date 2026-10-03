'use strict';
require('../scripts/local-test-env.cjs')();
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path'),os=require('node:os'),crypto=require('node:crypto');
const storage=require('../desktop/storage.cjs'),workspaceStorage=require('../desktop/workspace-storage.cjs'),{createRecoveryManager}=require('../desktop/recovery-manager.cjs');
const doc=name=>({format:'dake-stadio',version:3,documentId:'doc-'+name,width:80,height:80,name,background:'#fff',fonts:[],canvas:{objects:[]}});
const workspace=name=>({format:'dake-workspace',version:1,activeId:'session-'+name,documents:[{id:'session-'+name,data:doc(name),meta:{path:null,dirty:true}}]});
test('unread workspace and legacy recovery survive new autosaves, active discard and restart; explicit discard only consumes selected candidate',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'stadio-pending-'));
 try{
  await workspaceStorage.write(path.join(dir,'workspace-recovery.json'),workspace('old'));
  await workspaceStorage.write(path.join(dir,'workspace-recovery.json'),workspace('old-latest'));
  await storage.writeRecovery(path.join(dir,'recovery.json'),doc('legacy'));
  const manager=createRecoveryManager({directory:dir,storage,workspaceStorage});
  // A new autosave may beat the UI reading pending candidates.
  await manager.writeWorkspace(workspace('new'));
  let candidates=await manager.listPending();assert.equal(candidates.length,2);
  const old=candidates.find(p=>p.kind==='workspace'),legacy=candidates.find(p=>p.kind==='legacy');
  assert.equal((await manager.readPending(old.id)).data.documents[0].data.name,'old-latest');
  assert.equal((await manager.readPending(legacy.id)).data.name,'legacy');
  const pendingFile=path.join(dir,'pending-recovery',old.id+'.json'),hash=crypto.createHash('sha256').update(await fs.readFile(pendingFile)).digest('hex');
  await manager.writeWorkspace(workspace('newer'));await workspaceStorage.discard(path.join(dir,'workspace-recovery.json'));await storage.discardRecovery(path.join(dir,'recovery.json'));
  assert.equal(crypto.createHash('sha256').update(await fs.readFile(pendingFile)).digest('hex'),hash);
  const afterRestart=createRecoveryManager({directory:dir,storage,workspaceStorage});assert.equal((await afterRestart.listPending()).length,2);
  await afterRestart.discardPending(legacy.id);assert.equal((await afterRestart.listPending()).length,1);assert.equal((await afterRestart.readPending(old.id)).data.documents[0].data.name,'old-latest');
  await assert.rejects(()=>afterRestart.readPending('../recovery'),/RECOVERY_INVALID/);
 }finally{assert.equal(path.dirname(path.resolve(dir)),path.resolve(os.tmpdir()));assert.ok(path.basename(dir).startsWith('stadio-pending-'));await fs.rm(dir,{recursive:true,force:true});}
});
