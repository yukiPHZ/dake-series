require('./local-test-env.cjs')();
const {spawnSync}=require('node:child_process');const path=require('node:path');
const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;env.TEMP=env.TMP=path.resolve('test-output/tool-temp');
const child=spawnSync(require('electron'),['scripts/check-native-pdf.cjs'],{stdio:'inherit',env,windowsHide:true,timeout:120000});if(child.error)throw child.error;process.exitCode=child.status;
