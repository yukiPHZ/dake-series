
'use strict';
require('./local-test-env.cjs')();
const path=require('node:path'),{spawn}=require('node:child_process'),root=path.resolve(__dirname,'..');
const environment={...process.env};delete environment.ELECTRON_RUN_AS_NODE;
const child=spawn(require('electron'),['scripts/check-boundaries.cjs'],{cwd:root,env:environment,stdio:'inherit',windowsHide:true});
const timer=setTimeout(()=>{console.error('Boundary checks timed out');child.kill();},240000);
child.once('error',error=>{clearTimeout(timer);console.error(error);process.exitCode=1;});
child.once('exit',code=>{clearTimeout(timer);process.exitCode=code??1;});

