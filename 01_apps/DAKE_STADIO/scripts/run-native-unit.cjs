'use strict';
const {root}=require('./local-test-env.cjs')();const fs=require('node:fs'),path=require('node:path'),{spawnSync}=require('node:child_process');
const files=fs.readdirSync(path.join(root,'tests')).filter(name=>name.endsWith('.test.cjs')).map(name=>path.join('tests',name));
const result=spawnSync(process.execPath,['--test','--test-reporter=tap',...files],{cwd:root,encoding:'utf8',windowsHide:true});
const output=(result.stdout||'')+(result.stderr||'');const report={passed:result.status===0,exitCode:result.status,tests:Number(/# tests (\d+)/.exec(output)?.[1]||0),passes:Number(/# pass (\d+)/.exec(output)?.[1]||0),failures:Number(/# fail (\d+)/.exec(output)?.[1]||0),files,output};
fs.writeFileSync(path.join(root,'evidence','native-unit-results.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({...report,output:undefined},null,2));if(result.error)throw result.error;process.exitCode=result.status;
