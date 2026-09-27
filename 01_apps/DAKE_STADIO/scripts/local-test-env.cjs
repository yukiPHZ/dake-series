'use strict';
const fs=require('node:fs'),path=require('node:path');
module.exports=function localTestEnvironment() {
 const root=path.resolve(__dirname,'..');
 const temp=path.join(root,'test-output','tool-temp'),npmCache=path.join(root,'artifacts','npm-cache'),electronCache=path.join(root,'artifacts','electron-download-cache');
 for(const dir of [temp,npmCache,electronCache])fs.mkdirSync(dir,{recursive:true});
 process.env.PYTHONDONTWRITEBYTECODE='1';
 process.env.TEMP=temp;process.env.TMP=temp;process.env.TMPDIR=temp;
 process.env.npm_config_cache=npmCache;process.env.ELECTRON_CACHE=electronCache;
 return {root,temp,npmCache,electronCache};
};
