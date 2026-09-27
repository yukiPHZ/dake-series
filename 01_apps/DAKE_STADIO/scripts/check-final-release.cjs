'use strict';
const {root}=require('./local-test-env.cjs')();const fs=require('node:fs/promises'),{createReadStream}=require('node:fs'),path=require('node:path'),crypto=require('node:crypto'),assert=require('node:assert/strict');
async function hash(file){const value=crypto.createHash('sha256');for await(const chunk of createReadStream(file))value.update(chunk);return value.digest('hex')}
async function run(){
 const version=require('../package.json').version,distribution=path.join(root,'dist','DAKE_STADIO-win32-x64');
 const archive=JSON.parse(await fs.readFile(path.join(root,'evidence','archive-results.json'),'utf8'));
 assert.ok(archive.passed);assert.equal(path.basename(archive.zip),'DAKE_STADIO-'+version+'-win32-x64.zip');
 const directory=JSON.parse(await fs.readFile(path.join(root,'evidence','packaged-directory-results.json'),'utf8'));
 const extracted=JSON.parse(await fs.readFile(path.join(root,'evidence','packaged-extracted-results.json'),'utf8'));
 assert.ok(directory.passed&&extracted.passed);assert.deepEqual(directory.runtimeSha256,extracted.runtimeSha256);
 const runtimeSha256={executable:await hash(path.join(distribution,'DAKE_STADIO.exe')),appAsar:await hash(path.join(distribution,'resources','app.asar'))};
 assert.deepEqual(runtimeSha256,directory.runtimeSha256);
 const zipSha256=await hash(archive.zip);assert.equal(zipSha256,archive.zipSha256);
 const docs={};for(const name of ['README.md','ORIGINAL.md','TEST_REPORT.md','release_body.md']){const expected=await hash(path.join(root,name));assert.equal(await hash(path.join(distribution,name)),expected);assert.equal(await hash(path.join(archive.extracted,name)),expected);docs[name]=expected}
 const migration=JSON.parse(await fs.readFile(path.join(root,'artifacts','migration-0.1.0','manifest.json'),'utf8'));
 for(const [relative,entry]of Object.entries(migration.files))assert.equal(await hash(path.join(migration.source,relative)),entry.before,'Original changed: '+relative);
 assert.equal(await hash(path.join(root,'artifacts','migration-0.1.0',migration.baselinePackage)),migration.baselineSha256);
 for(const kind of ['banner','flyer','logo','card'])for(const ext of ['dake','png','svg','pdf'])assert.ok((await fs.stat(path.join(distribution,'examples',kind+'.'+ext))).size>100);
 const manual=JSON.parse(await fs.readFile(path.join(root,'evidence','manual-packaged-v2-results.json'),'utf8'));assert.ok(manual.passed);
 const report={passed:true,version,testedRuntimeMatchesFinalArchive:true,finalDocsExactlySynced:true,exampleIncluded:true,editableExamples:4,originalBaselineUnchanged:true,originalFilesVerified:Object.keys(migration.files).length,zipBytes:(await fs.stat(archive.zip)).size,zipSha256,runtimeSha256,docsSha256:docs,archiveFileCount:archive.fileCount,manualPackagedPassed:true};
 await fs.writeFile(path.join(root,'evidence','final-release-check.json'),JSON.stringify(report,null,2));
 await fs.writeFile(archive.zip+'.sha256',zipSha256+'  '+path.basename(archive.zip)+'\n');console.log(JSON.stringify(report,null,2));
}
run().catch(error=>{console.error(error);process.exitCode=1});
