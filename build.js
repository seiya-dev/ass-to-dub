#!/usr/bin/env node

// build requirements
const pkg = require('./package.json');
const fse = require('fs-extra');
const modulesCleanup = require('removeNPMAbsolutePaths');
const { compile } = require('nexe');

// main
(async function(){
    const buildStr = `${pkg.name}-${pkg.version}`;
    const acceptableBuilds = ['win64','linux64','macos64'];
    const buildType = process.argv[2];
    if(!acceptableBuilds.includes(buildType)){
        console.error(`[ERROR] unknown build type!`)
        process.exit();
    }
    await modulesCleanup('.');
    if(!fse.existsSync(`./builds`)){
        fse.mkdirSync(`./builds`);
    }
    const buildFull = `${buildStr}-${buildType}`;
    const buildDir = `./builds/${buildFull}`;
    if(fse.existsSync(buildDir)){
        fse.removeSync(buildDir);
    }
    fse.mkdirSync(buildDir);
    fse.mkdirSync(`${buildDir}/subtitles`);
    const buildConfig = {
        input: './main.js',
        output: `${buildDir}/${pkg.name}`,
        target: getTarget(buildType),
        resources: [
            './language/*',
        ],
    };
    console.log(`[Build] Build configuration: ${buildFull}`);
    await compile(buildConfig);
    const pCfgFile = `set_config.json`;
    if(fse.existsSync(`./${pCfgFile}`)){
        fse.copyFileSync(`./${pCfgFile}`, `${buildDir}/${pCfgFile}`);
    }
}());

function getTarget(bt){
    switch(bt){
        case 'win64':
            return 'windows-x64';
        case 'linux64':
            return 'linux-x64';
        case 'macos64':
            return 'macos-x64';
        default:
            return 'windows-x64';
    }
}
