#!/usr/bin/env node

// build-in
const path = require('path');
const fs = require('fs');

// plugins
const { Document, Packer, Table, Paragraph, TextRun, } = require('docx');
let file = '';

// main
(async function(){
    require('process').chdir(`${__dirname}/subtitles/`)
    let fls = filterByRegx(fs.readdirSync('.'),/(?<!\.Dub)\.ass$/);
    for(let f of fls){
        file = f;
        if(fs.existsSync(file)){
            console.log(`[INFO] Processing ${file}...`);
            await parseFile();
            console.log(`[INFO] Done!`);
        }
    }
}());

function filterByRegx(arr,regx) {
    return arr.filter(function(el) {
        return el.match(new RegExp(regx));
    });
}

// parse
async function parseFile(){
    let subs = fs.readFileSync(file, 'utf8');
    subs = subs.replace(/^\ufeff/,'').replace(/\r/g,'').split('\n');
    // sections
    let section = '',
        ass = {
            script_info: {},
            styles: {
                format: [],
                list: [],
            },
            events: {
                format: [],
                dialogue: [],
            },
            roles: {},
        };
    // collect
    for(let lineIndex in subs){
        let s = subs[lineIndex];
        // get section
        if(s.match(/^\[(.*)\]$/)){
            let mSec = s.match(/^\[(.*)\]$/)[1];
            switch(mSec) {
                case 'Script Info':
                    section = 'script_info';
                    break;
                case 'V4+ Styles':
                    section = 'v4_styles';
                    break;
                case 'Events':
                    section = 'events';
                    break;
                case 'Aegisub Project Garbage':
                    section = '';
                    break;
                default:
                    console.log(`[WARN] Unknown Section:`, mSec);
                    section = '';
            }
            continue;
        }
        // get commentary
        if(s.match(/^; (.*)/) || s.match(/^!: (.*)/)){
            let cmt = '';
            if(s.match(/^; (.*)/)){
                cmt = s.match(/^; (.*)/)[1];
            }
            if(s.match(/^!: (.*)/)){
                cmt = s.match(/^!: (.*)/)[1];
            }
            console.log(`[COMMENTARY]`, cmt);
            continue;
        }
        // get strings
        if(section != '' && s != ''){
            if(section == 'script_info'){
                let type = s.split(':')[0];
                let parm = s.replace(new RegExp(`^${type}: `),'');
                ass.script_info[type] = parm;
            }
            if(section == 'v4_styles'){
                let type = s.split(':')[0];
                let parm = s.replace(new RegExp(`^${type}: `),'');
                if(type == 'Format'){
                    ass.styles.format = parm.split(', ');
                    continue;
                }
                else{
                    parm = parm.split(',');
                }
                if(ass.styles.format.length > 0 && type == 'Style'){
                    let current = Object.assign(...ass.styles.format.map((k, i) => ({[k]: parm[i]})));
                    current = Object.assign({TextParam: parm.join(',')}, current);
                    ass.styles.list.push(current);
                }
            }
            if(section == 'events'){
                let type = s.split(':')[0];
                let parm = s.replace(new RegExp(`^${type}: `),'');
                let ptxt = '', ctxt = '';
                if(type == 'Format'){
                    ass.events.format = parm.split(', ');
                    continue;
                }
                else{
                    parm = parm.split(',');
                    ptxt = parm.slice(9).join(',');
                    parm = parm.slice(0, 9);
                    cprm = parm;
                    ctxt = ptxt.replace(/\{[^}]*\}/g, '');
                    parm.push(ptxt);
                }
                if(ass.events.format.length > 0 && type == 'Dialogue' && ptxt != ''){
                    let current = Object.assign(...ass.events.format.map((k, i) => ({[k]: parm[i]})));
                    current = Object.assign({CleanText: ctxt}, current);
                    let roleNames = current.Name.replace(/;;+/,';').replace(/^;+/,'').replace(/;+$/,'').split(';');
                    for(let r of roleNames){
                        let scurrent = current;
                        scurrent.Name = r.trim();
                        cprm[4] = scurrent.Name;
                        scurrent = Object.assign({CleanParam: cprm.join(',')}, scurrent);
                        ass.events.dialogue.push(scurrent);
                        if(scurrent.Name == ''){
                            console.log(`[WARN] Role name is missing on line ${lineIndex}`);
                        }
                        scurrent.Name = scurrent.Name.toUpperCase();
                        if(!ass.roles[scurrent.Name]){
                            ass.roles[scurrent.Name] = 0;
                        }
                        ass.roles[scurrent.Name]++;
                    }
                }
            }
        }
    }
    if(ass.script_info.ScriptType != 'v4.00+'){
        console.log(`[WARN] Supported only script types v4.00+!`);
        process.exit();
    }
    else{
        // make files
        let assFile = '';
        let srtFile = '';
        let docFile = new Document();
        let docRole = new Document();
        let fileName = path.join(file.replace(/(\\|\/)+$/g,'').replace(/\.ass$/,''));
        // make role list
        let rows = Object.keys(ass.roles).length+1;
        let roleTable = new Table({
            rows: rows,
            columns: 2,
            width: 100,
            widthUnitType: 'pct',
            margins: { left: '0.2cm', right: '0.2cm', },
            float: { relativeHorizontalPosition: 'center', },
        });
        roleTable.getCell(0, 0).Properties.setWidth('50%');
        roleTable.getCell(0, 1).Properties.setWidth('50%');
        roleTable.getCell(0, 0).addParagraph(new Paragraph('').addRun(new TextRun('Персонаж').bold()));
        roleTable.getCell(0, 1).addParagraph(new Paragraph('').addRun(new TextRun('Число строк').bold()));
        for(let r in Object.keys(ass.roles)){
            r = parseInt(r);
            let role = Object.keys(ass.roles)[r];   
            roleTable.getCell(r+1, 0).Properties.setWidth('50%');
            roleTable.getCell(r+1, 1).Properties.setWidth('50%');
            roleTable.getCell(r+1, 0).addParagraph(new Paragraph(role));
            roleTable.getCell(r+1, 1).addParagraph(new Paragraph(ass.roles[role]));
        }
        docRole.addTable(roleTable);
        let docRoleCont = await new Packer().toBuffer(docRole);
        try{
            fs.writeFileSync(`${fileName}.Roles.docx`, docRoleCont);
        }
        catch(e){
            console.log(e);
            console.log(`[ERROR] File ${fileName}.Roles.docx not saved!`);
        }
        // make new ass, srt and docx
        assFile = [
            `[Script Info]`,
            `Title: ${ass.script_info.Title}`,
            `Original Translation: `,
            `Original Editing: `,
            `Original Timing: `,
            `Synch Point: `,
            `Script Updated By: `,
            `Update Details: `,
            `ScriptType: v4.00+`,
            `Collisions: Normal`,
            `PlayResX: ${ass.script_info.PlayResX}`,
            `PlayResY: ${ass.script_info.PlayResY}`,
            `Timer: 0.0000`,
            `WrapStyle: ${ass.script_info.WrapStyle}`,
            `\r\n`,
        ].join(`\r\n`);
        // restore styles
        let assFileStyles = [`[V4+ Styles]`];
        assFileStyles.push(`Format: ${ass.styles.format.join(', ')}`);
        for(let s of ass.styles.list){
            assFileStyles.push(`Style: ${s.TextParam}`);
        }
        assFileStyles.push(`\r\n`);
        assFile += assFileStyles.join(`\r\n`);
        // ass header
        let assFileEvents = [`[Events]`];
        assFileEvents.push(`Format: ${ass.events.format.join(', ')}`);
        // prep doc table
        let docArr = [],
            current_row = -1,
            current_actor = undefined;
        // make subs and docx
        for(let s in ass.events.dialogue){
            s = parseInt(s);
            let dlgs = ass.events.dialogue;
            let dlgp = dlgs[s-1];
            let dlgc = dlgs[s];
            let actor = dlgc.Name.toUpperCase();
            assFileEvents.push(`Dialogue: ${dlgc.CleanParam},[${dlgc.Name.toUpperCase()}] ${dlgc.CleanText}`);
            srtFile += `${s+1}\r\n`;
            srtFile += `${assTimeToSrt(dlgc.Start)} --> ${assTimeToSrt(dlgc.End)}\r\n`;
            srtFile += `[${actor}] ${dlgc.CleanText.replace(/\\n/gi,'\r\n')}\r\n\r\n`;
            let cleanDialogDocx = dlgc.CleanText.replace(/\\n/gi,' ').replace(/  +/g,' ').trim();
            let cleanPrevDlDocx = dlgp ? dlgp.CleanText.replace(/\\n/gi,' ').replace(/  +/g,' ').trim() : '';
            if(actor == ''){
                current_actor = undefined;
            }
            if(current_actor != actor){
                current_actor = actor;
                docArr.push({
                    time: assTimeToDoc(dlgc.Start),
                    actor: current_actor,
                    text: cleanDialogDocx,
                });
                current_row++;
            }
            else{
                let startStrMatch = /^\(\d{1,2}(-|:)\d{2}\) |^\d{1,2}(-|:)\d{2} /;
                if(cleanPrevDlDocx.slice(-2) == '//' || assTimeToDoc(dlgc.Start, dlgp.End) == 5 && cleanPrevDlDocx.slice(-1) != '/' ){
                    if(cleanDialogDocx.match(startStrMatch)){ // cleanPrevDlDocx.slice(-2) == '//' &&
                        cleanDialogDocx = cleanDialogDocx.replace(startStrMatch,'');
                    }
                    docArr[current_row].text += ( cleanPrevDlDocx.slice(-2) != '//' ? ' //' : '' ) + assTimeToDoc(dlgc.Start);
                }
                if(cleanPrevDlDocx.slice(-1) == '/'  || assTimeToDoc(dlgc.Start, dlgp.End) == 1 && cleanPrevDlDocx.slice(-1) != '/' ){
                    docArr[current_row].text += ( cleanPrevDlDocx.slice(-1) != '/' ? ' /' : '' );
                }
                docArr[current_row].text += ' ' + cleanDialogDocx;
            }
        }
        // create doc
        let dlgTable = new Table({
            rows: docArr.length,
            columns: 3,
            width: 100,
            widthUnitType: 'pct',
            margins: { left: '0.2cm', right: '0.2cm', },
            float: { relativeHorizontalPosition: 'center', },
        });
        for(let s in docArr){
            s = parseInt(s);
            let str = docArr[s];   
            dlgTable.getCell(s, 0).Properties.setWidth('1.35cm');
            dlgTable.getCell(s, 1).Properties.setWidth('2.15cm');
            // dlgTable.getCell(s, 2).Properties.setWidth('50%');
            dlgTable.getCell(s, 0).addParagraph(new Paragraph(str.time));
            dlgTable.getCell(s, 1).addParagraph(new Paragraph(str.actor));
            dlgTable.getCell(s, 2).addParagraph(new Paragraph(str.text));
        }
        docFile.addTable(dlgTable);
        let docFileCont = await new Packer().toBuffer(docFile);
        // save
        assFileEvents.push(`\r\n`);
        assFile += assFileEvents.join(`\r\n`);
        fs.writeFileSync(`${fileName}.Dub.ass`, assFile);
        fs.writeFileSync(`${fileName}.Dub.srt`, srtFile);
        try{
            fs.writeFileSync(`${fileName}.Dub.docx`, docFileCont);
        }
        catch(e){
            console.log(e);
            console.log(`[ERROR] File ${fileName}.Dub.docx not saved!`);
        }
        // fs.writeFileSync(`${fileName}.Dub.json`, JSON.stringify(ass,null,'    '));
    }
}

function assTimeToSrt(time){
    return time.replace(/\./,',').padStart(11, '0').padEnd(12, '0');
}

function assTimeToDoc(time, timePrev){
    if(!timePrev){
        time = strToTimeArr(time);
        time[2] = Math.round(time[2]);
        if(time[2] > 60){
            time[2] = 0;
            time[1]++;
        }
        if(time[1] > 60){
            time[1] = 0;
            time[0]++;
        }
        time[2] = time[2].toString().padStart(2,'0');
        time[1] = time[1].toString().padStart(2,'0');
        if(time[0] == 0){
            time.shift();
        }
        return time.join(':');
    }
    else{
        let time1 = strToTimeArr(time);
        time1 = time1[0]*60*60 + time1[1]*60 + time1[2];
        let time2 = strToTimeArr(timePrev);
        time2 = time2[0]*60*60 + time2[1]*60 + time2[2];
        if(time1 - time2 > 4.99){
            return 5;
        }
        if(time1 - time2 > 0.99){
            return 1;
        }
        return 0;
    }
}

function strToTimeArr(time){
    time = time.split(':');
    time[0] = parseInt(time[0]);
    time[1] = parseInt(time[1]);
    time[2] = parseFloat(time[2]);
    return time;
}