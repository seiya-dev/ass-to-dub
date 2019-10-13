#!/usr/bin/env node

// build-in
const path = require('path');
const fs = require('fs');

// plugins
const { Document, Packer, Table, Paragraph, TableCell, TableRow, TextRun,
    VerticalAlign, WidthType, TableLayoutType } = require('docx');
let file = '', lang = {}, roles = {};

// predef config
let config = {};
const preConfig = {
    "use_never_join_role": true,
    "never_join_role": "CAPTION",
    "never_join_dialogues": false,
    "use_start_time": true,
    "use_end_time": false,
    "use_full_time_format": false,
    "use_full_time_hide_msec": false,
    "dont_split_subs_actors": false,
    "subs_actor_template": "[{actor}]",
    "subs_actor_template_joiner": "/ ",
    "subs_actor_template_before": "",
    "subs_actor_template_after": " ",
    "skip_create_ass_mod": false,
    "skip_create_srt_mod": false,
    "role_list_format": "txt",
};

// load config
config = Object.assign(config, preConfig);
if(fs.existsSync(`./set_config.json`)){
    let loadedConfig = require(`./set_config.json`);
    config = Object.assign(config, loadedConfig);
}

// main
(async function(){
    console.log(`== Advanced SubStation Alpha to Dialogue List ==`);
    console.log(`==             by  Seiya Loveless             ==`);
    const langRegx = /^set_([a-z]{2})$/;
    let setLangFile = filterByRegx(fs.readdirSync('./language/'),langRegx);
    setLangFile = setLangFile.length > 0 ? setLangFile[0].match(langRegx)[1] : 'en';
    if(fs.existsSync(`./language/${setLangFile}.json`)){
        lang = require(`./language/${setLangFile}.json`);
    }
    else{
        console.error(`[ERROR] Language file not found!`);
        process.exit();
    }
    if(fs.existsSync(`./set_roles.json`)){
        roles = require(`./set_roles.json`);
    }
    if(roles.toString() != '[object Object]'){
        roles = {};
    }
    if(typeof roles.male != 'object' || roles.male === null || roles.male.toString() == '[object Object]'){
        roles.male = [];
    }
    if(typeof roles.female != 'object' || roles.female === null || roles.female.toString() == '[object Object]'){
        roles.female = [];
    }
    require('process').chdir(`${__dirname}/subtitles/`);
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
    let dialogIndex = 0;
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
                    ctxt = ptxt.replace(/\{[^}]*\}/g, '');
                    parm.push(ptxt);
                }
                if(ass.events.format.length > 0 && type == 'Dialogue' && ptxt != ''){
                    dialogIndex++;
                    let cprm = parm.slice(0, 9);
                    let current = Object.assign(...ass.events.format.map((k, i) => ({[k]: parm[i]})));
                    current = Object.assign({CleanText: ctxt}, current);
                    let roleNames = current.Name.replace(/\t/g,' ').replace(/  +/g,' ')
                                        .replace(/;;+/g,';').replace(/^;+/g,'').replace(/;+$/g,'');
                    roleNames = roleNames.split(';').map(r => r.trim());
                    current.id    = dialogIndex;
                    current.Names = roleNames;
                    for(let r of roleNames){
                        let scurrent = current;
                        scurrent.Name = r;
                        if(!config.dont_split_subs_actors){
                            cprm[4] = scurrent.Name;
                        }
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
        // base fn
        let fileName = path.join(file.replace(/(\\|\/)+$/g,'').replace(/\.ass$/,''));
        // make role list
        let sroles = {
            male:[],
            female:[],
            unsorted:[],
        };
        for(let r in Object.keys(ass.roles)){
            r = parseInt(r);
            let role = Object.keys(ass.roles)[r];
            if(roles.male.includes(role)){
                sroles.male.push(`${role}\t${ass.roles[role]}`);
            }
            else if(roles.female.includes(role)){
                sroles.female.push(`${role}\t${ass.roles[role]}`);
            }
            else{
                sroles.unsorted.push(`${role}\t${ass.roles[role]}`);
            }
        }
        // txt roles
        let txtRole = '';
        txtRole = [].concat(
            [`${lang.male}`],[''],sroles.male,[''],
            [`${lang.female}`],[''],sroles.female,[''],
        );
        if(sroles.unsorted.length>0){
            txtRole = txtRole.concat(
                [`${lang.unsorted}`],[''],sroles.unsorted,[''],
            );
        }
        // role list docx
        txtRole.unshift(`${lang.character}\t${lang.dialogues}`,'');
        // docx roles
        let docRoleCont = '';
        let docRole = new Document();
        let rolesRows = [];
        for(let r in txtRole){
            roleStr = txtRole[r];
            roleStr = roleStr.split(`\t`);
            roleStr[1] = !roleStr[1] ? '' : roleStr[1];
            let rolesRow = new TableRow({
                children: [
                    addTableCell(roleStr[0], '50%'),
                    addTableCell(roleStr[1], '50%'),
                ],
            });
            rolesRows.push(rolesRow);
        }
        docRole.addSection({children:[
            addTable(rolesRows)
        ]});
        docRole = fixEmptyParagraph(docRole);
        docRoleCont = await Packer.toBuffer(docRole);
        // save role list
        switch(config.role_list_format){
            case 'docx':
                try{
                    fs.writeFileSync(`${fileName}.Roles.docx`, docRoleCont);
                }
                catch(e){
                    console.log(e);
                    console.log(`[ERROR] File ${fileName}.Roles.docx not saved!`);
                }
                break;
            case 'csv':
                fs.writeFileSync(`${fileName}.Roles.csv`, `\ufeff`+txtRole.join(`\r\n`).replace(/\t/g,';'));
                break;
            case 'txt':
            default:
                fs.writeFileSync(`${fileName}.Roles.txt`, txtRole.join(`\r\n`)); 
        }
        // make new ass, srt and docx
        let assFile = '';
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
        let srtFile = '';
        let docArr = [],
            subtitle_dlg_id = -1,
            current_row = -1,
            current_actor = undefined;
        // make subs and docx
        for(let s in ass.events.dialogue){
            s = parseInt(s);
            let dlgs = ass.events.dialogue;
            let dlgp = dlgs[s-1];
            let dlgc = dlgs[s];
            if(subtitle_dlg_id != dlgc.id){
                let subsActor = config.dont_split_subs_actors 
                    ? dlgc.Names.map(a => config.subs_actor_template.replace(/{actor}/,a)).join(config.subs_actor_template_joiner)
                    : config.subs_actor_template.replace(/{actor}/,dlgc.Name);
                subsActor = config.subs_actor_template_before + subsActor + config.subs_actor_template_after;
                assFileEvents.push(`Dialogue: ${dlgc.CleanParam},${subsActor}${dlgc.CleanText}`);
                srtFile += `${s+1}\r\n`;
                srtFile += `${assTimeToSrt(dlgc.Start)} --> ${assTimeToSrt(dlgc.End)}\r\n`;
                srtFile += `${subsActor}${dlgc.CleanText}`.replace(/\\n/gi,'\r\n') + `\r\n\r\n`;
                if(config.dont_split_subs_actors){
                    subtitle_dlg_id = dlgc.id;
                }
            }
            let actor = dlgc.Name;
            let cleanDialogDocx = dlgc.CleanText.replace(/\\n/gi,' ').replace(/  +/g,' ').trim();
            let cleanPrevDlDocx = dlgp ? dlgp.CleanText.replace(/\\n/gi,' ').replace(/  +/g,' ').trim() : '';
            if(actor == ''){
                current_actor = undefined;
            }
            if(current_actor != actor || config.use_never_join_role 
                    && current_actor == config.never_join_role || config.never_join_dialogues){
                current_actor = actor;
                docArr.push({
                    time: convTimeDoc(dlgc.Start),
                    tend: convTimeDoc(dlgc.End),
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
                docArr[current_row].tend = convTimeDoc(dlgc.End);
            }
        }
        // create doc
        let docFile = new Document();
        let docFileCont = '';
        let dialogRows = [];
        let dialogCellWidths = calcCellWidths();
        for(let s in docArr){
            s = parseInt(s);
            let str = docArr[s];
            let dialogCells = [];
            if(config.use_start_time){
                dialogCells.push(addTableCell(str.time, dialogCellWidths[0] + 'cm'));
            }
            if(config.use_end_time){
                dialogCells.push(addTableCell(str.tend, dialogCellWidths[1] + 'cm'));
            }
            dialogCells.push(addTableCell(str.actor, dialogCellWidths[2] + 'cm'));
            dialogCells.push(addTableCell(str.text, dialogCellWidths[3] + 'cm'));
            let dialogRow = new TableRow({
                children: dialogCells,
            });
            dialogRows.push(dialogRow);
        }
        docFile.addSection({children:[
            addTable(dialogRows)
        ]});
        docFile = fixEmptyParagraph(docFile);
        docFileCont = await Packer.toBuffer(docFile);
        // save
        assFileEvents.push(`\r\n`);
        assFile += assFileEvents.join(`\r\n`);
        if(!config.skip_create_ass_mod){
            fs.writeFileSync(`${fileName}.Dub.ass`, assFile);
        }
        if(!config.skip_create_srt_mod){
            fs.writeFileSync(`${fileName}.Dub.srt`, srtFile);
        }
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

function fixEmptyParagraph(doc){
    // Should fix empty paragraph bug
    let sections = doc.document.body.root.length;
    if(
        sections > 0 &&
        doc.document.body.root[0].rootKey == 'w:p' &&
        doc.document.body.root[0].properties.root.length == 0
    ){
        doc.document.body.root[0].deleted = true;
    }
    return doc;
}

function fixCellWidth(cell, size){
    let { TableCellWidth } = require('docx');
    let cellWidth = new TableCellWidth(size);
    cell.root[0].cellWidth = cellWidth;
    cell.properties.root.push(cellWidth);
    cell.properties.cellWidth = cellWidth;
    return cell;
}

function addTableCell(content, size){
    let cell = new TableCell({
        children: [new Paragraph({ text: content })],
        verticalAlign: VerticalAlign.CENTER,
        width: { size: size, },
    });
    cell = fixCellWidth(cell, size);
    return cell;
}

function addTable(content){
    let table = new Table({
        rows: content,
        width: { size: 100, type: WidthType.PERCENTAGE, },
        margins: { left: '0.2cm', right: '0.2cm', },
        layout: TableLayoutType.FIXED,
        float: { relativeHorizontalPosition: 'center', },
    });
    return table;
}

function calcCellWidths(){
    let w = [ 0, 0, 3.00, 16.32];
    if(config.use_start_time){
        w[0] = getTimeCellWidth();
        w[3] -= w[0];
    }
    if(config.use_end_time){
        w[1] = getTimeCellWidth();
        w[3] -= w[1];
    }
    w[3] -= w[2];
    return w;
}

function getTimeCellWidth(){
    if(config.use_full_time_format && !config.use_full_time_hide_msec){
        return 2.20;
    }
    if(config.use_full_time_format){
        return 1.70;
    }
    return 1.50;
}

function convTimeDoc(time){
    return config.use_full_time_format
                ? ( config.use_full_time_hide_msec ? assFullTimeToDoc(time) : assFullTimeWMSecToDoc(time) )
                : assTimeToDoc(time);
}

function assTimeToSrt(time){
    return time.replace(/\./,',').padStart(11, '0').padEnd(12, '0');
}

function assFullTimeToDoc(time){
    return time.replace(/\.\d{2}$/,'').padStart(8, '0');
}

function assFullTimeWMSecToDoc(time){
    return time.replace(/\./,':').padStart(11, '0');
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
