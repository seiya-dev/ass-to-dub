#!/usr/bin/env node

// build-in
const path = require('path');
const fs = require('fs');

// plugins
const { Document, Packer, Table, Paragraph, TableCell, TableRow, TextRun,
    VerticalAlign, WidthType, TableLayoutType } = require('docx');

// defined
let file = '', lang = {}, roles = {};
const docxStringSplitter = '{\\r\\n}';

// predef config
let config = {};
const preConfig = {
    "language": "ru",
    "use_never_join_role": true,
    "never_join_role": "CAPTION",
    "never_join_dialogues": false,
    "use_start_time": true,
    "use_end_time": false,
    "use_full_time_format": false,
    "use_full_time_hide_msec": false,
    "dont_split_subs_actors": false,
    "use_docx_string_splitter": false,
    "docx_join_time_short": 1,
    "docx_join_time_long": 5,
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
const configFile = __dirname + `/set_config.json`;
if(fs.existsSync(configFile)){
    let loadedConfig = require(configFile);
    config = Object.assign(config, loadedConfig);
}

// main
(async function(){
    console.log(`== Advanced SubStation Alpha to Dialogue List ==`);
    console.log(`==             by  Seiya Loveless             ==`);
    // set lang
    let langCode        = 'ru.json';
    const langFilesPath = __dirname + `/language/`;
    if(typeof config.language == 'string' && config.language.match(/^[a-z]{2}$/)){
        config.language += '.json';
        const langFiles = fs.readdirSync(langFilesPath);
        langCode = langFiles.includes(config.language) ? config.language : 'ru.json';
    }
    const langFile = langFilesPath + langCode;
    if(fs.existsSync(langFile)){
        lang = require(langFile);
    }
    else{
        console.error(`[ERROR] Language file not found!`);
        await doPause();
        process.exit();
    }
    const setRolesFile = __dirname + `/set_roles.json`;
    if(fs.existsSync(setRolesFile)){
        roles = require(setRolesFile);
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
    const subsDir = fs.readdirSync(process.cwd())
    let fls = filterByRegx(subsDir,/(?<!\.Dub)\.ass$/);
    if(fls.length<1){
        console.log(`[ERROR] No input files!`);
        await doPause();
        process.exit();
    }
    for(let f of fls){
        file = f;
        if(fs.existsSync(file)){
            console.log(`[INFO] Processing ${file}...`);
            try{
                await parseFile();
            }
            catch(e){
                console.log(e);
            }
            console.log(`[INFO] Done!`);
        }
    }
    await doPause();
    process.exit();
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
                    ass.styles.format = parm.split(',').map(p=>p.trim())
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
                    ass.events.format = parm.split(',').map(p=>p.trim())
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
                    if(typeof current.Name != 'string'){
                        current.Name = '';
                    }
                    if(current.Name == '' && typeof current.Actor == 'string'){
                        current.Name = current.Actor;
                    }
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
        return;
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
                fs.writeFileSync(`${fileName}.Roles.txt`, `\ufeff`+txtRole.join(`\r\n`)); 
        }
        // make new ass, srt and docx
        let assFile = '';
        assFile = [
            `[Script Info]`,
            `Title: ${ass.script_info.Title?ass.script_info.Title:''}`,
            `Original Translation: `,
            `Original Editing: `,
            `Original Timing: `,
            `Synch Point: `,
            `Script Updated By: `,
            `Update Details: `,
            `ScriptType: v4.00+`,
            `Collisions: Normal`,
            (ass.script_info.PlayResX?`PlayResX: ${ass.script_info.PlayResX}`:''),
            (ass.script_info.PlayResY?`PlayResX: ${ass.script_info.PlayResY}`:''),
            `Timer: 0.0000`,
            (ass.script_info.WrapStyle?`PlayResX: ${ass.script_info.WrapStyle}`:''),
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
                let cleanPrevDlDocx = docArr[current_row].text;
                let stringTimeMatch = `(?:\\()\\d{1,2}(-|:)\\d{2}(?:\\))`;
                let strStartReplace  = new RegExp(`^(|\\/+)(| +)${stringTimeMatch}`);
                let strEndReplace    = new RegExp(`(\\/+)(| +)${stringTimeMatch}$`);
                if(cleanPrevDlDocx.slice(-2) == '//'
                    || cleanDialogDocx.slice(0, 2) == '//'
                    || cleanDialogDocx.match(strStartReplace)
                    || cleanPrevDlDocx.match(strEndReplace)
                    || assTimeToDoc(dlgc.Start, dlgp.End) == 5 && cleanPrevDlDocx.slice(-1) != '/'
                ){
                    if(cleanDialogDocx.match(strStartReplace)){
                        cleanDialogDocx = cleanDialogDocx.replace(strStartReplace,'').trim();
                    }
                    if(cleanPrevDlDocx.match(strEndReplace)){
                        cleanPrevDlDocx = cleanPrevDlDocx.replace(strEndReplace,'').trim();
                    }
                    if(cleanPrevDlDocx.slice(-2) == '//'){
                        cleanPrevDlDocx = cleanPrevDlDocx.replace(/\/+$/,'').trim();
                    }
                    docArr[current_row].text = cleanPrevDlDocx
                        + (config.use_docx_string_splitter ? docxStringSplitter : ' ')
                        + '//' + assTimeToDoc(dlgc.Start);
                }
                if(cleanPrevDlDocx.slice(-1) == '/'
                    || assTimeToDoc(dlgc.Start, dlgp.End) == 1 && cleanPrevDlDocx.slice(-1) != '/'
                ){
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
            fs.writeFileSync(`${fileName}.Dub.ass`, `\ufeff`+assFile);
        }
        if(!config.skip_create_srt_mod){
            fs.writeFileSync(`${fileName}.Dub.srt`, `\ufeff`+srtFile);
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

function addTableCell(content, size){
    let cellContent = [];
    content = content.split(docxStringSplitter);
    for(let c of content){
        cellContent.push(new Paragraph({ text: c }));
    }
    let cell = new TableCell({
        children: cellContent,
        verticalAlign: VerticalAlign.CENTER,
        width: { size: size },
    });
    cell.properties.setWidth(size);
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
        if(time[2] > 59){
            time[2] = 0;
            time[1]++;
        }
        if(time[1] > 59){
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
        time1 = time1[0] * 60 * 60 + time1[1] * 60 + time1[2];
        let time2 = strToTimeArr(timePrev);
        time2 = time2[0] * 60 * 60 + time2[1] * 60 + time2[2];
        if(time1 - time2 > config.docx_join_time_long - 0.001){
            return 5;
        }
        if(time1 - time2 > config.docx_join_time_short - 0.001){
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

async function doPause(){
    console.log(`Press any key to continue...`);
    process.stdin.setRawMode(true);
    return new Promise(resolve => process.stdin.once('data', data => {
        const byteArray = [...data];
        if (byteArray.length > 0 && byteArray[0] === 3) {
            console.log('^C');
            process.exit(1);
        }
        process.stdin.setRawMode(false);
        resolve();
    }));
}
