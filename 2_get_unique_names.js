const fs = require('fs');
const path = require("path");
if (!process.argv[2]){
    process.exit(1)
}
const inputFile = process.argv[2]; //console.log(inputFile)
//const inputPath = process.argv[3]
let i = 0;
let id = 0;
let readout_id = 0;
let label_id = 0;
let line_id = 0;
let all_obj = [];
let all_conn = [];
let all_port = [];
// cbrxhemi cbrxglbh cbrxeucb cbehcrtw
let global_x_offset = 7; // -58, -115, -170, 7
let global_y_offset = -20; // 65, -40, 30, -20
let formula = '';
let readOutLcamp =[];
let readOutEpc =[];
let readOutTwta =[];
let scid;
let p;
let nameArray = [];

class Pair {
    constructor(s1, s2,p, initial_side,p2) {
        this.s1 = s1;
        this.s2 = s2;
        this.p =p;
        this.initial_side = initial_side;
        this.p2 =p2;
        this.ending_side = '';
    }

    isEqual(p) {
        if (this.s1 == p.s1 && this.s2 == p.s2) {
            return true;
        } else if (this.s2 == p.s1 && this.s1 == p.s2) {
            return true;
        }
        return false;
    }

    toString() {
        return '[(' + this.s1 + ',' + this.s2 + ') ' +this.p+' '+ this.initial_side + ' ' +this.p2+' '+ this.ending_side + ']';
    }

    [Symbol.for('nodejs.util.inspect.custom')]() {
        return this.toString();
    }
}

class Shape {
    constructor(name, x, y, h, w, p2, right = false, left = false, top = false, bottom = false, channels = 0) {
        this.name = name;
        this.x1 = x;
        this.y1 = y;
        this.x2 = x + w;
        this.y2 = y - h;
        this.height = h;
        this.width = w;
        this.p2 = p2
        this.right_multi = right;
        this.left_multi = left;
        this.top_multi = top;
        this.bottom_multi = bottom;
        this.channels = channels;
    }

    inBounds(x, y) {
        if (x >= this.x1 && x <= this.x2) {
            if (y <= this.y1 && y >= this.y2) {
                return true;
            }
        }
        return false;
    }

    toString() {
        return '{' + this.name + ': (' + this.x1 + ',' + this.y1 + ') (' + this.x2 + ',' + this.y2 + ')} ' + this.channels + this.left_multi + this.right_multi;
    }

    toLocaleString() {
        return '{' + this.name + ': (' + this.x1 + ',' + this.y1 + ') (' + this.x2 + ',' + this.y2 + ')} ' + this.channels + this.left_multi + this.right_multi;
    }

    valueOf() {
        return '{' + this.name + ': (' + this.x1 + ',' + this.y1 + ') (' + this.x2 + ',' + this.y2 + ')} ' + this.channels + this.left_multi + this.right_multi;
    }
}

class Connection {
    constructor(x1=0, y1=0, x2=0, y2=0, initial_side, ending_direction, origin = 0, p) {
        this.x1 = x1;
        this.y1 = y1;
        this.x2 = x2;
        this.y2 = y2;
        this.initial_side = initial_side;
        //this.ending_side = '';
        this.ending_direction = ending_direction;
        this.origin = origin;
        this.p =p;
    }

    toString() {
        return '[(' + this.x1 + ',' + this.y1 + ') (' + this.x2 + ',' + this.y2 + ',' + this.ending_direction +this.p+ ')]';
    }

    toLocaleString() {
        return '[(' + this.x1 + ',' + this.y1 + ') (' + this.x2 + ',' + this.y2 + ',' + this.ending_direction +this.p+ ')]';
    }

    valueOf() {
        return '[(' + this.x1 + ',' + this.y1 + ') (' + this.x2 + ',' + this.y2 + ',' + this.ending_direction +this.p+ ')]';
    }
}

const populateProps = (collection,groupsec) =>{ //console.log(line_arr)
    let twtaNum,readout,Mnemonic,formula;
    let arr =[];
    arr.push(collection)
    arr.forEach(el => {
        if(groupsec === "CAMP"){
            twtaNum = el[0].split(",")[2].trim().replace("\\","");
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic = el[k].split(",")[1].trim()//.split("5130"); console.log(Mnemonic)//
                formula = el[k].split(",")[1].trim().replace("\\","");
                readOutLcamp.push( {
                    "group": groupsec,
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
        }
        else   if(groupsec === "BEACONS"){
            let twta =  el[0].split(",")[1].split("-")
            twtaNum = twta[twta.length -1]; //console.log(twtaNum)
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic = el[k].split(",")[1].trim(); //console.log(Mnemonic)//
                formula = el[k].split(",")[1].trim(); //console.log(formula)
                readOutLcamp.push( {
                    "group": el[0].split(",")[1].trim(),
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
        }
        else if (groupsec === "EPC"){
            twtaNum = el[0].split(",")[1].split("-")[4];
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic = el[k].split(",")[1].trim(); //console.log(Mnemonic)//
                formula = el[k].split(",")[1];
                readOutLcamp.push({
                    "group": groupsec,
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
        }
        else if(groupsec === "TWTA") {
            twtaNum = el[0].split(",")[2];
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic = el[k].split(",")[1].trim(); //console.log(Mnemonic)//
                formula = el[k].split(",")[1];
                readOutLcamp.push({
                    "group": groupsec,
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
        }
        else if(groupsec === "RCVR") {
            let twta =  el[0].split(",")[1].split("-")
            twtaNum = twta[twta.length -1]; //console.log(twtaNum)//console.log(el[0].split(",")[1].trim())//
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                ///if (el[k].startsWith("ON-OFF-STATUS")) {//console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic = el[k].split(",")[1]; //console.log(Mnemonic)//
                formula = el[k].split(",")[1];  //console.log("group: "+el[0].split(",")[1].split("-")[1].slice(0,4))//console.log(formula)
                readOutLcamp.push({
                    "group": el[0].split(",")[1].trim(),
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
            //}
        }
        else if(groupsec === "TCR-XMTR") {
            twtaNum = el[0].split(",")[1].split("-")[2];
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic = el[k].split(",")[1]+"edge";  //console.log(readout+" : "+Mnemonic)//console.log(el[k].split(",")[1].split(scid.toString()+"-")[1])//
                formula = el[k].split(",")[1]; //console.log(formula)
                readOutLcamp.push({
                    "group": el[0].split(",")[1],
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
        }
        else if(groupsec === "TCR-CMDRX") {
            twtaNum = el[0].split(",")[1].split("-")[2];
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic = el[k].split(",")[1]+"edge"; //console.log(readout+" : "+Mnemonic)//.split(scid.toString()+"-")[1].split("=")[0])//
                formula = el[k].split(",")[1];
                readOutLcamp.push({
                    "group": el[0].split(",")[1],
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
        }
        else if(groupsec === "TCR-BBE") {
            twtaNum = el[0].split(",")[1].split("-")[2];
            for (let k = 1; k < el.length; k++) { //console.log(el[k])
                readout = el[k].split(",")[0]; //console.log(readout)
                Mnemonic =el[k].split(",")[1]+"edge"; //console.log(Mnemonic)//
                formula = el[k].split(",")[1];
                readOutLcamp.push({
                    "group": el[0].split(",")[1],
                    "twtaNum": twtaNum,
                    "readout": readout,
                    "Mnemonic": Mnemonic,
                    "formula": formula
                })
            }
        }
        else  if (collection.includes("SMALL-G22X-READOUTS")) {
            readOutLcamp.push({
                "group": "G22X-READOUTS",
                "twtaNum": "G22X-READOUTS",
                "readout": "G22X-READOUTS",
                "Mnemonic": line_arr[6] + "edge",
                "formula": line_arr[6]
            })
        }
    })
}

function process_label(label_arr) { //console.log(label_arr.split(','));
    let line = label_arr.split(',');
    let x_loc;
    let y_loc;
    let width;
    let height;
    let h_width;
    let h_height;

    if (line.length <= 6) {
        x_loc = parseInt(line[1]) + 10;
        y_loc = -1 * parseInt(line[2]);
        width = parseInt(line[3]);
        height = parseInt(line[4]);
        h_width = width / 2;
        h_height = height / 2;
        x_loc -= h_width;
        y_loc -= h_height;
        let text;
        text = line[5].replace(';', '').replace('"', '').replace('"', ''); //console.log(text)
        svg_elm = '\n<g id="label' + label_id + '" v:mID="' + id + '" v:groupContext="shape">\n';
        svg_elm += '<rect x="' + x_loc + '" y="' + (y_loc + 5) + '" width="' + (width / 2) + '" height="' + (height / 2) + '" opacity="0.0"/>';
        svg_elm += '<text x="0" y="0" v:langID="1033" font-size="10" fill="black">' + text.replace(":", "");
        label_id++;
        id++;
        svg_elm += '</text>\n';
        svg_elm += '</g>';
    }
    //console.log(svg_elm)
    else if (line.length > 6 && line[0] === "REALLY-SMALL-HOT-BUTTONS-MSG" || line[0].includes("HOT-BUTTONS")) {
        let props;
        x_loc = parseInt(line[2]) + 10;
        y_loc = -1 * parseInt(line[3]);
        width = parseInt(line[4]);
        height = parseInt(line[5]);
        h_width = width / 2;
        h_height = height / 2;
        svg_elm = '\n<g id="HOT_BUTTON' + label_id + '" v:mID="' + id + '" v:groupContext="shape">\n';
        svg_elm += '<rect x="' + x_loc + '" y="' + (y_loc + 5) + '" width="' + (width / 2) + '" height="' + (height / 2) + '" opacity="0.0"/>';
        svg_elm += ('\n<text x=" ' + parseInt(x_loc + (width / 2)) + '" y="' + (y_loc - 1) + '" font-size="' + '10' + '" fill="' + 'grey'+  '">' + line[1].toUpperCase().slice(6) + '</text>');
        svg_elm += ('\n<v:custProps>');
        svg_elm += ('\n<v:cp v:nameU="Link" v:lbl="Link" v:type="0" v:langID="1033" v:val="VT4(' + line[1].replace(';', '').toUpperCase().replaceAll('-', '_') + ')" />');
        svg_elm += ('\n<v:cp v:nameU="Trace" v:lbl="Trace" v:type="0" v:langID="1033" v:val="VT4()" />');
        svg_elm += ('\n</v:custProps>');
        label_id++;
        id++;
        //svg_elm += '</text>\n';
        svg_elm += '</g>';
    }//console.log(svg_elm)
    return svg_elm;
}

function process_object(line_arr) { //console.log(line_arr)
    readout_id += 1;
    label_id += 1;
    id += 1;
    let line = line_arr[0].split(', ');
    let classname = line[0]; //console.log(classname)
    let tName = line[1].split('-')[2]; //console.log(line[1])
    let twtaName = line[1].split('-')[line[1].split('-').length-2]; //console.log(twtaName)
    //objectNameList += classname +"\n";
    let group_name;
    if (classname.length > 25) {
        if(classname.includes("OMUX")){
            classname = classname.slice(0, 15)+"_OMUX";
        }
        else  if(classname.includes("CMUX")){
            classname = classname.slice(0, 15)+"_CMUX";
        }
        else if(classname.includes("IMUX")){
            classname = classname.slice(0, 15)+"_IMUX"; //console.log(classname)
        }
        else{
            classname = classname.slice(0, 19);
        }
    }
    if(classname.includes('BEACONS-HYBRID-MUX-TYPE-2')){
        classname = classname.replace("BEACONS-","").replace("-MUX","");
    }
    //console.log("class = "+classname);
    // if(classname.includes('TWTA')&& !classname.includes('EPC')&&  !classname.includes('IS10')&&  !classname.includes('IS9')){
    //     group_name = line[1].split('-')[2]+"T1_TWTA"+ '_' + line[1].split('-')[2];
    // }
    if (classname.includes('IS') && classname.match(/IS\d+-TWTA(?!-EPC)/)){ //console.log(line[1].split('-')[2])
        let name = line[1].split('-'); //console.log(name)
        if (name.some(part => part.includes('IS'))) {
            group_name = name[name.length - 2] + "_TWTA" + '_' + name[name.length - 2]; //console.log(group_name)
        }
        else{
            group_name = name[name.length - 2].slice(-2) + "_TWTA" + '_' + name[name.length - 2]; //console.log(group_name)
        }
    }
    else if(classname.includes('IS') && classname.match(/IS\d+.*EPC/)||classname.includes ("T8-EPC-DUAL")){ //console.log("class = "+classname);//console.log(line[1].split('-')[4])
        let name = line[1].split('-')[line[1].split('-').length -2]//.replace(/^\d+/, '')
        group_name = "E1_EPC"+ '_' + name; //console.log(group_name)
    }
    else if(classname.includes('IS') && (classname.match(/IS\d+-OPA-CAMP/)||classname.match(/IS\d+-CAMP/))){ //console.log("class = "+classname);//console.log(line[1].split('-')[4])
        let name = line[1].split('-')[line[1].split('-').length -2]//.replace(/^\d+/, '')
        group_name = name.slice(-2)+"_LCAMP"+ '_' + name; //console.log(group_name)
    }
    else if(classname.includes('SELECT')){ //console.log("class = "+classname);//console.log(line[1].split('-')[4])
        group_name = classname.replace("SELECT", "V")+ '_' + label_id; //console.log(group_name)
    }
     else {
        group_name = classname + '_' + label_id; //console.log(group_name)
    }
    //console.log(group_name);
    let lineInfo = group_name +"|"+ line_arr[0]; //console.log(line_arr[0])
    nameArray.push(lineInfo)
    // if (classname.length > 20) {
    //     group_name = classname.slice(0, 19) + '_' + label_id;
    // }
    // else {
    //     group_name = classname + '_' + label_id;
    // }
    if (classname === 'WIN-GROW-BTN' || classname === 'WIN-DISPLAY-HIDE-BTN'||   classname ==='WIN-SHRINK-BTN') {
        return '';
    }
    let objname = line[1];
    let alias = line[2];
    let x_loc = parseInt(line[3]);
    let y_loc = -1 * parseInt(line[4]);
    let width = parseInt(line[5]);
    let height = parseInt(line[6]);
    let h_width = width / 2;
    let h_height = height / 2;
    let att_count = parseInt(line[7].substring(5));
    let ch_count = parseInt(line[8].substring(4, line[8].length - 1)); //console.log(ch_count)
    let style = 'fill="wheat"';
    //let svg_elm = '';
    let extra = '';
    let props = '';
    let offset_x_axis = 0;
    let offset_y_axis = 0;
    let channels = 0;
    let shape_obj;
    let right_multi = false;
    let left_multi = false;
    let top_multi = false;
    let bottom_multi = false;
    let text_size = (width < 20) ? "8" : "10";
    let port =[];
    let c,objN;
    let currentDate = new Date();
    let dateString = currentDate.toString().split(" ").slice(0,5).join('-');
    for (let i = 0; i < line_arr.length; i++) {
        if(line_arr[i].startsWith('P')||line_arr[i].startsWith('IN')||line_arr[i].startsWith('OUT')||line_arr[i].startsWith('NONE')) {
            let n = line_arr[i].split(', ').length - 5;
            if (line_arr[i][0].split(', ') !== "POSITION") {
                objN= line_arr[i].split(', ')[0]
                port = line_arr[i].split(', ').slice(-n)
                if (line_arr[i].split(', ')[0]!== "POSITION"){
                    p= line_arr[i].split(', ')[0]
                }
                //console.log(p)
            }
        }
    }
    if (classname.includes('LORAL-MINI-IMUX-DUAL-MODE') || classname.includes('IS32-KU-IMUX-12CH-ODD')) {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
    }
    else if (classname.includes('HOT-ARROW')) {
        //style = 'fill="grey"';
    }
    else if ( classname === 'CHANNEL-STUB') {
        style = 'fill="grey"';
    }
    else if (classname === 'CONFIG-CHANNEL-OBJECT') {
        style = 'fill="none" stroke="black"';
        svg_elm = '\n<path d="M ' + x_loc + ' ' + (y_loc + h_height) + ' L ' + (x_loc + h_width) + ' ' + y_loc + ' L ' + x_loc + ' ' + (y_loc - h_height) + ' L ' + x_loc + ' ' + (y_loc + h_height) + '" fill="Balck" stroke="red"/>';
        svg_elm += '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + y_loc + '" x2="' + (x_loc + h_width) + '" y2="' + y_loc + '" stroke="Black" stroke-width="0.75"/>';
    }
    else if (classname === 'CHANNEL-POST') {
        style = 'fill="none" ';
        extra += '\n<path d="M ' + x_loc + ' ' + (y_loc + h_height) + ' L ' + (x_loc + h_width) + ' ' + y_loc + ' L ' + x_loc + ' ' + (y_loc - h_height) + ' L ' + x_loc + ' ' + (y_loc + h_height) + '" fill="black" stroke="black"/>';
        extra += `\n<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}" stroke="black" stroke-width="0.75"/>`;
    }
    else if (classname === 'VERTICAL-BORDER-LINE') {
        style = 'fill="none" stroke="Black"';
    }
    else if (classname === 'HORIZONTAL-BORDER-LINE') {
        style = 'fill="none" stroke="Black"';
    }
    else if (classname.includes('OPA-CAMP')||classname.includes('LCAMP-TWTA')||classname.includes('LCHAMP')||classname.includes('CHAMP')||classname.includes('CAMP')) {
        populateProps(line_arr, "LCAMP");
        style = 'fill="none" stroke="green" stroke-width="2"';
        right_multi = true;
        left_multi = true; //console.log(twtaName);
        extra = '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + y_loc +
            '" x2="' + (x_loc + h_width) + '" y2="' + y_loc +
            '" stroke="black" stroke-width="0.75"/>';
        extra += '<title>' + twtaName.slice(-2) + '_LCAMP'+'_' +twtaName + '</title>\n';
    }
    else if (classname.includes('EPC')) { // second twta
        populateProps(line_arr,"EPC");
        style = 'fill="none" stroke="black" stroke-width="2"';
        right_multi = true;
        left_multi = true; //console.log(twtaName);
        channels = 2;
        extra = '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + ((y_loc - h_height) + (height / 3)) +'" x2="' + (x_loc + h_width) + '" y2="' + ((y_loc - h_height) + (height / 3)) +'" stroke="black" stroke-width="0.75"/>';
        extra += '\n\n\n\n\n<line x1="' + (x_loc - h_width) + '" y1 ="' + ((y_loc - h_height) + ((height / 3) * 2)) +'" x2="' + (x_loc + h_width) + '" y2="' + ((y_loc - h_height) + ((height / 3) * 2)) +'" stroke="black" stroke-width="0.75"/>';
        extra += '<title>'  + 'E1_EPC'+'_' +twtaName + '</title>\n';
    }
    else if (classname.includes('TWTA')&& !classname.includes('EPC')) {
        populateProps(line_arr,"TWTA");
        style = 'fill="none" stroke="green" stroke-width="2"';
        right_multi = true;
        left_multi = true; //console.log(twtaName);
        extra = '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + y_loc +'" x2="' + (x_loc + h_width) + '" y2="' + y_loc +'" stroke="black" stroke-width="0.75"/>\n';
        extra += '<title>' + twtaName.slice(-2) + '_TWTA'+'_' +twtaName + '</title>';
        //extra +='</g>\n';
        // for(let i =1; i<line_arr.length; i++){ console.log(line_arr[i])
        //     if(line_arr[i].split(',')[0] !== "IN1" && line_arr[i].split(',')[0] !== "OUT1") {
        //         rout = line_arr[i].split(',')[0];console.log(rout)
        //         itn = line_arr[i].split(",")[1]; //console.log(Mnemonic)//
        //         fm = line_arr[i].split(",")[1];
        //         let x_loc = parseInt(line[3]);
        //         let y_loc = -1 * parseInt(line[4]);
        //         let width = parseInt(line[5]);
        //         let height = parseInt(line[6]);
        //         let h_width = width / 2;
        //         let h_height = height / 2;
        //         extra += '<g id="Readout.' + rout +"-" + readout_id + '" v:mID="' + id + '" v:groupContext="shape">\n';
        //         extra += '<rect x="' + x_loc + '" y="' + y_loc + '" width="' + width / 2 + '" height="' + height / 2 + '" fill="'+'#fff2cc'+'"/>\n';
        //         extra += '<title>'+'Readout.' + rout +"-" + readout_id + '</title>';
        //         extra +='\n<v:custProps>'
        //         extra += '\n<v:cp v:nameU="mnemonic" v:lbl="mnemonic" v:type="0" v:langID="1033" v:val="VT4( )" />';
        //         extra += '\n<v:cp v:nameU="Formula" v:lbl="Formula" v:type="0" v:langID="1033" v:val="VT4(  )" />\n';
        //         extra +='</v:custProps>\n';
        //         //extra +='</g>\n';
        //     }
        //     //label_id +=1;
        //
        //     id +=1;
        // }
        // readout_id += 1;
        //svg_elm += extra;
    }
    else if (classname.includes('TCR-XMTR')) {
        populateProps(line_arr,"TCR-XMTR");
        style = 'fill="none" stroke="green" stroke-width="2"';
        right_multi = true;
        left_multi = true;
        extra = '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + y_loc +'" x2="' + (x_loc + h_width) + '" y2="' + y_loc +'" stroke="black" stroke-width="0.75"/>';
    }
    else if (classname.includes('TCR-CMDRX')) {
        populateProps(line_arr,"TCR-CMDRX");
        style = 'fill="none" stroke="green" stroke-width="2"';
        right_multi = true;
        left_multi = true;
        extra = '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + y_loc +'" x2="' + (x_loc + h_width) + '" y2="' + y_loc +'" stroke="black" stroke-width="0.75"/>';
    }
    else if (classname.includes('TCR-BBE')) {
        populateProps(line_arr,"TCR-BBE");
        style = 'fill="none" stroke="green" stroke-width="2"';
        right_multi = true;
        left_multi = true;
        extra = '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + y_loc +'" x2="' + (x_loc + h_width) + '" y2="' + y_loc +'" stroke="black" stroke-width="0.75"/>';
    }
    else if (classname.includes('OMUX')||classname.includes('CMUX')) {
        left_multi = true;
        style = 'fill="none" stroke="black" stroke-width="0.75"';
    }
    else if (classname.includes('IMUX')) {
        right_multi = true;
        style = 'fill="none" stroke="black" stroke-width="0.75"';
    }
    else if (classname.includes('TX-POL-ANTENNA') || classname.includes('ANTENNA') && !classname.includes('OMNI') ) {
        style = 'fill="none"';
        top_multi = true;
        bottom_multi = true;
        channels = 2;
        extra = `\n<path d="M ${x_loc - h_width} ${y_loc + h_height} L ${x_loc - h_width + (width / 4)} ${y_loc + h_height - (height / 4)} L ${x_loc + h_width} ${y_loc + h_height - (height / 4)} L ${x_loc + h_width} ${y_loc + h_height - ((height / 4) * 3)} L ${x_loc - h_width + (width / 4)} ${y_loc + h_height - ((height / 4) * 3)} L ${x_loc - h_width} ${y_loc - h_height} L ${x_loc - h_width} ${y_loc + h_height}" fill="none" stroke="black"/>`;
    }
    else if (classname === "EPIC-ANTENNA") {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        extra += `<rect x="${x_loc - h_width}" y="${y_loc - h_height}" width="${width}" height="${height}" ${style} />`;
        extra += `<circle cx="${x_loc}" cy="${y_loc - h_height}" r="${width / 4}" ${style} />`;
    }
    else if (classname === "EPIC-ROUTER") {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        extra += `<rect x="${x_loc - h_width}" y="${y_loc - h_height}" width="${width}" height="${height}" ${style} />`;
        extra += `<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}"  stroke="black" stroke-width="0.75" />`;
        extra += `<line x1="${x_loc}" y1="${y_loc - h_height}" x2="${x_loc}" y2="${y_loc + h_height}" stroke="black" stroke-width="0.75" />`;
    }
    else if (classname.includes('COMM') || classname.includes('COMM-DUAL-IN-VH-POL-ANTENNA') || classname.includes('VH-POL-ANTENNA')  || classname.includes('COMM-QUAD-IN-POL')) {
        //style += 'fill="none"';
        extra = `\n<path d="M ${x_loc - h_width} ${y_loc + h_height} L ${x_loc - h_width + (width / 4)} ${y_loc + h_height - (height / 4)} L ${x_loc + h_width} ${y_loc + h_height - (height / 4)} L ${x_loc + h_width} ${y_loc + h_height - ((height / 4) * 3)} L ${x_loc - h_width + (width / 4)} ${y_loc + h_height - ((height / 4) * 3)} L ${x_loc - h_width} ${y_loc - h_height} L ${x_loc - h_width} ${y_loc + h_height}" fill="none" stroke="black"/>`;
    }
    else if (classname.includes('SPLITTER') || classname.includes('COUPLER')|| classname.includes('DIPLEXER-COMBINER')|| classname.includes('EPIC-FILTER-SPLITTER') ||classname.includes('BOEING-HYBRID-MUX')|| classname.includes('IS9-TRIPLE-BFN-COUPLER')) {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        if (classname.includes('COUPLER')) {
            channels = 2;
            left_multi = true;
        }
        else if (classname.includes('DUAL-BFN')) {
            channels = 2;
            right_multi = true;
        }
        else if (classname.includes('TRIPLE-BFN')) {
            channels = 3;
            right_multi = true;
        }
        else if (classname.includes('QUAD-BFN')) {
            channels = 4;
            right_multi = true;
        }
        else if (classname.includes('FIVE-BFN')) {
            channels = 5
            right_multi = true;
        }
        if (classname.includes( 'FIL-UNIT')) {
            style = 'fill="none" stroke="black" stroke-width="0.75"';
            channels = 2;
            right_multi = true;
        }
        else if (classname.includes('DUAL-BFN')) {
            channels = 2;
            right_multi = true;
        }
        else if (classname.includes('TRI-BFN')) {
            channels = 3;
            right_multi = true;
        }
        else if (classname.includes('QUAD-BFN')) {
            channels = 4;
            right_multi = true;
        }
        else if (classname.includes('SIX-BFN')) {
            channels = 6;
            right_multi = true;
        }
        else if (classname.includes('DUAL-BFN-COUPLER')) {
            channels = 2;
            left_multi = true;
        }
        else if (classname.includes('HYBRID-MUX')) {
            channels = 2;
            left_multi = true;
            right_multi = true;
        }
    }
    else if (classname === 'EPIC-CHANNELIZER-BLOCK') {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        channels = 5;
        right_multi = true;
        left_multi = true;
        for (let y = 1; y < 5; y++) {
            extra += `\n<g id="object${label_id}" v:mID="${id}" v:groupContext="shape">`;
            extra += `\n<rect x="${x_loc - h_width + 5}" y="${y_loc - h_height + (y * (height / 5)) - 13}" height="15" width="17" fill="none"/>`;
            extra += `\n<text x="${(x_loc - h_width + 5) / 2}" y="${(y_loc - h_height + (y * (height / 5)) - 13) / 2}" font-size="12">rx${y - 1}</text>`;
            extra += '\n</g>';
        }
        for (let y = 1; y < 5; y++) {
            extra += `\n<g id="object${label_id}" v:mID="${id}" v:groupContext="shape">`;
            extra += `\n<rect x="${x_loc + h_width - 19}" y="${y_loc - h_height + (y * (height / 5)) - 13}" height="15" width="17" fill="none"/>`;
            extra += `\n<text x="${x_loc + h_width - 19}" y="${y_loc - h_height + (y * (height / 5)) - 13}" font-size="12">tx${y - 1}</text>`;
            extra += '\n</g>';
        }
    }
    else if (classname.includes('JUNCTION') || classname === 'EPIC-CHANNELIZER-PORT') {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
    }
    else if (classname === 'TCR-UNIT') {
        style = 'fill="none" stroke="green" stroke-width="0.75"';
        channels = 2;
        right_multi = true;
    }
    else if (classname.includes('C-SWITCH')) {
        style = 'fill="wheat" stroke="Black"';
        extra += `<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}" stroke="black" stroke-width="0.75"/>`;
        extra += `\n<line x1="${x_loc}" y1="${y_loc + h_height}" x2="${x_loc}" y2="${y_loc - h_height}" stroke="black" stroke-width="0.75"/>`;
    }
    else if (classname.includes('IS32-TCR-TOGGLE')) {
        style = 'fill="wheat" stroke="Black"';
        // extra += `<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}" stroke="black" stroke-width="0.75"/>`;
        // extra += `<line x1="${x_loc}" y1="${y_loc + h_height}" x2="${x_loc}" y2="${y_loc - h_height}" stroke="black" stroke-width="0.75"/>`;
    }
    else if (classname.includes('DOWN-CONVERTER') || classname.includes('LNA')|| classname.includes('T8-KA-DC-CONVERTER')) {
        populateProps(line_arr,"RCVR")
        style = 'fill="none" stroke="green" stroke-width="2"';
    }
    else if (classname.includes('T-SWITCH')) {
        style = 'fill="wheat" stroke="black"';
        extra += `\n<g id="object${label_id}" v:mID="${id}" v:groupContext="shape">`;
        if (width > 20) {
            extra += `\n<rect x="${x_loc - h_width + 2}" y="${y_loc + 1}" height="${height - 15}" width="${width / 2}" fill="none"/>`;
            extra += `\n<text x="${x_loc - h_width + 2}" y="${y_loc + 1}" font-size="${text_size}">T</text>`;
        }
        else {
            extra += `\n<rect x="${x_loc - h_width + 2}" y="${y_loc + 4}" height="${height - 15}" width="${width / 2}" fill="none"/>`;
            //extra += `\n<text x="${x_loc - h_width + 2}" y="${y_loc + 4}" font-size="${text_size}">T</text>`;
        }
        extra += `\n<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}" stroke="black" stroke-width="0.75"/>`;
        extra += `\n<line x1="${x_loc}" y1="${y_loc + h_height}" x2="${x_loc}" y2="${y_loc - h_height}" stroke="black" stroke-width="0.75"/>`;
        extra += "\n" + "</g>";
    }
    else if (classname.includes('S-SWITCH')) {
        style = 'fill="wheat" stroke="Black"';
        extra += `\n<g id="object${label_id}" v:mID="${id}" v:groupContext="shape">\n`;
        if (width > 20) {
            extra += `\n<rect x="${x_loc}" y="${y_loc - height + 5}" height="${height}" width="${width / 2}" fill="none"/>`;
            extra += `\n<text x="${x_loc}" y="${y_loc - height + 5}" font-size="${text_size}">S</text>`;
        } else {
            extra += `\n<rect x="${x_loc}" y="${y_loc - height + 5}" height="${height}" width="${width / 2}" fill="none"/>`;
            extra += `\n<text x="${x_loc}" y="${y_loc - height + 5}" font-size="${text_size}">S</text>`;
        }

        extra += `\n<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}" stroke="black" stroke-width="0.75"/>`;
        extra += "\n" + "</g>";
    }
    else if (classname.includes('R-SWITCH')) { //console.log(classname)
        style = 'fill="wheat" stroke="Black"';
        extra += `\n<g id="object${label_id}" v:mID="${id}" v:groupContext="shape">`;
        if (width > 20) {
            extra += `\n<rect x="${x_loc - h_width + 2}" y="${y_loc + 1}" height="${height - 15}" width="${width/2}" fill="none"/>`;
            extra += `\n<text x="${x_loc - h_width + 2}" y="${y_loc + 1}" font-size="${text_size}">R</text>`
        }else {
            extra += `\n<rect x="${x_loc}" y="${y_loc - height + 5}" height="${height - 15}" width="${width / 2}" fill="none"/>`;
            extra += `\n<text x="${x_loc}" y="${y_loc - height + 5}" font-size="${text_size}">S</text>`;
        }

        extra += `\n<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}" stroke="black" stroke-width="0.75"/>`;
        extra += `\n<line x1=" ${x_loc}" y1=" ${y_loc - h_height } " x2=" ${x_loc} " y2="  ${y_loc - h_height }" stroke="black" stroke-width="0.75"/>`
        extra += "\n" + "</g>";
    }
    else if (classname.includes ("RECEIVER")) {
        populateProps(line_arr, "RCVR");
        style = 'fill="none" stroke="green" stroke-width="0.75"';
        extra = '\n<line x1="' + (x_loc - h_width) + '" y1 ="' + y_loc +
            '" x2="' + (x_loc + h_width) + '" y2="' + y_loc + '" stroke="black" stroke-width="0.75"/>';
    }
    else if (classname === "GND" ||classname ==="REPEATER-LOAD") {
        style = 'fill="none"';
        if (height > width) {
            extra = `\n<path d="M ${x_loc + h_width} ${y_loc + h_height} L ${x_loc - h_width} ${y_loc + h_height - (height / 4)} L ${x_loc + h_width} ${y_loc + h_height - (2 * (height / 4))} L ${x_loc - h_width} ${y_loc + h_height - (3 * (height / 4))} L ${x_loc + h_width} ${y_loc - h_height}" fill="none" stroke="black"/>`;
        } else {
            extra = `\n<path d="M ${x_loc - h_width} ${y_loc + h_height} L ${x_loc - h_width + (width / 4)} ${y_loc - h_height} L ${x_loc - h_width + ((width / 4) * 2)} ${y_loc + h_height} L ${x_loc - h_width + ((width / 4) * 3)} ${y_loc - h_height} L ${x_loc + h_width} ${y_loc + h_height}" fill="none" stroke="black"/>`;
        }
    }
    else if (classname === "BEACONS"|| classname === "BEACONS-TYPE-2" ||  classname ==="XMITRS") {
        populateProps(line_arr,"BEACONS")
        style = 'fill="none" stroke="green" stroke-width="0.75"';
    }
    else if (classname.includes("BEACONS-HYBRID-MUX") ||classname.includes('HYBRID-TYPE-2')) {
        channels = 2;
        left_multi = true;
        right_multi = true;
        style = 'fill="none" stroke="black" stroke-width="0.75"';
    }
    else if (classname === "BOEING-EPIC-FWD-FILTER-HIDDEN" || classname === "SHOW-ADCS-INFO-BUTTON") {
        style = 'fill="none"';
    }
    else if (classname.includes('BOEING-EPIC-FWD-FILTER-36MHZ') || classname === 'BOEING-EPIC-DL-FILTER-IS33-36MHZ') {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        extra = `<rect x="${x_loc - h_width}" y="${y_loc - h_height}" width="${width / 6}" height="${height}" fill="lightblue"/>`;
    }
    else if (classname.includes('BOEING-EPIC-FWD-FILTER-184MHZ') || classname === 'BOEING-EPIC-DL-FILTER-187MHZ') {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        extra = `<rect x="${x_loc - h_width}" y="${y_loc - h_height}" width="${width / 6}" height="${height}" fill="mediumslateblue"/>`;
    }
    else if (classname.includes('BOEING-EPIC-FWD-FILTER-125MHZ') || classname === 'BOEING-EPIC-DL-FILTER-36MHZ-ULPC1') {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        extra = `<rect x="${x_loc - h_width}" y="${y_loc - h_height}" width="${width / 6}" height="${height}" fill="wheat"/>`;
    }
    else if (classname.includes('BOEING-EPIC-FWD-FILTER-62MHZ') || classname === 'BOEING-EPIC-DL-FILTER-62MHZ') {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        extra = `<rect x="${x_loc - h_width}" y="${y_loc - h_height}" width="${width / 6}" height="${height}" fill="darkred"/>`;
    }
    else if (classname.includes('IS32-TCR-TX-HYBRID')||classname.includes('IS32-TCR-RX-HYBRID')) {
        style = 'fill="none" stroke="black" stroke-width="0.75"';
        //extra = `<rect x="${x_loc - h_width}" y="${y_loc - h_height}" width="${width / 6}" height="${height}" fill="wheat"/>`;
    }
    else if (classname === 'RECEIVER-TYPE-3') {
        style = 'fill="none" stroke="blue" stroke-width="0.75"';
        extra = `\n<line x1="${x_loc - h_width}" y1="${y_loc}" x2="${x_loc + h_width}" y2="${y_loc}" stroke="black" stroke-width="0.75"/>`;
    }
    else if (classname === 'OMNI-ANTENNA') {
        style = 'fill="none"';
        extra = `\n<path d="M ${x_loc - h_width} ${y_loc - h_height} L ${x_loc + h_width} ${y_loc + h_height} L ${x_loc + h_width} ${y_loc - h_height} L ${x_loc - h_width} ${y_loc + h_height} L ${x_loc - h_width} ${y_loc - h_height}" fill="none" stroke="black"/>`;
    }
    else {
        console.log("Not-Handled: "+classname);
    }
    x_loc -= width / 2;
    y_loc -= height / 2;
    p2 =p
    if (!classname.includes('BORDER')){//&&!classname.includes('LORAL-12CH-ODD-OMUX')&&!classname.includes('IS32-KU-IMUX-12CH-O')) {
        shape_obj = new Shape(group_name, x_loc, -1 * y_loc, height,width,p2, right_multi, left_multi, top_multi, bottom_multi, channels);
        all_obj.push(shape_obj); //console.log(shape_obj)
    }
    else {
        svg_elm = `\n<g id="${group_name}" v:mID="${id}" v:groupContext="shape">\n`; //console.log(group_name)
        svg_elm += `<rect x="${x_loc}" y="${y_loc}" width="${width}" height="${height}" ${style}/>`;
        svg_elm += '\n</g>';
        return svg_elm;
    }

    svg_elm = `\n<g id="${group_name}" v:mID="${id}" v:groupContext="shape">\n`; //console.log(group_name)
    svg_elm += `<rect x="${x_loc}" y="${y_loc}" width="${width}" height="${height}" ${style}/>`;

    let x = 1;
// process attributes
    let chn= '';
    let ct;
    let formula = '';
    let itn = '';
    for (let i = 0; i < att_count; i++) {
        formula = line_arr[x + i].split(', ')[1].replace(';', '').replace('"', '').replace('"', '');
        if(formula ==="NA"){
            itn="NA"
        }
        else {
            let m = formula.match(/\d{3,}.-\w+(\.\w+)?/);
            itn = m[0];
        }
    }
    x += att_count;
    let top = [];
    let bottom = [];
    let left = [];
    let right = [];
    let top_line = [];
    let right_line = [];
    let left_line = [];
    let bottom_line = [];
    let all_port = [];
    let line_arr2 =[];
    for (let i = 0; i < ch_count; i++) { //console.log(ch_count) //console.log ('LINE2: '+line_arr[1])//console.log ('LINE: '+line_arr +' x: '+x +' i: '+i); console.log('\n'+'line: '+ line_arr[0].split(',')[0])
        let line = line_arr[x + i].split(', ');//console.log (line[0])
        all_port.push(line); //console.log(all_port)
        if (line[1] === 'LEFT') { //console.log(line [1])
            let temp = [];
            for (let q = 5; q < line.length; q++) {
                temp.push(line[q]);
            }
            temp.push(line[0])
            left.unshift(temp);
            left_line.unshift(line);
        }
        else if (line[1] === 'RIGHT') {
            let temp = [];
            for (let q = 5; q < line.length; q++) {
                temp.push(line[q]); //console.log(temp)
            }
            temp.push(line[0]);
            right.unshift(temp); //console.log(right)
            right_line.unshift(line);
        }
        else if (line[1] === 'TOP') {
            let temp = [];
            for (let q = 5; q < line.length; q++) {
                temp.push(line[q]); //console.log(temp)
            }
            temp.push(line[0])
            top.unshift(temp);
            top_line.unshift(line);
        }
        else if (line[1] === 'BOTTOM') {
            let temp = [];
            for (let q = 5; q < line.length; q++) {
                temp.push(line[q].trim());
            }
            temp.push(line[0].trim());
            bottom.unshift(temp);
            bottom_line.unshift(line);
        }
    }
    let top_axis = parseInt(y_loc); //console.log(top_axis)
    let bottom_axis = parseInt(y_loc + height);
    let left_axis = parseInt(x_loc);
    let right_axis = parseInt(x_loc + width);
    if (classname.includes('MUX') || classname.includes('IMUX')|| classname.includes('LORAL')&&!classname.includes('SWITCH')) {
        let index = 0;
        ct=1;
        for (let n = all_port.length - 1; n >= 0; n--) {
            let chName,pN; //console.log(all_port[n][1])
            if (all_port[n][0][0] === 'P') {
                shape_obj.channels++;
                if(all_port[n][0][0].startsWith('P')){
                    chName = parseInt(all_port[n][0].slice(1)); //console.log(chName+"="+ct)

                    ct++
                }
                else{
                    chName = ''
                }
                pN=parseInt(all_port[n][0].slice(1))
                if(pN=== 1){
                    pN=pN+1;
                }
                chn+=ct+";";
                extra += '\n<g id="channel_' + chName + '" v:mID="' + id + '" v:groupContext="shape">\n';
                if(all_port[n][1].includes("TOP")){
                    extra += '<rect x="' + ((x_loc + (index * (width / all_port.length))) + (index * 2) -11) + '" y="' + (y_loc-3 ) + '" width="20" height="20" fill="none"/>';
                    extra += '\n<text x="' + (((x_loc + (index * (width / all_port.length))) + (index * 2) -11)) + '" y="' + (y_loc-3) + '" font-size="15">' + chName + '</text>';
                }
                else if(all_port[n][1].includes("BOTTOM")){
                    extra += '<rect x="' + ((x_loc + (index * (width / all_port.length))) + (index * 2) + 8) + '" y="' + (y_loc +23) + '" width="20" height="20" fill="none"/>';
                    extra += '\n<text x="' + (((x_loc + (index * (width / all_port.length))) + (index * 2) + 8)) + '" y="' + (y_loc +23 ) + '" font-size="15">' + chName + '</text>';
                }
                else if (all_port[n][1].includes("RIGHT") || all_port[n][1].includes("LEFT")){
                    extra += '<rect x="' + (x_loc + (h_width / 2)) + '" y="' + ((y_loc + (index * (height / all_port.length))) + (index * 2) + 8) + '" height="11" width="' + (width - 6) + '" fill="none"/>';
                    extra += '\n<text x="' + (x_loc + (h_width / 2) - 2) + '" y="' + ((y_loc + (index * (height / all_port.length))) + (index * 2) + 8) + '" font-size="17">' + chName + '</text>';
                }
                extra += '\n</g>'; //console.log(extra)
                label_id ++
                index++;
            }
        }
    }
    let right_channel_taken = [];
    let left_channel_taken = [];
    let top_channel_taken = [];
    let bottom_channel_taken = [];
    for (let j = 0; j < shape_obj.channels; j++) {
        right_channel_taken.push(false);
        left_channel_taken.push(false);
        top_channel_taken.push(false);
        bottom_channel_taken.push(false);
    }
    if (top.length !== 0) { //console.log(top)
        const top_offset = parseInt(width) / (top.length + 1);
        for (let i = 0; i < top.length; i++) { //console.log(top)
            p= top[i][top[i].length -1];
            p2= top[i][top[i].length -1]
            conn_x_loc_1 = conn_x_loc_2 = x_loc + ((i + 1) * top_offset);
            conn_y_loc_1 = conn_y_loc_2 = parseInt(top_axis);
            let axis = 'y';
            let ctop = top[i].filter(el=> !isNaN(parseInt(el))); //console.log(ctop)
            for (let dist of ctop) {
                if (axis === 'x') {
                    conn_x_loc_2 += parseInt(dist.replaceAll(';', ''));
                    axis = 'y';
                } else {
                    conn_y_loc_2 += (-1 * parseInt(dist.replaceAll(';', ''))); //console.log(conn_y_loc_2)
                    axis = 'x';
                }
            }
            if (axis === 'x') {
                ending = 'y';
            }
            else {
                ending = 'x';
            }
            if (classname.includes('SWITCH')|| classname.includes('TOGGLE') && !classname.includes('SWITCHABLE')) {
                label_id++;
                id++;
                extra += '\n<g id="object' + label_id + '" v:mID="' + id + '" v:groupContext="shape">';
                if (width > 20) {
                    extra += '\n<rect x="' + parseInt(x_loc + (width / 2)) + '" y="' + y_loc + '" height="11" width="11" fill="none"/>'; //console.log("top: " +x_loc +" w: "+ width)
                    //extra += '\n<text x="' + x_loc + (width / 2) + '" y="' + y_loc + '" font-size="' + text_size + '">' + top_line[i][0][1] + '</text>'; //console.log("top1: " +extra)//console.log("top: " +x_loc +" w: "+ width/2)
                    extra += '\n<text x=" '+ parseInt(x_loc +(width/ 2)) + '" y="' + y_loc + '" font-size="'+text_size + '">' + top_line[i][0][1]+ '</text>';

                }
                else {
                    extra += '\n<rect x="' + x_loc + (width / 2) + '" y="' + (y_loc - 1) + '" height="11" width="11" fill="none"/>';//console.log(x_loc)
                    extra += '\n<text x=" '+ parseInt(x_loc +(width/ 2)) + '" y="' + (y_loc-1) + '" font-size="'+ text_size + '">' + top_line[i][0][1]+ '</text>';
                    //extra += `\n<text x="${parseInt(x_loc) + (parseInt(width) / 2)}" y="${parseInt(y_loc)-1}" font-size="${text_size}">${top_line[i][0][1]}</text>`;
                }
                extra += '\n</g>';
            }
            else if (classname === 'HOT-ARROW') {
                extra = `\n<title>HOT_ARROW_${label_id}</title>`;
                extra += '\n<path d="M' + (parseInt(x_loc) + (parseInt(width) / 2)) + ' ' +(parseInt(y_loc) + parseInt(height)) + ' L ' + (parseInt(x_loc) + parseInt(width)) + ' ' + (y_loc) + ' L ' + (parseInt(x_loc) + (parseInt(width) / 2)) + ' ' + (parseInt(y_loc) + 5) + ' L ' + (x_loc) + ' ' + (y_loc) + ' L ' + (parseInt(x_loc) + (parseInt(width) / 2)) + ' ' + (parseInt(y_loc) + parseInt(height)) + '" fill="blue"/>'
            }
            if (top_multi === true && shape_obj.channels !== 0) {
                const diff = parseInt(width) / parseFloat(shape_obj.channels);
                let dist = width - parseInt(top_line[i][2]); //console.log(diff+" : "+dist)
                conn_point = Math.floor(Math.abs(dist / diff)); //console.log(conn_point)
                if (conn_point < 0) {
                    conn_point = 0;
                }
                if (conn_point >= shape_obj.channels) {
                    conn_point = shape_obj.channels - 1;
                }
                while (conn_point < top_channel_taken.length - 1 && top_channel_taken[conn_point] === true) {
                    conn_point += 1;
                }
                while (conn_point > 0 && top_channel_taken[conn_point] === true) {
                    conn_point -= 1;
                }
                top_channel_taken[conn_point] = true;

            }
            else if (bottom_multi === true) {
                conn_point = 0;
            }
            else {
                conn_point = -1;
            }
            if (conn_point === -1) {
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), conn_x_loc_2, parseInt(-1 * conn_y_loc_2), 'top', ending,top_line[i][2],p)); //console.log(all_conn)
            }
            else {
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), conn_x_loc_2, parseInt(-1 * conn_y_loc_2), 'top' + conn_point.toString(), ending, top_line[i][2],p)); //onsole.log(all_conn)
            }
        }
    }
    if (bottom.length !== 0) { //console.log(bottom[0])
        const bottom_offset = parseInt(width) / (bottom.length + 1);
        for (let i = 0; i < bottom.length; i++) {
            p= bottom[i][bottom[i].length -1];
            //p2= bottom[i][bottom[i].length -1];
            conn_x_loc_1 = conn_x_loc_2 = parseInt(x_loc) + ((i + 1) * parseInt(bottom_offset));
            conn_y_loc_1 = conn_y_loc_2 = parseInt(bottom_axis);
            let axis = 'y';
            let cbottom = bottom[i].filter(el=> !isNaN(parseInt(el))); //console.log(cbottom)
            for (let dist of cbottom ) { //console.log(dist)
                if (axis === 'x') {
                    conn_x_loc_2 += parseInt(dist.replace(';', '')); //console.log(conn_x_loc_2)
                    axis = 'y';
                } else {
                    conn_y_loc_2 += -1 * parseInt(dist.replace(';', ''));
                    axis = 'x';
                }
            }
            if (axis === 'x') {
                ending = 'y';
            }
            else {
                ending = 'x';
            }
            if (classname.includes('SWITCH') || classname.includes('TOGGLE') && !classname.includes('SWITCHABLE')) {
                label_id++;
                id++;
                extra += '\n<g id="object' + label_id + '" v:mID="' + id + '" v:groupContext="shape">\n';
                if (width > 20) {
                    extra += '\n<rect x="' + parseInt(x_loc + (width / 2)) + '" y="' + parseInt(y_loc + (height-10))+ '" height="11" width="11" fill="none"/>';//console.log("btop: " +x_loc +" w: "+ width)
                    //extra += '\n<text x="' + x_loc + (width / 2) + '" y="' + y_loc +(height-10)+ '" font-size="' + text_size + '">' + bottom_line[i][0][1] + '</text>'; //console.log("btop1: " +extra)
                    extra += '\n<text x="' + parseInt(x_loc + (width / 2)) + '" y="' + parseInt(y_loc  +(height-10))+ '" font-size="' + text_size + '">' + bottom_line[i][0][1] + '</text>'; //console.log("btop2: " +extra)

                } else {
                    extra += '\n<rect x="' + parseInt(x_loc + (width / 2)) + '" y="' + parseInt(y_loc  +(height-10))+ '" height="11" width="11" fill="none"/>';
                    extra += '\n<text x="' + parseInt(x_loc + (width / 2)) + '" y="' + parseInt(y_loc  +(height-10))+ '" font-size="' + text_size + '">' + bottom_line[i][0][1] + '</text>'; //console.log("btop2: " +extra)
                }
                extra += '\n</g>';
            }
            else if (classname === 'HOT-ARROW') {
                extra = `\n<title>HOT_ARROW_${label_id}</title>`;
                extra += `\n<path d="M${parseInt(x_loc) + (parseInt(width) / 2)} ${parseInt(y_loc) + parseInt(height)} L ${parseInt(x_loc) + parseInt(width)} ${y_loc} L ${parseInt(x_loc) + (parseInt(width) / 2)} ${parseInt(y_loc) + 5} L ${x_loc} ${y_loc} L ${parseInt(x_loc) + (parseInt(width) / 2)} ${parseInt(y_loc) + parseInt(height)}" fill="blue"/>`;
            }
            if (bottom_multi === true && shape_obj.channels !== 0) {
                const diff = width / parseFloat(shape_obj.channels);
                const dist = width - parseInt(bottom_line[i][2]);
                conn_point = parseInt(Math.floor(Math.abs(dist / diff)));
                if (conn_point < 0) {
                    conn_point = 0;
                }
                if (conn_point >= shape_obj.channels) {
                    conn_point = shape_obj.channels - 1;
                }
                while (conn_point < bottom_channel_taken.length - 1 && bottom_channel_taken[conn_point] === true) {
                    conn_point += 1;
                }
                while (conn_point > 0 && bottom_channel_taken[conn_point] === true) {
                    conn_point -= 1;
                }
                bottom_channel_taken[conn_point] = true;
            }
            else if (top_multi === true) {
                conn_point = 0;
            }
            else {
                conn_point = -1;
            }
            if (conn_point === -1) {
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), conn_x_loc_2, (-1 * conn_y_loc_2), 'bottom', ending,bottom_line[i][2],p));
            }
            else {
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), conn_x_loc_2, (-1 * conn_y_loc_2), 'bottom' + conn_point.toString(), ending, bottom_line[i][2],p)); //console.log(all_conn)
            }
        }
    }
    if (left.length !== 0) { //console.log(left)
        const leftOffset = height / (left.length + 1);
        for (let i = 0; i < left.length; i++) { //console.log(left)
            p= left[i][left[i].length -1]
            //p2= left[i][left[i].length -1]
             conn_x_loc_1 =conn_x_loc_2= left_axis;
             conn_y_loc_1 = conn_y_loc_2= parseInt(y_loc + left_line[i][2]); //console.log("loc2: " +conn_x_loc_2)
            let axis = 'x';
            let cleft = left[i].filter(el=> !isNaN(parseInt(el))); //console.log(cleft)
            for (let dist of cleft) { //console.log("dist: " +dist)
                if (axis === 'x') {
                    conn_x_loc_2 += parseInt(dist.replace(';', '')); //console.log("loc2+dist: " +conn_x_loc_2)
                    axis = 'y';
                } else {
                    conn_y_loc_2 += parseInt(-1 * dist.replace(';', '')); //console.log("Eloc2+dist: " +conn_x_loc_2)
                    axis = 'x';
                }
            }
            if (axis === 'x') {
                ending = 'y';
            }
            else {
                ending = 'x';
            }
            if (classname.includes('SWITCH') || classname.includes('TOGGLE') && !classname.includes('SWITCHABLE')) {
                label_id += 1;
                id += 1;
                extra += `\n<g id="object${label_id}" v:mID="${id}" v:groupContext="shape">\n`;
                if (width > 20) {
                    extra += '\n<rect x="' +parseInt(x_loc + 2) + '" y="' + parseInt(y_loc + 5)+ '" height="11" width="11" fill="none"/>';
                    extra += '\n<text x="' + parseInt(x_loc + 2) + '" y="' + parseInt(y_loc +5)+ '" font-size="' + text_size + '">' + left_line[i][0][1] + '</text>';//{left_line[i][0][1]}</text>`; //console.log("x: "+(x_loc+2)+" "+ "y: "+(y_loc+5)+ " Tx: "+left_line[i][0][1])
                } else {
                    extra += '\n<rect x="' +parseInt(x_loc + 2) + '" y="' + parseInt(y_loc) + '" height="11" width="11" fill="none"/>';
                    extra += '\n<text x="' + parseInt(x_loc + 2) + '" y="' + parseInt(y_loc) + '" font-size="' + text_size + '">' + left_line[i][0][1] + '</text>'; //console.log("x2: "+(x_loc+2)+" "+ "y2: "+(y_loc+5)+ " Tx2: "+left_line[i][0][1])
                }
                extra += '\n</g>';
            }
            else if (classname === 'HOT-ARROW') {
                extra = `\n<title>HOT_ARROW_${label_id}</title>`;
                extra += `\n<path d="M${parseInt(x_loc) + parseInt(width)} ${parseInt(y_loc) + (parseInt(height) / 2)} L ${x_loc} ${y_loc} L ${parseInt(x_loc) + 5} ${parseInt(y_loc) + (parseInt(height) / 2)} L ${x_loc} ${parseInt(y_loc) + parseInt(height)} L ${parseInt(x_loc) + parseInt(width)} ${parseInt(y_loc) + (parseInt(height) / 2)}" fill="blue"/>`;
            }
            //let conn_point;
            if (left_multi===true && shape_obj.channels !== 0) {
                let diff = parseInt(height) / parseFloat(shape_obj.channels);
                let dist = parseInt(left_line[i][2]);
                 conn_point = parseInt(Math.floor(Math.abs(dist / diff)));
                if (conn_point < 0) {
                    conn_point = 0;
                }
                if (conn_point >= shape_obj.channels) {
                    conn_point = shape_obj.channels - 1;
                }
                while (conn_point < left_channel_taken.length - 1 && left_channel_taken[conn_point]=== true) {
                    conn_point += 1;
                }
                while (conn_point > 0 && left_channel_taken[conn_point]) {
                    conn_point -= 1;
                }
                left_channel_taken[conn_point] = true;
            }
            else if (right_multi===true) { //console.log(conn_point)
                conn_point = 0;
            }
            else {
                conn_point = -1;  //console.log(conn_point)
            } //console.log(classname)
            if (conn_point === -1) { //console.log(conn_point)
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), conn_x_loc_2, (-1 * conn_y_loc_2), 'left', ending, parseInt(left_line[i][2]),p)); //console.log(all_conn);
            } else { //console.log(conn_point)
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), conn_x_loc_2, (-1 * conn_y_loc_2), 'left' + conn_point.toString(), ending, parseInt(left_line[i][2]),p)); //console.log(all_conn);
            }
        }
    }
    if (right.length !== 0) { //console.log(right.flat(1))
        let rightOffset = height / (right.length + 1);
        let arr=[]; let p='';
        for (let i = 0; i < right.length; i++) {
            p= right[i][right[i].length -1];
            //p2= right[i][right[i].length -1];//console.log(p)
            conn_x_loc_1 = right_axis;
            conn_x_loc_2 = right_axis;
            conn_y_loc_1 = conn_y_loc_2 = parseInt(y_loc + right_line[i][2]);  //console.log("loc2: " +conn_x_loc_2)
            //if ('HYBRID-MUX-_86' === group_name) {
            //console.log(right_line[i]);
            //}
            let axis = 'x';
            let cRight = right[i].filter(el=> !isNaN(parseInt(el)));  //console.log(cRight)
            for (let dist of cRight) { //console.log("dist: " +dist)
                if (axis === 'x') {
                    conn_x_loc_2 += parseInt(dist.replace(';', '')); //console.log("loc2+dist: " +conn_x_loc_2)
                    axis = 'y';
                } else {
                    conn_y_loc_2 += parseInt(-1 * dist.replace(';', ''));
                    axis = 'x';
                }
            }
            if (axis === 'x') {
                ending = 'y';
            }
            else {
                ending = 'x';
            }
            if (classname.includes('SWITCH')|| classname.includes('TOGGLE') && !classname.includes('SWITCHABLE')) {
                label_id += 1;
                id += 1;
                extra += '\n<g id="object' + label_id + '" v:mID="' + id +'" v:groupContext="shape">\n';

                if (width > 20) {
                    extra += '\n<rect x="' + parseInt(x_loc + (width - 9)) + '" y="' + parseInt(y_loc + 5) + '" height="11" width="11" fill="none"/>'; //console.log("Rtop: " +x_loc +" w: "+ width)
                    extra += '\n<text x="' + parseInt(x_loc + (width - 9)) + '" y="' + parseInt(y_loc + 5) + '" font-size="' + text_size + '">' + right_line[i][0][1] + '</text>'; //console.log("Rtop1: " +extra)//console.log("Rtop: " +x_loc +" w: "+ (width-9))
                } else {
                    extra += '\n<rect x="' + parseInt(x_loc + (width - 5)) + '" y="' + (y_loc) + '" height="11" width="5" fill="none"/>'; //console.log("else:" + x_loc)
                    extra += '\n<text x="' + parseInt(x_loc + (width - 5)) + '" y="' + (y_loc) + '" font-size="' + text_size + '">' + right_line[i][0][1] + '</text>'; //console.log("Rtop2: " +extra)//console.log("Rtop: " +x_loc +" w: "+ (width-5))
                }
                extra += '\n</g>';
            }
            else if (classname === 'HOT-ARROW') {
                extra = `\n<title>HOT_ARROW_${label_id}</title>`;
                extra += `\n<path d="M${parseInt(x_loc) + (parseInt(width) / 2)} ${parseInt(y_loc) + parseInt(height)} L ${parseInt(x_loc) + parseInt(width)} ${y_loc} L ${parseInt(x_loc) + (parseInt(width) / 2)} ${parseInt(y_loc) + 5} L ${x_loc} ${y_loc} L ${parseInt(x_loc) + (parseInt(width) / 2)} ${parseInt(y_loc) + parseInt(height)}" fill="blue"/>`;
            }

            if (right_multi === true && shape_obj.channels !== 0) {
                const diff = parseInt(height) / parseFloat(shape_obj.channels);
                const dist = parseInt(right_line[i][2]);
                let conn_point = Math.floor(Math.abs(dist / diff));
                if (conn_point < 0) {
                    conn_point = 0;
                }
                if (conn_point >= shape_obj.channels) {
                    conn_point = shape_obj.channels - 1;
                }
                while (conn_point < right_channel_taken.length - 1 && right_channel_taken[conn_point] === true) {
                    conn_point += 1;
                }
                while (conn_point > 0 && right_channel_taken[conn_point] === true) {
                    conn_point -= 1;
                    right_channel_taken[conn_point] = true;
                }
            }
            else if (left_multi === true) {
                conn_point = 0;
            }
            else {
                conn_point = -1;
            }
            if (conn_point === -1) {
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), conn_x_loc_2, (-1 * conn_y_loc_2), 'right', ending, parseInt(right_line[i][2]),p));
            }
            else {
                all_conn.push(new Connection(conn_x_loc_1, (-1 * conn_y_loc_1), parseInt(conn_x_loc_2), (-1 * conn_y_loc_2), 'right' + conn_point.toString(), ending, parseInt(right_line[i][2]),p)); ///console.log(all_conn)
            }
        }
    }
    svg_elm += extra;
    if (classname.includes('SWITCH') || classname.includes('TOGGLE') && !classname.includes('SWITCHABLE')) { //console.log(true)
        props+= ('\n<v:cp v:nameU="Orbital_PN" v:lbl="Orbital/PN" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="Vendor_PN" v:lbl="Vendor_PN" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="Vendor_ICD#" v:lbl="Vendor_ICD" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="Mass" v:lbl="Mass" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="DC_Power" v:lbl="DC_Power" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="Loss" v:lbl="Loss" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="Position" v:lbl="Position" v:type="0" v:langID="1033" v:val="VT4(' + 1 + ')" />');
        props += ('\n<v:cp v:nameU="Random_Access" v:lbl="Random_Access" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="Mnemonic" v:lbl="Mnemonic" v:type="0" v:langID="1033" v:val="VT4(' + itn.replaceAll(scid+'-',"") + ')" />');
        props +=('\n<v:cp v:nameU="Bank" v:lbl="Bank" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += ('\n<v:cp v:nameU="Number" v:lbl="Number" v:type="0" v:langID="1033" v:val="VT4()" />');
        props += '\n<v:cp v:nameU="SwitchClass" v:lbl="SwitchClass" v:type="0" v:langID="1033" v:val="VT4(' + classname + ')" />';
        props += '\n<v:cp v:nameU="j-Ports" v:lbl="j-Ports" v:type="0" v:langID="1033" v:val="VT4()" />';
        if (objname !== 'NONE') {
            props += '\n<v:cp v:nameU="Designator" v:lbl="Designator" v:type="0" v:langID="1033" v:val="VT4(' + objname + ')" />';
        } else {
            props += '\n<v:cp v:nameU="Designator" v:lbl="Designator" v:type="0" v:langID="1033" v:val="VT4()" />';
        }
        props += '\n<v:cp v:nameU="Position" v:lbl="Position" v:type="0" v:langID="1033" v:val="VT4(' + 1 + ')" />';
        props += '\n<v:cp v:nameU="Orientation" v:lbl="Orientation" v:type="0" v:langID="1033" v:val="VT4()" />';
    }
    if (classname.includes('RECEIVER')) {
        props+= '\n<v:cp v:nameU="Orbital_PN" v:lbl="Orbital_PN" v:type="0" v:langID="1033" v:val="VT4()" />';
        props += '\n<v:cp v:nameU="Vendor_PN" v:lbl="Vendor_PN" v:type="0" v:langID="1033" v:val="VT4()" />';
        props += '\n<v:cp v:nameU="Vendor_ICD#" v:lbl="Vendor_ICD#" v:type="0" v:langID="1033" v:val="VT4()" />';
        props += '\n<v:cp v:nameU="Mass" v:lbl="Mass" v:type="0" v:langID="1033" v:val="VT4()" />';
        props += '\n<v:cp v:nameU="DC_Power" v:lbl="DC_Power" v:type="0" v:langID="1033" v:val="VT4()" />';
        props += '\n<v:cp v:nameU="Gain" v:lbl="Gain" v:type="0" v:langID="1033" v:val="VT4()" />';
        props += '\n<v:cp v:nameU="Mnemonic" v:lbl="Mnemonic" v:type="0" v:langID="1033" v:val="VT4(' + itn.replaceAll(scid+'-',"") + ')" />';
        if (objname !== 'NONE') {
            props += '\n<v:cp v:nameU="Designator" v:lbl="Designator" v:type="0" v:langID="1033" v:val="VT4(' + objname + ')" />';
        } else {
            props += ('\n<v:cp v:nameU="Designator" v:lbl="Designator" v:type="0" v:langID="1033" v:val="VT4()" />')
            props += ('\n<v:cp v:nameU="Function" v:lbl="Function" v:type="0" v:langID="1033" v:val="VT4()" />')
            props += ('\n<v:cp v:nameU="Spare" v:lbl="Spare" v:type="0" v:langID="1033" v:val="VT4()" />')
            props += ('\n<v:cp v:nameU="LO" v:lbl="LO" v:type="0" v:langID="1033" v:val="VT4()" />')
        }
    }
    else if (classname.includes('DOWN-CONVERTER')) {
        props += ('\n<v:cp v:nameU="Mnemonic" v:lbl="Mnemonic" v:type="0" v:langID="1033" v:val="VT4(' + itn + ')" />');
        if (objname !== 'NONE') {
            props += ('\n<v:cp v:nameU="Designator" v:lbl="Designator" v:type="0" v:langID="1033" v:val="VT4(' + objname + ')" />');
        } else {
            props += ('\n<v:cp v:nameU="Designator" v:lbl="Designator" v:type="0" v:langID="1033" v:val="VT4()" />');
        }
    }
        // else if (classname.includes('TCR-XMTR')) {
        //     props = ('\n<v:cp v:nameU="Link" v:lbl="Link" v:type="0" v:langID="1033" v:val="VT4(' + line_arr[line_arr.length - 1].split(": ")[1].replace(';', '').toUpperCase().replaceAll('-','_').replace(scid, 'G'+scid.slice(-2)) + ')" />');
        //     props += ('\n<v:cp v:nameU="Trace" v:lbl="Trace" v:type="0" v:langID="1033" v:val="VT4()" />');
    // }
    else if (classname.includes('HOT-ARROW')) {
        //console.log(classname); console.log("line: "+line_arr[line_arr.length-1])
        props += ('\n<v:cp v:nameU="Link" v:lbl="Link" v:type="0" v:langID="1033" v:val="VT4(' + line_arr[line_arr.length - 1].split(":")[1].replace(';', '').toUpperCase().replaceAll('-','_').trim() + ')" />');
        //props += ('\n<v:cp v:nameU="Trace" v:lbl="Trace" v:type="0" v:langID="1033" v:val="VT4()" />');
    }
    else if (classname.includes('ANTENNA')) {//console.log(line_arr)
        let prt1, prt2;
        let prtaux=''; let prtaux2 ='';
        for(let i =1; i<line_arr.length; i++)
            if (classname.includes('QUAD')){
                if(line_arr[i].includes('TOP')){prt1 = line_arr[i].split(',')[0];prtaux = line_arr[i].split(',')[0];
                }
                else if(line_arr[i].includes('BOTTOM')) {prt2 = line_arr[i].split(',')[0];prtaux2 = line_arr[i].split(',')[0];
                }
            }
            else if (classname.includes('DUAL')){
                if(line_arr[i].includes('TOP')){prt1 = line_arr[i].split(',')[0]
                }
                else if(line_arr[i].includes('BOTTOM')) {prt2 = line_arr[i].split(',')[0]
                }
            }
            else {
                if(line_arr[i].includes('NONE')){prt1 = line_arr[0].split(',')[2].replaceAll('"','');prtaux = ''; prt2=''; //console.log(prt1)
                }
            }
        props += ('\n<v:cp v:nameU="topPort" v:lbl="topPort" v:type="0" v:langID="1033" v:val="VT4('+prt1+')" />');
        props += ('\n<v:cp v:nameU="bottomPort" v:lbl="bottomPort" v:type="0" v:langID="1033" v:val="VT4('+prt2+')" />');
        props += ('\n<v:cp v:nameU="topAux" v:lbl="topAux" v:type="0" v:langID="1033" v:val="VT4('+prtaux+')" />');
        props += ('\n<v:cp v:nameU="bottomAux" v:lbl="bottomAux" v:type="0" v:langID="1033" v:val="VT4('+prtaux2+')" />');
    }
    else if (classname.includes('MUX')||classname.includes('LORAL')&&!classname.includes('SWITCH')) {
        //console.log(line_arr[0].split(',')[0].match(/\b(\w*MUX\w*)\b/)[1])
        props += ('\n<v:cp v:nameU="Mux_Type" v:lbl="Mux_Type" v:type="0" v:langID="1033" v:val="VT4('+line_arr[0].split(',')[0].match(/\b(\w*MUX\w*)\b/)[1]+')" />');
        props += ('\n<v:cp v:nameU="Channel_Name" v:lbl="Channel_Name" v:type="0" v:langID="1033" v:val="VT4('+"1;"+chn+')" />');
    }
    svg_elm += ('\n<v:custProps>');
    svg_elm += props;
    svg_elm += ('\n</v:custProps>');
    label_id += 1;
    id += 1;
    svg_elm += ('\n</g>');
    return svg_elm;
}

function processG22xreadouts (line_arr) { //console.log(line_arr.split(",")[1])
    let c=0;
    if (line_arr[1].trim() === ("SMALL-G22X-READOUTS")) { //console.log(line_arr)
        readOutLcamp.push({
            "group":  line_arr[6].replaceAll(/[^\w\s.-]/g,"").toString().trim(),
            "twtaNum": "G22X",
            "readout":  "G22X-READOUTS",
            "Mnemonic": line_arr[6] + "edge",
            "formula": line_arr[6]
        })
        c++;
    }
}

function process_readout(line_arr) { //console.log(line_arr)
    let pin, mnemonic, group,readOut;
    let twtaNum, readout, Mnemonic, formula;
    let prop = '';
    line_arr = line_arr.split(","); //console.log(line_arr);
    processG22xreadouts(line_arr)
    for (let i = 0; i < line_arr.length; i++) {
        line_arr[i] = line_arr[i].trim();
    }
    if (line_arr[1] === "UTC-CLOCK" || line_arr[1] === "DISPLAY-HEADERS") {
        return "";
    }
    let x_loc = parseInt(line_arr[2]) + 20; //console.log(x_loc)
    let y_loc = -1 * parseInt(line_arr[3]) + 8;
    let width = parseInt(line_arr[4]);
    let height = parseInt(line_arr[5]) //.replace(",", ""));
    let h_width = width / 2;
    let h_height = height / 2;
    x_loc -= h_width;
    y_loc -= h_height;

    if (line_arr[6].includes("[")) {
        group=  line_arr[6].replaceAll(/[^\w\s.-]/g,"").toString().trim(); //console.log(group)
        readOut = line_arr[1].split('SMALL-')[1]; //console.log(readOut)
        pin = "G22X"; //console.log("rpin:"+pin)
    }
    else{
        group = line_arr[6]; //console.log(group)
        readOut = line_arr[7];

        if (line_arr[6].includes("CBRX")) {
            let str = line_arr[6].split('-')[1];
            pin = "C0" + str[str.length - 1]; //console.log(pin)
        }
        else if (line_arr[6].includes("RCVR")) {
            pin = line_arr[6].split("-")[2]; //console.log("rpin:"+pin)
        }
        else if (line_arr[6].includes("BEACON")) {
            let pn = line_arr[6].split("-")
            pin = pn[pn.length -1]
        }
        else {
            if (line_arr[6].split("-").length >= 5) {
                pin = line_arr[6].split("-")[4]; //console.log(pin)
            }
            else {
                pin = line_arr[6].split("-")[2]; //console.log(pin)

            }
        }
    }
    //console.log(pin);
    readOutLcamp.forEach(el => { let color= '#fff2cc'; //console.log(el)
        //readOut.match(/AGC-VOLTAGE/)||readOut.match(/AGC-LEVEL/)||readOut.match(/RANGING-STATUS/)||readOut.match(/POWER-MODE/)||readOut.match(/LIN-STEP/)||readOut.match(/LCHAMP-OUTLEVEL/)||readOut.match(/LCHAMP-RFBLANK/)||readOut.match(/LCHAMP-BAND/)||readOut.match(/LCHAMP-GAIN/)||readOut.match(/LOAD-CURRENT/)||readOut.match(/HELIX-CURRENT/)||readOut.match(/FIL-VOLTAGE/)||
        //if(readOut.match(/BUS-CURRENT/)||readOut.match(/HELIX-CURRENT/)||readOut.match(/ARU-STS/)){color="wheat"}else{color="wheat"}
        svg_elm = '\n<g id="Readout.'+readOut+"_" + label_id + '" v:mID="' + id + '" v:groupContext="shape">\n';
        svg_elm += '<rect x="' + x_loc + '" y="' + y_loc + '" width="' + width / 2 + '" height="' + height / 2 + '" fill="'+color+'"/>';
        let twta = el.twtaNum.replaceAll('"', '').trim(); //console.log('pin: '+ pin+" "+'twta: '+twta+ ' '+'G: '+group+ ' '+'elg:'+el.group+" "+'R: '+readOut+ ' '+  el.readout.replaceAll('"', ''));
        if (pin === twta && readOut === el.readout.replaceAll('"', '')&& group.includes(el.group.trim())) { //console.log('pin: '+ pin+" "+'twta: '+twta+ ' '+'G: '+group+ ' '+el.group+" "+'R: '+readOut+ ' '+  el.readout.replaceAll('"', ''));//console.log(true) //('pin: '+ pin+" "+'twta: '+twta);
            Mnemonic= el.Mnemonic.replaceAll(/[^\w\s.-]/g,"").toString().trim(); //console.log(readOut +" : "+el.Mnemonic)
            if(Mnemonic.startsWith("safe-symbol")||Mnemonic.startsWith("the")){
                mnemonic =Mnemonic.split(scid.toString())[1].replace("-","").replace(" edge","").trim(); //console.log(readOut +" : " +mnemonic)
            }
            else if(Mnemonic.startsWith(scid.toString())){
                mnemonic =Mnemonic.split(" ")[0].replace(scid.toString()+"-","");
            }
            else if (Mnemonic.startsWith("round")) {
                let parts = Mnemonic.trim().split(/\s+/);
                if (parts.length > 1) {
                    mnemonic = parts[1].replace(scid.toString() + "-", "");
                } else {
                    mnemonic = "["+Mnemonic.split('round')[1].replace(scid.toString() + "-", "")+"]";
                }
            }
            else if(Mnemonic.startsWith("if")){ //console.log(Mnemonic.split( scid.toString()+"-")[1].split(" ")[0])
                if(Mnemonic.includes("round")) {
                    mnemonic = Mnemonic.split("round")[1].replace(scid.toString()+"-", "").replace(" else", ""); //console.log(mnemonic)
                }else if(Mnemonic.includes("OK")) { //console.log(Mnemonic)
                    mnemonic = Mnemonic.split("OK")[1].replace(" else if " + scid.toString()+"-", "").replace("1 then Trip else", "");//console.log("mn: "+mnemonic)
                }else if(Mnemonic.includes("--")){ //console.log(Mnemonic)
                    mnemonic = Mnemonic.split("--")[1].replace(" else " + scid.toString()+"-", "").replace("as dd.d", "");
                }
                else if(Mnemonic.includes("CMR")){ //console.log(Mnemonic)
                    mnemonic = Mnemonic.split( scid.toString()+"-")[1].split(" ")[0]; //console.log(mnemonic)
                }
                else if(Mnemonic.includes("edge")){ //console.log(Mnemonic)
                    mnemonic = Mnemonic.split( scid.toString()+"-")[1].split(" ")[0]; //console.log(mnemonic)
                }
            }
            twtaNum=twta; //console.log("TEST: "+ el.formula)
            formula=el.formula.replaceAll('"', '').replace(";","").replaceAll(scid.toString()+"-","");//console.log("TEST: "+formula);
            prop += '\n<v:cp v:nameU="mnemonic" v:lbl="mnemonic" v:type="0" v:langID="1033" v:val="VT4(' + mnemonic + ')" />';
            prop += '\n<v:cp v:nameU="PinID" v:lbl="PinID" v:type="0" v:langID="1033" v:val="VT4()" />';
            prop += '\n<v:cp v:nameU="InternalID" v:lbl="InternalID" v:type="0" v:langID="1033" v:val="VT4()" />';
            prop += '\n<v:cp v:nameU="TwtaNum" v:lbl="TwtaNum" v:type="0" v:langID="1033" v:val="VT4(' + twtaNum + ')" />';
            prop += '\n<v:cp v:nameU="Formula" v:lbl="Formula" v:type="0" v:langID="1033" v:val="VT4(' + formula + ')" />';
        }
        svg_elm += ('\n<v:custProps>');
        svg_elm += prop;
        svg_elm += ('\n</v:custProps>');
    })
    label_id += 1;
    id += 1;
    svg_elm += ('\n</g>'); //console.log(svg_elm)
    return  svg_elm;
}

const runApp =(fileName)=>{
//let storage_dir = "./"//'C:\\Users\\scotch\\rcap\\rcap\\ascii_convert\\automation'; //\\rcap_34_realtime\\static\\etc\\rcap_34_realtime';
let section = 'display';
let filename= fileName.split('\\')[3]
//console.log('opening ' + filename);
  let file =fileName
    let local_lines = []
    let width, height;
    try {
        let i = 0;
        let lines = fs.readFileSync(file, 'utf-8').split('\n');
        let section = '';
        let width, height,displayName;
        let local_lines = [];
        let svg_elm = ''; 
        while (i < lines.length) {
            if (lines[i].trim().length > 0) {
                if (lines[i].trim()[0] === '*') {
                    if (lines[i].includes('OBJECT LIST')) {
                        section = 'object list';
                    } else if (lines[i].includes('READOUT LIST')) {
                        section = 'readout list';
                    } else if (lines[i].includes('LABEL LIST')) {
                        section = 'label list';
                    }
                    i++;
                    continue;
                } 
                else {
                    if (lines[i].includes('DispWidth')) {
                        width = lines[i].split(', ')[0].split(': ')[1]; //console.log("W: " +width)
                        height = lines[i].split(', ')[1].split(': ')[1]; //console.log("H: " +height)
                        i++;
                        continue;
                    } 
                    else if (lines[i].includes('Map')) {
                        displayName = lines[i].split(', ')[0].split(':')[1].trim().replace(/"/g, '').replace(' ', '_').replaceAll('-', '_').replace(/\([^)]*\)/g, ''); //console.log("Dispaly: " +displayName)
                        scid = lines[i].split(', ')[0].split(': ')[1].split(" ")[0].replace('"',''); //console.log("scid: " +scid)
                        i++;
                        continue;
                    }
                    else if (section === 'object list') {
                        local_lines.push(lines[i].trim());
                        i+=1; //console.log(local_lines)
                        while (lines[i][0] === ' ') {
                            local_lines.push(lines[i].trim());
                            i+=1; //console.log(local_lines)
                        }//console.log(local_lines)
                        svg_elm += process_object(local_lines);
                        local_lines = [];
                        continue;
                    } else if (section === 'readout list') {
                        while (!lines[i].includes(';')) {
                            local_lines.push(lines[i].trim());
                            i++; //console.log(local_lines)
                        }
                        local_lines.push(lines[i].trim());
                        i++;
                        let arr='';
                        for (let j=0; j<local_lines.length; j++) {
                            arr += ' '+ local_lines[j].trim()
                        }
                        svg_elm += process_readout(arr);
                        local_lines = [];
                        continue;
                    } else if (section === 'label list') {

                        while (!lines[i].includes(';')) {
                            local_lines.push(lines[i].trim());
                            i++;
                        }
                        local_lines.push(lines[i].trim());
                        i++; //console.log(local_lines)
                        let arr='';
                        for (let j=0; j<local_lines.length; j++) {
                            arr += ' '+ local_lines[j].trim()
                        }
                        svg_elm += process_label(arr);
                        local_lines = [];
                        continue;
                    }
                }
            }
            i++;
        }
        let final_conn = [];
        let mod_list = [0.125, 0.25, 0.375, 0.5, 0.625, 0.75, 0.875];
        for (let c of all_conn) {  //console.log(all_conn)
            let shape1 = '';
            for (let k = 0; k < 13; k++) {
                for (let s of all_obj) { //console.log(c)
                    if (s.inBounds(c.x1 + k, c.y1)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1 + k, c.y1)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1 - k, c.y1)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1, c.y1 + k)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1, c.y1 - k)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1 + k, c.y1 + k)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1 + k, c.y1 - k)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1 - k, c.y1 + k)) {
                        shape1 = s;
                        break;
                    } else if (s.inBounds(c.x1 - k, c.y1 - k)) {
                        shape1 = s;
                        break;
                    }
                    for (let mod of mod_list) {
                        if (s.inBounds(c.x1 + (k * mod), c.y1 + k)) {
                            shape1 = s;
                            break;
                        } else if (s.inBounds(c.x1 + k, c.y1 + (k * mod))) {
                            shape1 = s;
                            break;
                        } else if (s.inBounds(c.x1 - (k * mod), c.y1 - k)) {
                            shape1 = s;
                            break;
                        } else if (s.inBounds(c.x1 - k, c.y1 - (k * mod))) {
                            shape1 = s;
                            break;
                        } else if (s.inBounds(c.x1 + (k * mod), c.y1 - k)) {
                            shape1 = s;
                            break;
                        } else if (s.inBounds(c.x1 + k, c.y1 - (k * mod))) {
                            shape1 = s;
                            break;
                        } else if (s.inBounds(c.x1 - (k * mod), c.y1 + k)) {
                            shape1 = s;
                            break;
                        } else if (s.inBounds(c.x1 - k, c.y1 + (k * mod))) {
                            shape1 = s;
                            break;
                        }
                    }
                    if (shape1 !== '') {
                        break;
                    }
                }
                if (shape1 !== '') {
                    break;
                }
            }
            let ending_side = '';
            let shape2 = new Shape('', 0, 0, 0, 0,''); //console.log(shape2)
            for (let k = 0; k < 13; k++) {
                for (let s of all_obj) { //console.log(all_obj)
                    if (s.inBounds(c.x2 + k, c.y2) && s.name !== shape1.name) {
                        shape2 = s; 
                        break;
                    } else if (s.inBounds(c.x2 - k, c.y2) && s.name !== shape1.name) {
                        shape2 = s; //console.log(s)
                        break;
                    } else if (s.inBounds(c.x2, c.y2 + k) && s.name !== shape1.name) {
                        shape2 = s;
                        break;
                    } else if (s.inBounds(c.x2, c.y2 + k) && s.name !== shape1.name) {
                        shape2 = s;
                        break;
                    } else if (s.inBounds(c.x2, c.y2 - k) && s.name !== shape1.name) {
                        shape2 = s;
                        break;
                    } else if (s.inBounds(c.x2 + k, c.y2 + k) && s.name !== shape1.name) {
                        shape2 = s;
                        break;
                    } else if (s.inBounds(c.x2 + k, c.y2 - k) && s.name !== shape1.name) {
                        shape2 = s;
                        break;
                    } else if (s.inBounds(c.x2 - k, c.y2 + k) && s.name !== shape1.name) {
                        shape2 = s;
                        break;
                    } else if (s.inBounds(c.x2 - k, c.y2 - k) && s.name !== shape1.name) {
                        shape2 = s;
                        break;
                    } else {
                        for (let mod of mod_list) {
                            if (s.inBounds(c.x2 + (k * mod), c.y2 + k) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            } else if (s.inBounds(c.x2 + k, c.y2 + (k * mod)) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            } else if (s.inBounds(c.x2 - (k * mod), c.y2 - k) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            } else if (s.inBounds(c.x2 - k, c.y2 - (k * mod)) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            } else if (s.inBounds(c.x2 + (k * mod), c.y2 - k) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            } else if (s.inBounds(c.x2 + k, c.y2 - (k * mod)) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            } else if (s.inBounds(c.x2 - (k * mod), c.y2 + k) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            } else if (s.inBounds(c.x2 - k, c.y2 + (k * mod)) && s.name !== shape1.name) {
                                shape2 = s;
                                break;
                            }
                            if (shape2.name !== '') {
                                break;
                            }
                        }
                    }
                }
            } //console.log(shape2)
            if (shape1.name !== '' && shape2.name !== '') {
                let found = false;
                const pair = new Pair(shape1.name, shape2.name, c.p, c.initial_side,shape2.p2); //console.log(pair)
                let p1_dist, p2_dist, closest_side = '';
                if (c.ending_direction === 'x') {
                    p1_dist = Math.abs(c.x2 - shape2.x1);
                    p2_dist = Math.abs(c.x2 - shape2.x2);
                } else {
                    p1_dist = Math.abs(c.y2 - shape2.y1);
                    p2_dist = Math.abs(c.y2 - shape2.y2);
                }

                if (p1_dist <= p2_dist) {
                    if (c.ending_direction === 'x') {
                        closest_side = 'left';
                    } else {
                        closest_side = 'top';
                    }
                } else {
                    if (c.ending_direction === 'x') {
                        closest_side = 'right';
                    } else {
                        closest_side = 'bottom';
                    }
                }
                let dist = -1, diff = -1, section = -1, height_t = -1;
                if (closest_side === 'left' && shape2.left_multi === true) {
                    dist = Math.abs(shape2.y1 - c.y2);
                    diff = shape2.height / parseFloat(shape2.channels);
                    height_t = shape2.height;
                    section = Math.floor(dist / diff);
                    closest_side += section.toString();
                } else if (closest_side === 'left' && shape2.right_multi === true) {
                    closest_side += '0';
                } else if (closest_side === 'right' && shape2.right_multi === true) {
                    dist = Math.abs(shape2.y1 - c.y2);
                    diff = shape2.height / parseFloat(shape2.channels);
                    section = Math.floor(dist / diff);
                    height_t = shape2.height;
                    closest_side += section.toString();
                } else if (closest_side === 'right' && shape2.left_multi === true) {
                    closest_side += '0';
                }
                if (closest_side === 'top' && shape2.top_multi === true) {
                    let dist = Math.abs(shape2.x1 - c.x2);
                    let diff = shape2.width / parseFloat(shape2.channels);
                    let width_t = shape2.width;
                    let section = Math.floor(dist / diff);
                    closest_side += section.toString();
                } else if (closest_side === 'top' && shape2.bottom_multi === true) {
                    closest_side += '0';
                }
                if (closest_side === 'bottom' && shape2.bottom_multi === true) {
                    let dist = Math.abs(shape2.x1 - c.x2);
                    let diff = shape2.width / parseFloat(shape2.channels);
                    let width_t = shape2.width;
                    let section = Math.floor(dist / diff);
                    closest_side += section.toString();
                } else if (closest_side === 'bottom' && shape2.top_multi === true) {
                    closest_side += '0';
                }
                pair.ending_side = closest_side;
                for (const f of final_conn) {
                    if (f.isEqual(pair)) {
                        if (pair.initial_side !== pair.ending_side) {
                            if (pair.initial_side === f.initial_side || pair.s1 === f.s2) {
                                found = true;
                                break;
                            }
                        } else if (pair.initial_side === pair.ending_side && pair.initial_side === f.initial_side && pair.initial_side === f.ending_side) {
                            found = true;
                            break;
                        }
                    }
                }
                if (found === false) {
                    //console.log("pair:" +pair);
                    final_conn.push(pair);
                } //console.log(final_conn)
            }
        }
        id += 1;
        let currentDate = new Date();
        let dateString = currentDate.toString().split(" ").slice(0,5).join('-'); //console.log(dateString)
        let background = '\n<g id="background" v:mID="'+ id + '" v:groupContext="shape">\n';
        background += '<rect x="0" y="0" width="'+ width+'" height="'+height+'" fill="wheat"/>';
        background += `\n<v:custProps>`;
        // for (const f of final_conn) {  //console.log(final_conn)
        //         background += `\n<v:cp v:nameU="conn${conn_index}" v:lbl="conn${conn_index}" v:type="0" v:langID="1033" v:val="VT4(${f.s1}, ${f.s2}, ${f.initial_side},${f.p},${f.ending_side},${f.p2}` + ')" />';
        //     conn_index += 1;
        // }
        // let connInfo =JSON.parse(fs.readFileSync('./output/cxn_file.json'))
        // for(let i =0; i< connInfo.length; i++){
        //     //conn_index = 0;
        //     background += `\n<v:cp v:nameU="conn${i}" v:lbl="conn${i}" v:type="0" v:langID="1033" v:val="VT4(${connInfo[i].shp1}, ${connInfo[i].shp2}, ${connInfo[i].shp1_port},${connInfo[i].shp1_side},${connInfo[i].shp2_port},${connInfo[i].shp2_side}` + ')" />';
        //     //conn_index += 1;
        // }
        background += '\n</v:custProps>';
        label_id += 1;
        id += 1;
        background += '\n</g>';
        svg_elm +=`\n<text x="${parseInt(width)+parseInt(200)}" y="${0}" text-anchor="end" font-family="Arial" font-size="14" dy="0.5em" fill="blue">Date Modified: ${dateString}</text>`;
        svg_elm = `<svg style="margin: 0 auto;" version="1.1" width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg" xmlns:v="http://schemas.microsoft.com/visio/2003/SVGExtensions;">${background}${svg_elm}</svg>`;
        let shpAmount = nameArray.length
        let blob = ""
        for (let lj=0;lj < shpAmount; lj++){
            blob += nameArray[lj] + "\n"
        }; //console.log(blob)
        let newFileName = filename.replace(".","-") + ".txt"
        fs.writeFileSync("./output/"+newFileName, blob,'utf8')
    }catch (e) {
        console.log(e);
    }
};runApp(inputFile)




