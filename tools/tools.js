var fs = require("fs");
var path = require("path");

var filelist = [];

var cmd = process.argv;
if (cmd.length <= 2) {
    alert('参数错误，参数个数：0');
}else{
    var args = cmd.slice(2);
    if (args.length < 2) {
        alert('参数错误，json文件路径或ktjson.js路径参数不正确');
    }else if(args instanceof Array) {
        var josn_path = args[0];
        readDir(josn_path);
        console.log(filelist.length);
        if (filelist.length < 1) {
            alert('路径错误，当前路径下不包含json文件');
            return;
        }
    
        var kt_path = args[1];
        console.log(kt_path);
        if (kt_path.indexOf("ktjson.js") == -1) {
            alert('路径错误，不包含ktjson.js');
            return;
        }
            
        writefileHead(kt_path);
        writeJs(kt_path);
        writefileEnd(kt_path);    
    }
}

//读取当前路径下所有json文件
function readDir(path) {
    try {
        var pa = fs.readdirSync(path);
        pa.forEach(function (ele, index) {
            var info = fs.statSync(path + "/" + ele);
            if (!info.isDirectory() && ele.indexOf(".json") > 0) {
                filelist.push(path + "/" + ele);
            }
        });
    } catch (e) {
       
    }
}

//拼接所有json文件内容
function writeJs(dest_dir) {
    for (let index = 0; index < filelist.length; index++) {
        var ele = filelist[index];
        var data = fs.readFileSync(ele);
        fs.appendFileSync(dest_dir, get_file_name(ele) + ":" + data + ",", { encoding: 'utf8', mode: 438 /*=0666*/, flag: 'a' });
    }
}

/**
 * 拼接ktjons.js头部"var ktjson={"
 * @param {*} dest_dir 
 */
function writefileHead(dest_dir) {
    fs.openSync(dest_dir, 'w');
    fs.appendFileSync(dest_dir, "var ktjson={", { encoding: 'utf8', mode: 438 /*=0666*/, flag: 'a' });
}

/**
 * 拼接ktjons.js尾部"};"
 * @param {*} dest_dir 
 */
function writefileEnd(dest_dir) {
    fs.appendFileSync(dest_dir, "};", { encoding: 'utf8', mode: 438 /*=0666*/, flag: 'a' });
}

/**
 * 获取文件名字
 * @param {*} pa 
 */
function get_file_name(pa) {
    var name = path.basename(pa);
    return name.substring(0, name.indexOf("."));
}