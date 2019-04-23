var xlsx = require('./lib/xlsx-to-json.js');
var g_path = require('path');
var shell = require('child_process');
var fs = require('fs');
var glob = require('glob');
var config = require('./config.json');

/**
 * all commands
 */
var commands = {
    // 帮助
    "--help": {
    "alias": ["-h"],
    "desc": "show this help manual.",
    "action": showHelp
  },
    // 导出指定excel的指定sheet
    "--export": {
    "alias": ["-e"],
    "desc": "export excel to json. --export [files]",
    "action": exportJson
  },
    // 导出指定路径下的所有excel
    "--all": {
    "alias": ["-a"],
    "desc": "export excel to json. --export [files]",
    "action": exportAllJson,
    "default": true
  }
};

 //cache of command's key ("--help"...)
var keys = Object.keys(commands);

var alias_map = {}; // mapping of alias_name -> name
for (var key in commands) {
    var alias_array = commands[key].alias;
    alias_array.forEach(function (e) {
    alias_map[e] = key;
  });
}

var parsed_cmds = []; //cmds of parsed out.
parsed_cmds = parseCommandLine(process.argv);
parsed_cmds.forEach(function (e) {
  exec(e);
});

function checkExcel(path) {
  var suffix = path.substr(path.lastIndexOf("."));
  if (suffix == ".xlsx") {
    return true;
  }
  
  console.error("路径不是excel文件："+path);
  return false;
}

function checkDir(path) {
  var stat = fs.statSync(path);
  if (stat.isDirectory()) {
    return true;
  }

  console.error("路径不是文件夹："+path);
  return false;
}

/**
 * export all json
 * @param {} args 
 */
function exportAllJson(args) {
  if (args instanceof Array) {

    if(args.length == 2){
      var src = args[0];
      var dst = args[1];

      //是文件 存在 是excel表
      if (!checkDir(src)) {
        return;
      }

      //是文件夹 存在
      if (!checkDir(src)) {
        return;
      }

      glob(src+"/**/[^~$]*.xlsx", function (err, files) {
        if (err) {
          console.error("exportJson error:", err);
          throw err;
        }
        
        files.forEach(function (element, index, array) {
          console.log(index+":"+element);
          xlsx.toJson(element,"", dst, config.xlsx.head, config.xlsx.arraySeparator);
        });
  
      });

    } else {
      console.log("参数错误请检查，参数长度："+args.length);
    }

  }
}


/**
 * export json
 * args: --export [cmd_line_args] [.xlsx files list].
 */
function exportJson(args) {

  if (args instanceof Array) {

    if(args.length == 3){
      var sheet = args[0];
      var excel_file = args[1];
      var dst = args[2];

      //是文件 存在 是excel表
      if (!checkExcel(excel_file)) {
        return;
      }

      //是文件夹 存在
      if (!checkDir(dst)) {
        return;
      }      

      if (excel_file.indexOf(".xlsx") != -1) {
        xlsx.toJson(excel_file, sheet, dst, config.xlsx.head, config.xlsx.arraySeparator);
      } else {
        console.log("参数错误请检查，后缀错误");
      }

    } else {
      console.log("参数错误请检查，参数长度："+args.length);
    }

  }
}

function isEmptyArr(arr) {
  if(arr == undefined || arr.length == 0) {
    return true;
  }
  return false;
}

/**
 * show help
 */
function showHelp() {
  var usage = "usage: \n";
  for (var p in commands) {
    if (typeof commands[p] !== "function") {
      usage += "\t " + p + "\t " + commands[p].alias + "\t " + commands[p].desc + "\n ";
    }
  }

  usage += "\nexamples: ";
  usage += "\n\n $node index.js --export\n\tthis will export all files configed to json.";
  usage += "\n\n $node index.js --export ./excel/foo.xlsx ./excel/bar.xlsx\n\tthis will export foo and bar xlsx files.";

  console.log(usage);
}


/**************************** parse command line *********************************/

/**
 * execute a command
 */
function exec(cmd) {
  if (typeof cmd.action === "function") {
    cmd.action(cmd.args);
  }
}


/**
 * parse command line args
 */
function parseCommandLine(args) {
  var parsed_cmds = [];

  if (args.length < 2) {
    alert('参数错误，至少输入两个参数');
  } else {
    var cli = args.slice(2);

    var pos = 0;
    var cmd;

    cli.forEach(function (element, index, array) {
      //replace alias name with real name.
      if (element.indexOf('--') === -1 && element.indexOf('-') === 0) {
        cli[index] = alias_map[element];
      }

      //parse command and args
      if (cli[index].indexOf('--') === -1) {
        cmd.args.push(cli[index]);
      } else {

        if (keys[cli[index]] == "undefined") {
          throw new Error("not support command:" + cli[index]);
        }

        pos = index;
        cmd = commands[cli[index]];
        if (typeof cmd.args == 'undefined') {
          cmd.args = [];
        }
        parsed_cmds.push(cmd);
      }
    });
  }

  return parsed_cmds;
}

/**
 * default command when no command line argas provided.
 */
function defaultCommand() {
  if (keys.length <= 0) {
    throw new Error("Error: there is no command at all!");
  }

  for (var p in commands) {
    if (commands[p]["default"]) {
      return commands[p];
    }
  }

  if (keys["--help"]) {
    return commands["--help"];
  } else {
    return commands[keys[0]];
  }
}

/*************************************************************************/
