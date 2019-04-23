var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');
var moment = require('moment');
var config = require('../config.json');

var arraySeparator;

module.exports = {

  noArgRun: function (dir) {
    var config = require('../config.json');
    var glob = require('glob');
    var _xlsx = require('./xlsx-to-json')

    glob(dir + "/**/[^~$]*.xlsx", function (err, files) {
      if (err) {
        console.error("exportJson error:", err);
        throw err;
      }

      files.forEach(function (element, index, array) {
        _xlsx.toJson(path.join('', element), path.join(dir, "/json"), config.xlsx.head, config.xlsx.arraySeparator);
      });

    });
  },

  toJson: function (src_excel_file, sheet_name, dest_dir, head, separator,callback) {
    arraySeparator = separator;

    if (!fs.existsSync(dest_dir)) {
      fs.mkdirSync(dest_dir);
    }

    console.log("parsing excel:", src_excel_file);
    var excel = xlsx.parse(src_excel_file);

    let len = sheet_name.length <= 0?1:excel.worksheets.length;
    _toJson(excel, sheet_name, dest_dir, len, head,callback);
  },

};

function _toJson(excel, sheet_name, dest_dir, len, head,callback) {

  for (var i_sheet = 0; i_sheet < len; i_sheet++) {
    var sheet = excel.worksheets[i_sheet];
    console.log("name:" + sheet.name + " cols:" + sheet.maxCol, "rows:" + sheet.maxRow);

    if (sheet.name == sheet_name || len <= 1) {
      if (sheet.name.indexOf('@') == -1 && sheet.name[0] !== "#") {
        var output;
        try {
          output = _parseSheet(sheet, head);
        } catch (error) {
          alert(src_excel_file + "表解析出错,请检查第"+error+"行");
        }
        // scan rest sheets for external keys
        for (var j = i_sheet; j < excel.worksheets.length; j++) {
          var rest_sheet = excel.worksheets[j];
          var rest_sheet_name = rest_sheet.name;
          if ((rest_sheet_name + '').indexOf('@') != -1) {
            var temp = rest_sheet_name.split('@');
            var external_key = temp[0];
            var sheet_name = temp[1];
            var external_values = _parseSheet(rest_sheet, head);
            if (sheet.name === sheet_name) {
              for (var k in output) {
                if (output[k] && external_values[k]) {
                  output[k][external_key] = external_values[k];
                }
              }
            }
          }
        }

        //因为输出大小的限制,这里不再对数据进行填充字符了
        output = JSON.stringify(output, null, 0);
        // output = sheet.name + ":" + output;
        console.log("output: "+output);

        var dest_file = path.resolve(dest_dir, sheet.name + ".json");
        console.log("dest_file: "+dest_file);
        fs.writeFile(dest_file, output, function (err) {
          if (err) {
            console.log("error：", err);
            alert("转换过程中出现错误==>"+err);
            throw err;
          }
  
          callback && callback();
          console.log('完成===>dest_file')
        });  
      }

    }
  }
}

var flag = 0;
function name(str) {
  while (flag <= str.length) {
    var index = str.indexOf(':');
    if ( index > -1) {
      var temp = str.substring(flag, index);
      console.log("temp= "+temp+"| index= "+index);
      flag = index;
      name(temp);
    }
  }
}

/**
 * util method for detect object is an array
 * @param o
 * @returns {boolean}
 */
function isArray(o) {
  return Object.prototype.toString.call(o) === '[object Array]';
}

var OUTPUT_ARRAY = 0;
var OUTPUT_OBJ_VALUE = 1;
var OUTPUT_OBJ_ARRAY = 2;

/**
 * parse one sheet and return the result as a json object or array
 *
 * @param sheet
 * @param head
 * @returns {*}
 * @private
 */
function _parseSheet(sheet, head) {
  //从这里开始解析表数据
  if (sheet.data && sheet.data.length > 0) {

    var row_head = sheet.data[head - 1];

    var col_type = []; //column data type
    var col_name = []; //column name
    var objOutput = OUTPUT_ARRAY;

    //读取表头 解读列名字和列数据类型
    //parse headline to get column name & column data type
    for (var i_cell = 0; i_cell < row_head.length; i_cell++) {
      var name = row_head[i_cell].value;
	  console.log("xxxx:",i_cell,row_head[i_cell].value);
      if (typeof name == 'undefined' || !name) {
        break;
      }

      var type = 'basic';

      if ((name + '').indexOf('#') != -1) {
        var temp = name.split('#');
        name = temp[0];
        type = temp[1];
        if (type) {
          type = trim(type);
          // if there exists an id type, change the whole output to a json object
          if (type === 'id') {
            objOutput = OUTPUT_OBJ_VALUE;
          } else if (type === 'id[]') {
            objOutput = OUTPUT_OBJ_ARRAY;
          }
        }
      }

      col_type.push(type);
      col_name.push(name);
    }

    var output = objOutput ? {} : [];

    for (var i_row = head; i_row < sheet.maxRow; i_row++) {
      var row = sheet.data[i_row];
      //if (typeof row == 'undefined' || !row[0]) {
      if (typeof row == 'undefined') {
        throw i_row;
        break;
      }
      // var json_obj = {};
      var json_obj = [];
      var id;

      for (var i_col = 0; i_col < col_type.length; i_col++) {
        var cell = row[i_col];
        var type = col_type[i_col].toLowerCase().trim();
        //console.log(col_name[i_col],cell.value,typeof(cell),type);

        // if the cell is empty, do not export the key

        switch (type) {
          case 'id': // id
            if (cell) {
              id = cell.value;
            }
            break;
          case 'id[]': // id[]
            if (cell) {
              id = cell.value;
              if (!output[id]) {
                output[id] = [];
              }
            }
            break;
          case 'basic': // number string boolean date
            if (cell) {
              if (isDateType(cell.value)) {
                parseDateType(json_obj, col_name[i_col], cell.value);
              } else {
                var _t;
                if(cell.value){
                  _t = cell.value;
                }else{
                  _t = null;
                }
                json_obj.push(_t);
                // json_obj[col_name[i_col]] = cell.value;
              }
            }else{
              json_obj.push(null);
            }
            break;
          case 'date':
            parseDateType(json_obj, col_name[i_col], cell.value);
            break;
          case 'string':
            if (cell) {
              if (isDateType(cell.value)) {
                parseDateType(json_obj, col_name[i_col], cell.value);
              } else {
                if (cell.value === null || cell.value.toString() === 'null') {
                  json_obj.push("");
                  // json_obj[col_name[i_col]] = "";
                } else {
                  json_obj.push(cell.value.toString());
                  // json_obj[col_name[i_col]] = cell.value.toString();
                }
              }
            } else {
              json_obj.push("");
              // json_obj[col_name[i_col]] = "";
            }
            break;
          case 'number':
            //+xxx.toString() '+' means convert it to number
            var isNumber = !isNaN(+cell.value.toString());
            if (cell && isNumber) {
              json_obj.push(Number(cell.value));
              // json_obj[col_name[i_col]] = Number(cell.value);
            } else {
              json_obj.push(0);
              // json_obj[col_name[i_col]] = 0;
            }
            if (!isNumber) {
              console.log("cell[" + (i_row + 1) + "," + (i_col + 1) + "]: not a number");
            }
            break;
          case 'bool':
            if (cell) {
              json_obj[col_name[i_col]] = toBoolean(cell.value.toString());
            } else {
              json_obj.push(false)
              // json_obj[col_name[i_col]] = false;
            }
            break;
          case '{}': //support {number boolean string date} property type
            if (cell && cell.value && cell.value.trim().length) {
              parseObject(json_obj, col_name[i_col], cell.value);
            }else{
              json_obj.push(null);
            }
            break;
          case '[]': //[number] [boolean] [string]  todo:support [date] type
            if (cell && cell.value) {
              parseBasicArrayField(json_obj, col_name[i_col], cell.value);
            }else{
              json_obj.push([]);
            }
            break;
          case '[{}]':
            if (cell && cell.value && cell.value.trim().length) {
              parseObjectArrayField(json_obj, col_name[i_col], cell.value);
            } else {
              json_obj.push([]);
              // json_obj[col_name[i_col]] = [];
            }
            break;
          default:
            console.log('********************************************');
            console.log('unrecognized type', cell.value, typeof (cell.value));
            alert('unrecognized type', cell.value, typeof (cell.value));
            break;
        }

        //console.log('********************************************');
        //console.log("--->",i_row,type,json_obj[col_name[i_col]], col_name[i_col],cell);

      }
      if (objOutput === OUTPUT_OBJ_VALUE) {
        output[id] = json_obj;
      } else if (objOutput === OUTPUT_OBJ_ARRAY) {
        output[id].push(json_obj);
      } else if (objOutput === OUTPUT_ARRAY) {
        output.push(json_obj);
      }
    }
	  //console.log("output******",output);
    return output;
  }
}

/**
 * parse date type
 * row:row of xlsx
 * key:col of the row
 * value: cell value
 */
function parseDateType(row, key, value) {
  rwo.push(convert2Date(value));
  // row[key] = convert2Date(value);
  //console.log(value,key,row[key]);
}

/**
 * convert string to date type
 * value: cell value
 */
function convert2Date(value) {
  var dateTime = moment(value);
  if (dateTime.isValid()) {
    return dateTime.format("YYYY-MM-DD HH:mm:ss");
  } else {
    return "";
  }
}

/**
 * parse object array.
 */
function parseObjectArrayField(row, key, value) {

  var obj_array = [];

  if (value) {
    if (value.indexOf(',') !== -1) {
      obj_array = value.split(',');
    } else {
      obj_array.push(value.toString());
    }
  }

  // if (typeof(value) === 'string' && value.indexOf(',') !== -1) {
  //     obj_array = value.split(',');
  // } else {
  //     obj_array.push(value.toString());
  // };

  // var result = [];
  
  // result.push(array2object(obj_array));

  // obj_array.forEach(function (e) {
  //   if (e) {
  //     result.push(array2object([e]));
  //   }
  // });
  //新版不使用key了,这里直接push到数组里边
  // row[key] = result;
  row.push(array2object(obj_array));
}

/**
 * parse object from array.
 *  for example : [a:123,b:45] => {'a':123,'b':45}
 */
function array2object(array) {
  var result = [];
  array.forEach(function (e) {
    if (e) {
      var kv = e.trim().split(':');
      if (isNumber(kv[1])) {
        kv[1] = Number(kv[1]);
      } else if (isBoolean(kv[1])) {
        kv[1] = toBoolean(kv[1]);
      } else if (isDateType(kv[1])) {
        kv[1] = convert2Date(kv[1]);
      }
      // 新版不使用key了,这里直接push到数组里边
      // result[kv[0]] = kv[1];
      result.push(kv);
    }
  });
  return result;
}

/**
 * parse object
 */
function parseObject(field, key, data) {
  field[key] = array2object(data.split(';'));
}


/**
 * parse simple array.
 */
function parseBasicArrayField(field, key, array) {
  var basic_array;

  if (typeof array === "string") {
    basic_array = array.split(arraySeparator);
  } else {
    basic_array = [];
    basic_array.push(array);
  }

  var result = [];
  if (isNumberArray(basic_array)) {
    basic_array.forEach(function (element) {
      result.push(Number(element));
    });
  } else if (isBooleanArray(basic_array)) {
    basic_array.forEach(function (element) {
      result.push(toBoolean(element));
    });
  } else { //string array
    result = basic_array;
  }
  // console.log("basic_array", result + "|||" + cell.value);
  field.push(result);
  // field[key] = result;
}

/**
 * convert value to boolean.
 */
function toBoolean(value) {
  return value.toString().toLowerCase() === 'true';
}

/**
 * is a boolean array.
 */
function isBooleanArray(arr) {
  return arr.every(function (element, index, array) {
    return isBoolean(element);
  });
}

/**
 * is a number array.
 */
function isNumberArray(arr) {
  return arr.every(function (element, index, array) {
    return isNumber(element);
  });
}

/**
 * is a number.
 */
function isNumber(value) {

  if (typeof (value) == "undefined") {
    return false;
  }

  if (typeof value === 'number') {
    return true;
  }
  return !isNaN(+value.toString());
}

/**
 * boolean type check.
 */
function isBoolean(value) {

  if (typeof (value) == "undefined") {
    return false;
  }

  if (typeof value === 'boolean') {
    return true;
  }

  var b = value.toString().trim().toLowerCase();

  return b === 'true' || b === 'false';
}

//delete all space
function trim(str) {
  return str.replace(/(^\s+)|(\s+$)/g, "");
}

/**
 * date type check.
 */
function isDateType(value) {
  if (value) {
    var str = value.toString();
    return moment(new Date(value), "YYYY-M-D", true).isValid() || moment(value, "YYYY-M-D H:m:s", true).isValid() || moment(value, "YYYY/M/D H:m:s", true).isValid() || moment(value, "YYYY/M/D", true).isValid();
  }
  return false;
}
