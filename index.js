var fs = require('fs');
var xlsx = require('xlsx');
var cvcsv = require('csv');
var _ = require('lodash');

module.exports = {
  toJson:function(excel_file,output,callback,useDictionary)
  {
      if(!excel_file) 
      {
        console.error("params error...");
        process.exit(1);
      }
      _toJson(excel_file,output,callback,useDictionary);
  },
  sheetsToJson(excel_file,sheets,output,callback,useDictionary)
  {
    if(!excel_file||!(sheets instanceof(Array)))
    {
      console.error("params error...");
      process.exit(1);
    }
    _sheetToJson(excel_file,sheets,output,callback,useDictionary);
  },
  jsonToLua:function(jsonObject,output)
  {
    var stream = fs.createWriteStream(output, { flags : 'w' });
    stream.write(_toLua(jsonObject));
  }
};


function _toJson(excel_file,output,callback,useDictionary) 
{ 
  var file = xlsx.readFile(excel_file);
  console.log(file.SheetNames);
  file.SheetNames.forEach((sheetName,index)=>{
    parse(xlsx.utils.make_csv(file.Sheets[sheetName]), output?(output+sheetName+".json"):null, callback,useDictionary,sheetName);
  });
}

function _sheetToJson(excel_file,sheets,output,callback,useDictionary)
{
  var file = xlsx.readFile(excel_file);
  console.log(file.SheetNames);
  file.SheetNames.forEach((sheetName,index)=>{
    if(sheets.indexOf(sheetName)!=-1)
    {
      parse(xlsx.utils.make_csv(file.Sheets[sheetName]), output?(output+sheetName+".json"):null, callback,useDictionary,sheetName);
    }
  });
}

function _toLua(obj) {
    'use strict';
    if (obj === null || obj === undefined) {
        return "nil";
    }
    if (!_.isObject(obj)) {
        if (typeof obj === 'string') {
            return '"' + obj + '"';
        }
        return obj.toString();
    }
    var result = "{";
    var isArray = obj instanceof Array;
    var len = _.size(obj);
    var i = 0;
    _.forEach(obj, function(v, k) {
        if (isArray) {
            result += _toLua(v);
        } else {
            result += '["' + k + '"] = ' + _toLua(v);
        }
        if (i < len - 1) {
            result += ",";
        }
        ++i;
    });
    result += "}";
    return result;
}

function parse(csv, output, callback,useDictionary,sheetName) 
{
  var record = [];
  var header = [];

  if(useDictionary == true)
  {
    record = {}
  }

  cvcsv()
    .from.string(csv)
    .transform( function(row){
      row.unshift(row.pop());
      return row;
    })
    .on('record', function(row, index){
      if(index===1)return;
      if(index === 0) {
        header = [];
        row.forEach((c,i)=>{
          if(/^[A-Za-z]+$/.test(c))
          {
            header[i]=row[i];
          }
        });
        //console.log(header);

      }else{
        //console.log("ox:"+row)
        if(row[1].trim().length>0)
        {

          var obj = {};
          header.forEach(function(column, index) {
            var v=row[index].trim();
            var isnum=/^\d+(?=\.{0,1}\d+$|$)/;
            if(v.indexOf(',')!=-1)
            {
              v=v.split(',');
              v.forEach((item,ind)=>{
                if(isnum.test(item))
                {
                  v[ind]=Number(item);
                }
              })
            }
            else if(isnum.test(v))
            {
              //console.log("nubmer:"+v);
              v=Number(v);
            }
            obj[column.trim()] = v;
          })
          useDictionary==true?record[row[1].trim()]=obj:record.push(obj);
        }
      }
    })
    .on('end', function(count){
      if(output !== null) {
        var stream = fs.createWriteStream(output, { flags : 'w' });
        stream.write(JSON.stringify(record));
        callback(null, record,sheetName);
      } else {
        callback(null, record,sheetName);
      }
      
    })
    .on('error', function(error){
      callback(error, null,sheetName);
    });
}
