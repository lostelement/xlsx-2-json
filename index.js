var fs = require('fs');
var xlsx = require('xlsx');
var cvcsv = require('csv');

module.exports = {
  toJson:function(excel_file,output,callback)
  {
      if(!excel_file||!output) 
      {
        console.error("params error...");
        process.exit(1);
      }
      _toJson(excel_file,output,callback);
  }
};


function _toJson(excel_file,output,callback) 
{ 
  var file = load_xlsx_file(excel_file);
  file.Sheets.forEach(sheet,index)=>{
    parse(toCsv(sheet), output+file.SheetNames[index]+".json", callback);
  }
}

function load_xlsx_file = function(input) {
  return xlsx.readFile(input);
}


function toCsv(file) {
  return xlsx.utils.make_csv(file);
}

function parse(csv, output, callback) 
{
  var record = [];
  var header = [];

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
          record.push(obj);
        }
      }
    })
    .on('end', function(count){
      // when writing to a file, use the 'close' event
      // the 'end' event may fire before the file has been written
      if(output !== null) {
        var stream = fs.createWriteStream(output, { flags : 'w' });
        stream.write(JSON.stringify(record));
        callback(null, record);
      } else {
        callback(null, record);
      }
      
    })
    .on('error', function(error){
      callback(error, null);
    });
}
