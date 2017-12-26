# xlsx-2-json

##1.Excel文件配置规范

>1.字段名*(key)*必须为英文字母
>
>2.字段名*(key)*在第一行
>
>3.第二行是字段描述，不做处理
>
>4.辅助表格字段名字用中文，就不会生成在*json*里
>
>5.第一列必须为必填字段，必须有值*（如:Id）*，没有值不会生成在json里
>
>6.遍历xlsx文件里的sheet，在output目录输出以sheet name为名字的json文件

##2.脚本调用：
例：在build.js里
```
	const xlsx2json = require("xlsx-gen-json");
	xlsx2json.toJson("./excel/gameconfig.xlsx","./src/config/",function(e,r){
		e?console.log(e):console.log(r);
	});
```
