var latte_xlsx = require("../../");
var Fs = require("fs");
var data = Fs.readFileSync("./a.xlsx");
latte_xlsx.read(data, function(err, data) {
	console.log(err, data);
});