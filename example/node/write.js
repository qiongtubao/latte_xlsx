var latte_xlsx = require("../../");
var Fs = require("fs");
var data = {
	sheets: ["Sheet1"],
	Sheet1: [
		["a", "b", "c", "d"],
		[1 , 2 , 3, 4]
	]
};
latte_xlsx.write(data, function(err, buffer) {
	Fs.writeFileSync("b.xlsx", new Buffer(buffer));
	console.log("ok");
});