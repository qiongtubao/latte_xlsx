
		var latte_lib = require("latte_lib")
			, Async = latte_lib.async
			, Zip = require("latte_zip")
			, utf8 = latte_lib.utf8;
		function Xlsx() {
			this.info = {};
		};
		(function() {
			function unescapeXML(s) { 
				return (s || '')
				.replace(/&amp;/g, '&')
				.replace(/&lt;/g, '<')
				.replace(/&gt;/g, '>')
				.replace(/&quot;/g, '"')
				.replace(/&#x27;/g, '\''); 
			}
			this.readSharedStrings = function(callback) {
				var self = this;
				this.zip.getFile("xl/sharedStrings.xml", function(error, xml) {
					if(error) {
						return callback(error);
					}
					var sharedStrings = self.info.sharedStrings = [];
					if(xml) {
						var sis = xml.split("<si>");
						var i = sis.length;
						while(--i) {
							var ts = sis[i].split(/<t.*?>/g);
							var tl = ts.length;
							var text = "";
							for(var ti = 0; ti < tl; ti++){
								text += unescapeXML(ts[ti].substring(0, ts[ti].indexOf("</t>")));
							}
							sharedStrings[i - 1] = utf8.decode(text);
						}
					}
					callback(error, 1);
				});
			}
			function getOneXmlText(data, xmlName, end) {
				var data = data.substr(data.indexOf(xmlName)+(xmlName.length+1));
				return data.substring(0,data.indexOf((end || xmlName))-2);
			}
			var getAttr = function(s, n) { 
				s = s.substr(s.indexOf(n + '="') + n.length + 2); 
				return s.substring(0, s.indexOf('"')); 
			}
			this.readCore = function(callback) {
				var self = this;
				this.zip.getFile("docProps/core.xml", function(error, xml) {
					if(error) {
						return callback(error);
					}
					self.info.creator = getOneXmlText(xml, "dc:creator");
					self.info.lastModifiedBy = getOneXmlText(xml,"cp:lastModifiedBy");
					self.info.created = new Date(getOneXmlText(xml, "dcterms:created xsi:type=\"dcterms:W3CDTF\"","dcterms:modified"));
					self.info.modified = new Date(getOneXmlText(xml , "dcterms:modified xsi:type=\"dcterms:W3CDTF\"","dcterms:modified"));
					
					callback(error, 1);
				});
			}
			this.readWorkBook = function(callback) {
				var self = this;
				this.info.sheets = [];
				this.zip.getFile("xl/workbook.xml", function(error, xml) {
					if(error) {
						return callback(error);
					}
					self.info.worksheets = [];
					var matchs = xml.match("activeTab=\"([0-9]*)\"");
					if(matchs) {
						self.info.activeWorksheet = parseInt(matchs[1]);
					}else{
						self.info.activeWorksheet = 0;
					}

					var sheets = xml.split("<sheet ");
					var i = sheets.length;
					while(--i) {
						var sheet = sheets[i].substring(sheets[i].indexOf("name"), sheets[i].indexOf("/>"));
						var name = getAttr(sheet, "name");
						if(name) {
							self.info.worksheets.unshift({ name: name, data: []});
							var id = getAttr(sheet, "id");
							var rId = getAttr(sheet, "r:id");
							self.info.sheets[rId.substring(3) - 1] = name;
						}
					}
					callback(error, 1);
				});
			}
			this.readStyle = function( callback) {
				var self = this;
				this.zip.getFile("xl/styles.xml", function(error, xml){
					if(error) {
						return callback(error);
					}
					var numFmts = xml.split("<numFmt ");
					var numFmtDatas = ['General', '0', '0.00', '#,##0', '#,##0.00',,,,, '0%', '0.00%', '0.00E+00', '# ?/?', '# ??/??', 'mm-dd-yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM',
					'h:mm', 'h:mm:ss', 'm/d/yy h:mm',,,,,,,,,,,,,,, '#,##0 ;(#,##0)', '#,##0 ;[Red](#,##0)', '#,##0.00;(#,##0.00)', '#,##0.00;[Red](#,##0.00)',,,,, 'mm:ss', '[h]:mm:ss', 'mmss.0', '##0.0E+0', '@'];
					var i = numFmts.length;
					while(--i) {
						var t = numFmts[i];
						numFmtDatas[+getAttr(t, "numFmtId")] = getAttr(t, "formatCode");
					}
					var cellXfsDatas = xml.substr(xml.indexOf("cellXfs")).split("<xf ");
					var styles = [];
					i = cellXfsDatas.length;
					while(--i) {
						var id = getAttr(cellXfsDatas[i], "numFmtId");
						var f = numFmtDatas[id];
						if(f.indexOf("m") > -1) {
							t = "date";
						}else if(f.indexOf("0") > -1) {
							t = "number";
						}else if(f === "@") {
							t = "string";
						}else{
							t = "unknown";
						}
						styles.unshift({formatCode: f, type: t});
					}
					self.info.styles = styles;
					callback(error, 1);
				});
			}
				var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
				function alphaNum(s) { 
					var t = 0; 
					if (s.length === 2) { 
						t = alphaNum(s.charAt(0)) + 1; 
					} 
					return t * 26 + alphabet.indexOf(s.substr(-1)); 
				}
			this.readSheetDatas = function(callback) {
				var self = this;
				var worksheets = this.info.worksheets;
				var len = worksheets.length;
				var styles = this.info.styles
					, sharedStrings = this.info.sharedStrings;
				var funcs = [];
				var debugData = {};
				this.works = {};
				while(len--) {
					(function(i) {
						funcs.push(function(callback) {
							self.zip.getFile("xl/worksheets/sheet"+(i+1)+".xml", function(error, xml) {
								if(error) {
									return callback(error);
								}
								var rowDatas = xml.split("<row ");
								var w = worksheets[i];
								w.name = utf8.decode(w.name);
								self.works[w.name] = [];
								w.table = rowDatas[0].indexOf("<tableParts ") > 0;
								var refs = getAttr(rowDatas[0].substr(rowDatas[0].indexOf("<dimension")), "ref").split(":");
								var t = refs[1];
								if(!t) {
									return callback(error, 1);
								}
								w.maxCol = alphaNum(t.match(/[a-zA-Z]*/g)[0]) +1;
								w.maxRow = +t.match(/\d*/g).join("");
								var j = rowDatas.length;
								while(--j) {
									var row = w.data[+getAttr(rowDatas[j], "r") - 1] = [];
									var columns = rowDatas[j].split("<c ");
									var k = columns.length;
									self.works[w.name][+getAttr(rowDatas[j], "r") -1] = [];
									while(--k) {
										var cell = columns[k];
										var f = styles[+getAttr(cell, "s")] || {type: "General", formatCode: "General" };
										var t = getAttr(cell, "t") || f.type;
										var val = cell.substring(cell.indexOf("<v>") + 3, cell.indexOf("</v>"));
										val = val ? + val: "";
										val = isNaN(val)? "": val;
										switch(t) {
											case "s": val = sharedStrings[val]; break;
											case "b": val = val === 1;break;
											case "date": val = convertDate(val); break;
										}
										row[alphaNum(getAttr(cell, "r").match(/[a-zA-Z]*/g)[0])] = {value: val, formatCode: f.formatCode};
										self.works[w.name][+getAttr(rowDatas[j], "r") -1][alphaNum(getAttr(cell,"r").match(/[a-zA-Z]*/g)[0])] = val;
									}	

								}
								callback(error, 1);
							});
						});
					})(len);
				};
				Async.parallel(funcs, function(err) {
					callback(err, 1);
				});

			}
			this.read = function(data, callback) {
				this.xlsxData = data;
				this.zip = Zip.create(data);
				var self = this;
				Async.parallel([
					function(cb) {
						self.readSharedStrings(cb);
					},
					function(cb) {
						self.readCore(cb);
					},
					function(cb) {
						self.readWorkBook(cb);
					},
					function(cb) {
						self.readStyle(cb);
					}				
				], function() {
					self.readSheetDatas(function(error) {
						callback(error, self);
					});
				});
			}

			this.writeTheme1 = function() {
				this.zip.updateEntry("xl/theme/theme1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
						<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">\
							<a:themeElements>\
								<a:clrScheme name="Office">\
									<a:dk1>\
										<a:sysClr val="windowText" lastClr="000000"/>\
									</a:dk1>\
									<a:lt1>\
										<a:sysClr val="window" lastClr="FFFFFF"/>\
									</a:lt1>\
									<a:dk2>\
										<a:srgbClr val="1F497D"/>\
									</a:dk2>\
									<a:lt2>\
										<a:srgbClr val="EEECE1"/>\
									</a:lt2>\
									<a:accent1>\
										<a:srgbClr val="4F81BD"/>\
									</a:accent1>\
									<a:accent2>\
										<a:srgbClr val="C0504D"/>\
									</a:accent2>\
									<a:accent3>\
										<a:srgbClr val="9BBB59"/>\
									</a:accent3>\
									<a:accent4>\
										<a:srgbClr val="8064A2"/>\
									</a:accent4>\
									<a:accent5>\
										<a:srgbClr val="4BACC6"/>\
									</a:accent5>\
									<a:accent6>\
										<a:srgbClr val="F79646"/>\
									</a:accent6>\
									<a:hlink>\
										<a:srgbClr val="0000FF"/>\
									</a:hlink>\
									<a:folHlink>\
										<a:srgbClr val="800080"/>\
									</a:folHlink>\
								</a:clrScheme>\
								<a:fontScheme name="Office">\
									<a:majorFont>\
										<a:latin typeface="Cambria"/>\
										<a:ea typeface=""/>\
										<a:cs typeface=""/>\
										<a:font script="Jpan" typeface="MS P????"/>\
										<a:font script="Hang" typeface="?? ??"/>\
										<a:font script="Hans" typeface="??"/>\
										<a:font script="Hant" typeface="????"/>\
										<a:font script="Arab" typeface="Times New Roman"/>\
										<a:font script="Hebr" typeface="Times New Roman"/>\
										<a:font script="Thai" typeface="Tahoma"/>\
										<a:font script="Ethi" typeface="Nyala"/>\
										<a:font script="Beng" typeface="Vrinda"/>\
										<a:font script="Gujr" typeface="Shruti"/>\
										<a:font script="Khmr" typeface="MoolBoran"/>\
										<a:font script="Knda" typeface="Tunga"/>\
										<a:font script="Guru" typeface="Raavi"/>\
										<a:font script="Cans" typeface="Euphemia"/>\
										<a:font script="Cher" typeface="Plantagenet Cherokee"/>\
										<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>\
										<a:font script="Tibt" typeface="Microsoft Himalaya"/>\
										<a:font script="Thaa" typeface="MV Boli"/>\
										<a:font script="Deva" typeface="Mangal"/>\
										<a:font script="Telu" typeface="Gautami"/>\
										<a:font script="Taml" typeface="Latha"/>\
										<a:font script="Syrc" typeface="Estrangelo Edessa"/>\
										<a:font script="Orya" typeface="Kalinga"/>\
										<a:font script="Mlym" typeface="Kartika"/>\
										<a:font script="Laoo" typeface="DokChampa"/>\
										<a:font script="Sinh" typeface="Iskoola Pota"/>\
										<a:font script="Mong" typeface="Mongolian Baiti"/>\
										<a:font script="Viet" typeface="Times New Roman"/>\
										<a:font script="Uigh" typeface="Microsoft Uighur"/>\
										<a:font script="Geor" typeface="Sylfaen"/>\
									</a:majorFont>\
									<a:minorFont>\
										<a:latin typeface="Calibri"/>\
										<a:ea typeface=""/>\
										<a:cs typeface=""/>\
										<a:font script="Jpan" typeface="MS P????"/>\
										<a:font script="Hang" typeface="?? ??"/>\
										<a:font script="Hans" typeface="??"/>\
										<a:font script="Hant" typeface="????"/>\
										<a:font script="Arab" typeface="Arial"/>\
										<a:font script="Hebr" typeface="Arial"/>\
										<a:font script="Thai" typeface="Tahoma"/>\
										<a:font script="Ethi" typeface="Nyala"/>\
										<a:font script="Beng" typeface="Vrinda"/>\
										<a:font script="Gujr" typeface="Shruti"/>\
										<a:font script="Khmr" typeface="DaunPenh"/>\
										<a:font script="Knda" typeface="Tunga"/>\
										<a:font script="Guru" typeface="Raavi"/>\
										<a:font script="Cans" typeface="Euphemia"/>\
										<a:font script="Cher" typeface="Plantagenet Cherokee"/>\
										<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>\
										<a:font script="Tibt" typeface="Microsoft Himalaya"/>\
										<a:font script="Thaa" typeface="MV Boli"/>\
										<a:font script="Deva" typeface="Mangal"/>\
										<a:font script="Telu" typeface="Gautami"/>\
										<a:font script="Taml" typeface="Latha"/>\
										<a:font script="Syrc" typeface="Estrangelo Edessa"/>\
										<a:font script="Orya" typeface="Kalinga"/>\
										<a:font script="Mlym" typeface="Kartika"/>\
										<a:font script="Laoo" typeface="DokChampa"/>\
										<a:font script="Sinh" typeface="Iskoola Pota"/>\
										<a:font script="Mong" typeface="Mongolian Baiti"/>\
										<a:font script="Viet" typeface="Arial"/>\
										<a:font script="Uigh" typeface="Microsoft Uighur"/>\
										<a:font script="Geor" typeface="Sylfaen"/>\
									</a:minorFont>\
								</a:fontScheme>\
								<a:fmtScheme name="Office">\
									<a:fillStyleLst>\
										<a:solidFill>\
											<a:schemeClr val="phClr"/>\
										</a:solidFill>\
										<a:gradFill rotWithShape="1">\
											<a:gsLst>\
												<a:gs pos="0">\
													<a:schemeClr val="phClr">\
														<a:tint val="50000"/>\
														<a:satMod val="300000"/>\
													</a:schemeClr>\
												</a:gs>\
												<a:gs pos="35000">\
													<a:schemeClr val="phClr">\
														<a:tint val="37000"/>\
														<a:satMod val="300000"/>\
													</a:schemeClr>\
												</a:gs>\
												<a:gs pos="100000">\
													<a:schemeClr val="phClr">\
														<a:tint val="15000"/>\
														<a:satMod val="350000"/>\
													</a:schemeClr>\
												</a:gs>\
											</a:gsLst>\
											<a:lin ang="16200000" scaled="1"/>\
										</a:gradFill>\
										<a:gradFill rotWithShape="1">\
											<a:gsLst>\
												<a:gs pos="0">\
													<a:schemeClr val="phClr">\
														<a:shade val="51000"/>\
														<a:satMod val="130000"/>\
													</a:schemeClr>\
												</a:gs>\
												<a:gs pos="80000">\
													<a:schemeClr val="phClr">\
														<a:shade val="93000"/>\
														<a:satMod val="130000"/>\
													</a:schemeClr>\
												</a:gs>\
												<a:gs pos="100000">\
													<a:schemeClr val="phClr">\
														<a:shade val="94000"/>\
														<a:satMod val="135000"/>\
													</a:schemeClr>\
												</a:gs>\
											</a:gsLst>\
											<a:lin ang="16200000" scaled="0"/>\
										</a:gradFill>\
									</a:fillStyleLst>\
									<a:lnStyleLst>\
										<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">\
											<a:solidFill>\
												<a:schemeClr val="phClr">\
													<a:shade val="95000"/>\
													<a:satMod val="105000"/>\
												</a:schemeClr>\
											</a:solidFill>\
											<a:prstDash val="solid"/>\
										</a:ln>\
										<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">\
											<a:solidFill>\
												<a:schemeClr val="phClr"/>\
											</a:solidFill>\
											<a:prstDash val="solid"/>\
										</a:ln>\
										<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">\
											<a:solidFill>\
												<a:schemeClr val="phClr"/>\
											</a:solidFill>\
											<a:prstDash val="solid"/>\
										</a:ln>\
									</a:lnStyleLst>\
									<a:effectStyleLst>\
										<a:effectStyle>\
											<a:effectLst>\
											<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">\
												<a:srgbClr val="000000"><a:alpha val="38000"/>\
												</a:srgbClr>\
											</a:outerShdw>\
											</a:effectLst>\
										</a:effectStyle>\
										<a:effectStyle>\
											<a:effectLst>\
												<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">\
												<a:srgbClr val="000000">\
													<a:alpha val="35000"/>\
												</a:srgbClr>\
												</a:outerShdw>\
											</a:effectLst>\
										</a:effectStyle>\
										<a:effectStyle>\
											<a:effectLst>\
												<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">\
													<a:srgbClr val="000000">\
														<a:alpha val="35000"/>\
													</a:srgbClr>\
												</a:outerShdw>\
											</a:effectLst>\
											<a:scene3d>\
												<a:camera prst="orthographicFront">\
													<a:rot lat="0" lon="0" rev="0"/>\
												</a:camera>\
												<a:lightRig rig="threePt" dir="t">\
													<a:rot lat="0" lon="0" rev="1200000"/>\
												</a:lightRig>\
											</a:scene3d>\
											<a:sp3d>\
												<a:bevelT w="63500" h="25400"/>\
											</a:sp3d>\
										</a:effectStyle>\
									</a:effectStyleLst>\
									<a:bgFillStyleLst>\
										<a:solidFill>\
											<a:schemeClr val="phClr"/>\
										</a:solidFill>\
										<a:gradFill rotWithShape="1">\
											<a:gsLst>\
												<a:gs pos="0">\
													<a:schemeClr val="phClr">\
														<a:tint val="40000"/>\
														<a:satMod val="350000"/>\
													</a:schemeClr>\
												</a:gs>\
												<a:gs pos="40000">\
													<a:schemeClr val="phClr">\
														<a:tint val="45000"/>\
														<a:shade val="99000"/>\
														<a:satMod val="350000"/>\
													</a:schemeClr>\
												</a:gs>\
												<a:gs pos="100000">\
													<a:schemeClr val="phClr">\
														<a:shade val="20000"/>\
														<a:satMod val="255000"/>\
													</a:schemeClr>\
												</a:gs>\
											</a:gsLst>\
											<a:path path="circle">\
												<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>\
											</a:path>\
										</a:gradFill>\
										<a:gradFill rotWithShape="1">\
											<a:gsLst>\
												<a:gs pos="0">\
													<a:schemeClr val="phClr">\
														<a:tint val="80000"/>\
														<a:satMod val="300000"/>\
													</a:schemeClr>\
												</a:gs>\
												<a:gs pos="100000">\
													<a:schemeClr val="phClr">\
														<a:shade val="30000"/>\
														<a:satMod val="200000"/>\
													</a:schemeClr>\
												</a:gs>\
											</a:gsLst>\
											<a:path path="circle">\
											<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>\
											</a:path>\
										</a:gradFill>\
									</a:bgFillStyleLst>\
								</a:fmtScheme>\
							</a:themeElements>\
							<a:objectDefaults/>\
							<a:extraClrSchemeLst/>\
						</a:theme>\
					');
			}

			this.writeRels = function() {
				this.zip.updateEntry("_rels/.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
						<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
						<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\
						<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\
						<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\
					</Relationships>');
			}

			this.writeXlWorkSheets = function() {
				var opts = this.opts;
				this.zip.updateEntry("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
					<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
					<dc:creator>'
					+ (opts.creator || 'latte') 
				+ '</dc:creator><cp:lastModifiedBy>' 
					+ (opts.lastModifiedBy || 'latte') 
				+ '</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">'
					+ (opts.created || new Date()).toISOString() 
				+ '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' 
					+ (opts.modified || new Date()).toISOString() 
				+ '</dcterms:modified></cp:coreProperties>');
			}

				function escapeXML(s) { 
					return (s || '')
					.replace(/&/g, '&amp;')
					.replace(/</g, '&lt;')
					.replace(/>/g, '&gt;')
					.replace(/"/g, '&quot;')
					.replace(/'/g, '&#x27;'); 
				} 
				var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
				function alphaNum(s) { 
					var t = 0; 
					if (s.length === 2) {
						t = alphaNum(s.charAt(0)) + 1; 
					} 
					return t * 26 + alphabet.indexOf(s.substr(-1)); 
				}
				function numAlpha(i) { 
					var t = Math.floor(i / 26) - 1; 
					return (t > -1 ? numAlpha(t) : '') + alphabet.charAt(i % 26); 
				}
			this.write = function(data, opts, callback) {
				this.works = data;
				if(typeof opts == "function") {
					callback = opts;
					opts = {};
				}
				if(!opts) { opts = {}};
				this.opts = opts;
				opts.font = opts.font || "Calibri";
				opts.fontSize = opts.fontSize || 1;
				opts.fontColor = opts.fontColor || "000";
				var defaultFontName = opts.defaultFontName || "Calibri";
				var defaultFontSize = opts.defaultFontName || 11;
				var defaultFontColor = opts.defaultFontColor || "000";
				var defaultCellBorders = opts.defaultCellBorders;
				var defaultHAlign = opts.defaultHAlign;
				var defaultVAlign = opts.defaultVAlign;
				var defaultTextRotation = opts.defaultTextRotation;
				var processTime = Date.now();
				var sharedStrings = [[], 0];
				var zip = this.zip = Zip.create();
				this.writeRels();			
				this.writeTheme1();
				this.writeXlWorkSheets();
				var styles = new Array(1);
			    var borders = new Array(1);
			    var fonts = new Array(1);
				var	w = data.sheets.length;
				var contentTypes = [[], []];
				var props = [];
				var xlRels = [];
				var worksheets = [];

				while(w--) {
					var id = w+1;
					var worksheet = data[data.sheets[w]]; 
					var s = '';
					var columns = [];
			    	var merges = [];
			    	var i = -1; 
			    	var l = worksheet.length;
			    	while (++i 	< l) {
			    		var j = -1; 
						var k = worksheet[i].length;
						var height = opts.forcedRowHeight && opts.forcedRowHeight[i] >= 0 ? opts.forcedRowHeight[i] : opts.rowHeight;
						s += '<row r="' + (i + 1) + '"' + (height >= 0 ? ' customHeight="1" ht="' + height + '"' : '') + ' x14ac:dyDescent="0.25">';
			    		while (++j < k) {
			    			var cell = worksheet[i][j]; 
							var val = cell.hasOwnProperty('value') ? cell.value : cell; 
							var t = ''; 
							var style = {
								borders: cell.borders || defaultCellBorders, 
								hAlign: cell.hAlign || defaultHAlign,
								textRotation: cell.textRotation || defaultTextRotation,
								vAlign: cell.vAlign === 'bottom' ? '' : (cell.vAlign || defaultVAlign), // default is bottom
								bold: cell.bold,
								italic: cell.italic,
								fontName: cell.fontName,
								fontColor: cell.fontColor,
								fontSize: cell.fontSize,
								formatCode: cell.formatCode || 'General'
							};
							var colWidth = 0;
							if (val && typeof val === 'string' && !isFinite(val) || cell.forceString) { 
								// If value is string, and not string of just a number, place a sharedString reference instead of the value
		            			val = escapeXML(val);
								sharedStrings[1]++; // Increment total count, unique count derived from sharedStrings[0].length
								index = sharedStrings[0].indexOf(val);
								colWidth = val.length;
								if (index < 0) {
								 	index = sharedStrings[0].push(val.replace(/\n/g, '&#10;')) - 1; 
								}
								val = index;
								t = 's';
							} else if (typeof val === 'boolean') { 
								val = (val ? 1 : 0); t = 'b'; 
								colWidth = 1;
							} else if (typeof(val) === 'date') { 
								val = convertDate(val); 
								style.formatCode = cell.formatCode || 'mm-dd-yy'; 
								colWidth = val.length;
							} else if (typeof val === 'object') {
								// unsupported value
								val = null
							} else {
								// number, or string which is a number 
								colWidth = (''+val).length;
							}

							var style = JSON.stringify(style);
							var index = styles.indexOf(style);
							if (index < 0) { 
								style = styles.push(style) - 1; 
							} else { 
								style = index; 
							}
							// keeps largest cell in column, and autoWidth flag that may be set on any cell
							if (columns[j] == null) {
								columns[j] = {autoWidth: false, max:0};
							}
							if (cell.autoWidth) {
								columns[j].autoWidth = true;
							}
							if (colWidth > columns[j].max) {
								columns[j].max = colWidth;
							}
							if (cell.colWidth) {
								columns[j].autoWidth = true;
								columns[j].max = cell.colWidth;
							}

							if (cell.colSpan > 1 && cell.rowSpan > 1) {
								merges.push([numAlpha(j) + (i + 1), numAlpha(j+cell.colSpan-1) + (i + cell.rowSpan)]);
								merged = [j, 0];
								for (var m = 0; m < cell.colSpan-1; m++) {
									merged.push(cell);
								}
								data[i].splice.apply(data[i], merged);
								for (var n = 1; n < cell.rowSpan; n++) {
									merged = [j-1, 0];
									for (var m = 0; m < cell.colSpan; m++) {
										merged.push(cell);
									}
									data[i+n].splice.apply(data[i+n], merged);
								}
								k += cell.colSpan-1;
							} else if (cell.colSpan > 1) {
								// horizontal merge. ex: B12:E12. Add missing cells (with same attribute but value) to current row
								merges.push([numAlpha(j) + (i + 1), numAlpha(j+cell.colSpan-1) + (i + 1)]);
								merged = [j, 0]
								for (var m = 0; m < cell.colSpan-1; m++) {
									merged.push(cell);
								}
								data[i].splice.apply(data[i], merged);
								k += cell.colSpan-1;
							} else if (cell.rowSpan > 1) {
								// vertical merge. ex: B12:B15. Add missing cells (with same attribute but value) to next columns
								for (var m = 1; m < cell.rowSpan; m++) {
									if (data[i+m]) {
										data[i+m].splice(j, 0, cell)
									} else {
										// readh the end of data
										cell.rowSpan = m;
										break;
									}
								}
								merges.push([numAlpha(j) + (i + 1), numAlpha(j) + (i + cell.rowSpan)]);
							}
							if (cell.rowSpan > 1 || cell.colSpan > 1) {
								// deletes value, rowSpan and colSpan from cell to avoid refering it from copied cells
								delete cell.value;
								delete cell.rowSpan;
								delete cell.colSpan;
							}
							s += '<c r="' 
								+ numAlpha(j) + (i + 1) + '"' 
								+ (style ? ' s="' + style + '"' : '') 
								+ (t ? ' t="' + t + '"' : '');
							if (val != null) {
								s += '>' 
								+ (cell.formula ? '<f>' + cell.formula + '</f>' : '') 
								+ '<v>' 
									+ val 
								+ '</v></c>';
							} else {
								s += '/>';
							}
			    		}
			    		s += '</row>';
			    	}
			    	var cols = []
					if (opts.forcedColWidth) {
						for (i = 0; i < opts.forcedColWidth.length; i++) {
							if (opts.forcedColWidth[i] >= 0) {
								cols.push('<col customWidth="1" min="', i+1, '" max="', i+1, '" width="', options.forcedColWidth[i], '"/>');
							}
						}
					} else {
						for (i = 0; i < columns.length; i++) {
							if (columns[i].autoWidth) {
								cols.push('<col min="', i+1, '" max="', i+1, '" width="', columns[i].max, '" bestFit="1"/>');
							}
						}
					}
					// only add cols definition if not empty
					if (cols.length > 0) {
						cols = ['<cols>'].concat(cols, ['</cols>']).join('');
					}

					s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
						+ '<dimension ref="A1:' + numAlpha(worksheet[0].length - 1) + worksheet.length + '"/><sheetViews>'
						+ '<sheetView ' + (w === opts.activeWorksheet ? 'tabSelected="1" ' : '') + ' workbookViewId="0"'
						//+ (worksheet.view ? ' view="' + worksheet.view + '"' : '')
						+ '/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
						+ cols
						+ '<sheetData>'
						+ s 
						+ '</sheetData>';
					if (merges.length > 0) {
						s += '<mergeCells count="' + merges.length + '">';
						for (i = 0; i < merges.length; i++) {
							s += '<mergeCell ref="' + merges[i].join(':') + '"/>';
						}
						s += '</mergeCells>';
					}
					var paper_size = 9;
					var orientation = 'portrait';
					var margins = {};
					var page_scale = 100;
					if (opts.page) {
						if (opts.page.margins) margins = opts.page.margins;
						if (opts.page.paper_size) paper_size = opts.page.paper_size;
						if (opts.page.orientation) orientation = opts.page.orientation;
						if (opts.page.page_scale) page_scale = opts.page.page_scale;
					}
					s += '<pageMargins left="' + (margins.left || 0.7) +  '" right="' + (margins.right || 0.7) +  '" top="' + (margins.top || 0.75) +  '" bottom="' + (margins.bottom || 0.75) +  '" header="' + (margins.header || 0.3) +  '" footer="' + (margins.footer || 0.3) +  '"/>';
					s += '<pageSetup paperSize="' + paper_size + '" orientation="' + orientation + '" scale="' + page_scale + '"/>'
					/*if (worksheet.table) { 
						s += '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'; 
					}*/
					this.zip.updateEntry('xl/worksheets/sheet' + id + '.xml', s + '</worksheet>');

					/*if (worksheet.table) {
						i = -1; l = data[0].length; t = numAlpha(data[0].length - 1) + data.length;
						s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="' + id
							+ '" name="Table' + id + '" displayName="Table' + id + '" ref="A1:' + t + '" totalsRowShown="0"><autoFilter ref="A1:' + t + '"/><tableColumns count="' + data[0].length + '">';
						while (++i < l) { 
							s += '<tableColumn id="' + (i + 1) + '" name="' + (data[0][i].hasOwnProperty('value') ? data[0][i].value : data[0][i]) + '"/>'; 
						}
						s += '</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';

						xl.folder('tables').file('table' + id + '.xml', s);
						xlWorksheets.folder('_rels').file('sheet' + id + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table' + id + '.xml"/></Relationships>');
						contentTypes[1].unshift('<Override PartName="/xl/tables/table' + id + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>');
					}*/

					contentTypes[0].unshift('<Override PartName="/xl/worksheets/sheet' + id + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
					props.unshift(escapeXML(data.sheets[w]) || 'Sheet' + id);
					xlRels.unshift('<Relationship Id="rId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + id + '.xml"/>');
					worksheets.unshift('<sheet name="' + (escapeXML(data.sheets[w]) || 'Sheet' + id) + '" sheetId="' + id + '" r:id="rId' + id + '"/>');

				}

				var i = styles.length; 
				var t = [];
				while (--i) { 
					// Don't process index 0, already added
					var style = JSON.parse(styles[i]);

					// cell formating, refer to it if necessary
					if (style.formatCode !== 'General') {
						index = numFmts.indexOf(style.formatCode);
						if (index < 0) { 
							index = 164 + t.length; 
							t.push('<numFmt formatCode="' + style.formatCode + '" numFmtId="' + index + '"/>'); 
						}
						style.formatCode = index
					} else {
						style.formatCode = 0
					}

					// border declaration: add a new declaration and refer to it in style
					var borderIndex = 0
					if (style.borders) {
						border = ['<border>']
						// order is significative
						for (var edge in {left:0, right:0, top:0, bottom:0, diagonal:0}) {
							if (style.borders[edge]) {
								var color = style.borders[edge];
								// add transparency if missing
								if (color.length === 6) {
									color = 'FF'+color;
								}
								border.push('<', edge, ' style="thin">', '<color rgb="', style.borders[edge], '"/></', edge, '>');
							} else {
								border.push('<', edge, '/>');
							}
						}
						border.push('</border>');
						border = border.join('');
						// try to reuse existing border
						borderIndex = borders.indexOf(border);
						if (borderIndex < 0) {
							borderIndex = borders.push(border) - 1;
						}
					}

					// font declaration: add a new declaration and refer to it in style
					var fontIndex = 0
					if (style.bold || style.italic || style.fontSize || style.fontName || style.fontColor) {
						font = ['<font>']
						if (style.bold) {
							font.push('<b/>');
						}
						if (style.italic) {
							font.push('<i/>');
						}
						font.push('<sz val="', style.fontSize || defaultFontSize, '"/>');
						font.push('<color rgb="', style.fontColor || defaultFontColor, '"/>');
						font.push('<name val="', style.fontName || defaultFontName, '"/>');
						font.push('<family val="2"/>', '</font>');
						font = font.join('');
						// try to reuse existing font
						fontIndex = fonts.indexOf(font);
						if (fontIndex < 0) {
							fontIndex = fonts.push(font) - 1;
						}
					}

					// declares style, and refer to optionnal formatCode, font and borders
					styles[i] = ['<xf xfId="0" fillId="0" borderId="', 
						borderIndex, 
						'" fontId="',
						fontIndex,
						'" numFmtId="',
						style.formatCode,
						'" ',
						(style.hAlign || style.vAlign || style.textRotation? 'applyAlignment="1" ' : ' '),
						(style.formatCode > 0 ? 'applyNumberFormat="1" ' : ' '),
						(borderIndex > 0 ? 'applyBorder="1" ' : ' '),
						(fontIndex > 0 ? 'applyFont="1" ' : ' '),
						'>'
					];
					if (style.hAlign || style.vAlign || style.textRotation) {
						styles[i].push('<alignment');
						if (style.hAlign) {
							styles[i].push(' horizontal="', style.hAlign, '"');
						}
						if (style.vAlign) {
							styles[i].push(' vertical="', style.vAlign, '"');
						}
						if (style.textRotation) {
							styles[i].push(' textRotation="', style.textRotation, '"');
						}
						styles[i].push('/>');
					}
					styles[i].push('</xf>');
					styles[i] = styles[i].join('');
				}
				t = t.length ? '<numFmts count="' + t.length + '">' + t.join('') + '</numFmts>' : '';

				this.zip.updateEntry('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
					+ t + '<fonts count="'+ fonts.length + '" x14ac:knownFonts="1"><font><sz val="' + defaultFontSize + '"/><color rgb="' + defaultFontColor + '"/><name val="' + defaultFontName + '"/><family val="2"/>'
					+ '<scheme val="minor"/></font>' + fonts.join('') + '</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
					+ '<borders count="' + borders.length + '"><border><left/><right/><top/><bottom/><diagonal/></border>'
					+ borders.join('') + '</borders><cellStyleXfs count="1">'
					+ '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + styles.length + '"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
					+ styles.join('') + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>'
					+ '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
					+ '<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
					+ '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>');

				// [Content_Types].xml
				zip.updateEntry('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
					+ contentTypes[0].join('') + '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
					+ contentTypes[1].join('') + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');

				// docProps/app.xml
				zip.updateEntry('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>XLSX.js</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>'
					+ data.sheets.length + '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' + props.length + '" baseType="lpstr"><vt:lpstr>' + props.join('</vt:lpstr><vt:lpstr>')
					+ '</vt:lpstr></vt:vector></TitlesOfParts><Manager></Manager><Company>Microsoft Corporation</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>');

				// xl/_rels/workbook.xml.rels
				zip.updateEntry('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
					+ xlRels.join('') + '<Relationship Id="rId' + (xlRels.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
					+ '<Relationship Id="rId' + (xlRels.length + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
					+ '<Relationship Id="rId' + (xlRels.length + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>');

				// xl/sharedStrings.xml
				zip.updateEntry('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'
					+ sharedStrings[1] + '" uniqueCount="' + sharedStrings[0].length + '"><si><t>' + sharedStrings[0].join('</t></si><si><t>') + '</t></si></sst>');

				// xl/workbook.xml
				zip.updateEntry('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
					+ '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView '
					+ (opts.activeWorksheet ? 'activeTab="' + opts.activeWorksheet + '" ' : '') + 'xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/></bookViews><sheets>'
					+ worksheets.join('') + '</sheets><calcPr calcId="145621"/></workbook>');
				
				
				

				return callback(null, zip.buffer);
			}
		}).call(Xlsx.prototype);
		(function() {
			this.read = function(data, callback) {
				var xlsx = new Xlsx();
				return xlsx.read(data, callback);
			}
			this.write = function(data, options, callback) {
				var xlsx = new Xlsx();
				return xlsx.write(data, options, callback);
			}
		}).call(module.exports);
	