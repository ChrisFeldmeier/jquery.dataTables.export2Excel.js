/*
* File:        jquery.dataTables.export2Excel.js
* Version:     1.0.0.
* Author:      Christoph Feldmeier
* 
* Copyright 2015 Christoph Feldmeier, all rights reserved.
*
* Following js Files must be includes in your project
* // https://github.com/eligrey/FileSaver.js/    
* // https://github.com/eligrey/Blob.js
* // https://github.com/stephen-hardy/xlsx.js or (http://oss.sheetjs.com/js-xlsx/xlsx.core.min.js)
*
* The MIT License (MIT)
* 
* Copyright (c) 2015 Christoph Feldmeier
* 
* Permission is hereby granted, free of charge, to any person obtaining a copy of
* this software and associated documentation files (the "Software"), to deal in
* the Software without restriction, including without limitation the rights to
* use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
* the Software, and to permit persons to whom the Software is furnished to do so,
* subject to the following conditions:
* 
* The above copyright notice and this permission notice shall be included in all
* copies or substantial portions of the Software.
* 
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
* FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
* COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
* IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
* CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
* 
*/
(function ($) {

    $.fn.export2Excel = function (DataTable, options) {

        function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }        

        function datenum(v, date1904) {
            if (date1904) v += 1462;
            var epoch = Date.parse(v);
            return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
        };

        function sheet_from_array_of_arrays(data, opts) {
            var ws = {};
            var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
            for (var R = 0; R != data.length; ++R) {
                for (var C = 0; C != data[R].length; ++C) {
                    if (range.s.r > R) range.s.r = R;
                    if (range.s.c > C) range.s.c = C;
                    if (range.e.r < R) range.e.r = R;
                    if (range.e.c < C) range.e.c = C;
                    var cell = { v: data[R][C] };
                    if (cell.v == null) continue;
                    var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

                    if (typeof cell.v === 'number') cell.t = 'n';
                    else if (typeof cell.v === 'boolean') cell.t = 'b';
                    else if (cell.v instanceof Date) {
                        cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                        cell.v = datenum(cell.v);
                    }
                    else cell.t = 's';

                    ws[cell_ref] = cell;
                }
            }
            if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
            return ws;
        };

        function Workbook() {
            if (!(this instanceof Workbook)) return new Workbook();
            this.SheetNames = [];
            this.Sheets = {};
        };

        var oTable = this;
        var oTableData = null;
        var exportData = [];

        var defaults = {
            sFileName: "Export.xlsx",
            sWorkBookName: "Export",
            sBookType: "xlsx",
            sWriteType: "binary",
            sStreamType: "application/octet-stream",
            bWithColumnsTitle: true,
            aNotDisplayColumns: [
                //"COLUMN_NAME1",
                //"COLUMN_NAME2",
            ],
            aDisplayColumns: [
                //"COLUMN_NAME1",
                //"COLUMN_NAME2",
            ]            
        };

        var properties = $.extend(defaults, options);
                      
        var generateWorkbook = function(worksheet) {
            var wb = new Workbook();            
            wb.SheetNames.push(defaults.sWorkBookName);
            wb.Sheets[defaults.sWorkBookName] = worksheet;            
            var WorkBookOut = XLSX.write(wb, { bookType: defaults.sBookType, bookSST: true, type: defaults.sWriteType });                    
            saveAs(new Blob([s2ab(WorkBookOut)], { type: defaults.sStreamType }), defaults.sFileName);
        };

        var init = function() { 
           
            oTableData = DataTable.data().toArray();            
           
            // loop through data
            oTableData.forEach(function(row) {
                var rItem = [];                                
                for(var item in row) {    
                    // not display columns          
                    if (defaults.aNotDisplayColumns.length > 0) {
                        if (defaults.aNotDisplayColumns.indexOf(item) == -1) {
                            rItem.push(row[item]);
                        }
                    }
                    // display columns
                    if (defaults.aDisplayColumns.length > 0) {
                        if (defaults.aDisplayColumns.indexOf(item) > -1) {
                            rItem.push(row[item]);
                        }
                    }

                    // default
                    if (defaults.aNotDisplayColumns.length == 0 && defaults.aDisplayColumns.length == 0) {
                        rItem.push(row[item]);
                    }
                };
                exportData.push(rItem);
            }, this);
            
            // set header title            
            var ColumnNames = [];                    
            for (var ColumnName in oTableData[0]) {
                // not display column header      
                if (defaults.aNotDisplayColumns.length > 0) {
                    if (defaults.aNotDisplayColumns.indexOf(ColumnName) == -1) {
                        ColumnNames.push(ColumnName);
                    }
                }
                // display columns header    
                if (defaults.aDisplayColumns.length > 0) {
                    if (defaults.aDisplayColumns.indexOf(ColumnName) > -1) {
                        ColumnNames.push(ColumnName);
                    }
                }

                // default 
                if (defaults.aNotDisplayColumns.length == 0 && defaults.aDisplayColumns.length == 0) {
                    ColumnNames.push(ColumnName);
                }
            };
                
            // export with column titles
            if(defaults.bWithColumnsTitle) {
                exportData.unshift(ColumnNames);                           
            }

            /*var ColumnTitles = [];
            // export columns header text
            if (defaults.bShowColumnsHeaderTitle) {
                /*$(DataTable.columns().header()).each(function (index) {                    
                    ColumnTitles.push($(this).text());
                });*/

                /*for (var Column in DataTable.settings()[0].aoColumns) {
                    var Col = DataTable.settings()[0].aoColumns[Column];
                    var data = Col.data;
                    /*if(Col.mData.length > 0) {
                        data = Col.mData;
                    }*/
                    //console.log(Col);
                  //  ColumnTitles.push(Col.title);
               // }

            //}
            //exportData.unshift(ColumnTitles);
             
            generateWorkbook(sheet_from_array_of_arrays(exportData));
        };
        return init();
    };

})(jQuery);

