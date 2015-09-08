# jquery.dataTables.export2Excel.js
DataTables.js Extension to implement nativ XLSX Export 

#Install
Following js Files must be includes in your project
* https://github.com/eligrey/FileSaver.js/    
* https://github.com/eligrey/Blob.js
* https://github.com/stephen-hardy/xlsx.js or (http://oss.sheetjs.com/js-xlsx/xlsx.core.min.js)

# Example usage: (all Parameters are optional)
```javascript
var expert2excel = new $.fn.export2Excel(AppCustomerFullSearch.oDataTable,
    {
        // Filename of the export File
        sFileName: "CRM-Search.xlsx",
        // Name of the worksheet
        sWorkBookName: "Sheet1",
        // Book Type
        sBookType: "xlsx",
        // Export Column Titles 
        bWithColumnsTitle: true,
        // Only Export some specific columns
        "aDisplayColumns": [
            "COLUMN1",
            "COLUMN2",
            "COLUMN3",
        ],
        // Columns which should not exported
        /*aNotDisplayColumns: [
            "COLUMN1",
            "COLUMN2",
            "COLUMN3"
        ],*/
        
    });
```
