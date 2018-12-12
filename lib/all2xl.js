const uri = 'data:application/vnd.ms-excel;base64,',
    template = '<html dir="rtl" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
        '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">' +
        '<head>' +
        '<!--[if gte mso 9]>' +
        '<xml><x:ExcelWorkbook>' +
        '<x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>{worksheet}</x:Name><x:WorksheetOptions x:rightToLeft="1"><x:DisplayGridlines/></x:WorksheetOptions>' +
        '</x:ExcelWorksheet ></x:ExcelWorksheets>' +
        '</x:ExcelWorkbook></xml><![endif]-->' +
        '</head><body><table>{table}</table></body></html>';
        base64=function(s){return window.btoa(unescape(encodeURIComponent(s)));};

        format=function(s,c){
            s.replace(/ *\[[^)]*\] */g, "");
            return s.replace(/{(\w+)}/g,function(m,p){
                return c[p].replace(/ *\[[^)]*\] */g, "");
            })
        };

        function StringToExcel(tableId,worksheetName){

                var table=$(tableId),
                    ctx={worksheet:worksheetName,table:tableId},
                    href=uri+base64(format(template,ctx));
                console.log(table.html());
                console.log(table);
            return href;
        }
        function HtmlToExcel(tableId,worksheetName){
                var table=$(tableId);
                    ctx={worksheet:worksheetName,table:table.html()};
                    href=uri+base64(format(template,ctx));
                console.log(table.html());
                console.log(table);
            return href;
        }
