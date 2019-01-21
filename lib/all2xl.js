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
        function HtmlToExcel(tableId,worksheetName){
            var table=document.getElementById(tableId);
            ctx={worksheet:worksheetName,table:table.innerHTML};
            href=uri+base64(format(template,ctx));
            window.location.replace(href);
        }
