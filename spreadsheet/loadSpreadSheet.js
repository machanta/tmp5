function demoHandler(){
    let upload = document.getElementById("Luckyexcel-demo-file");
    if(upload){
        window.onload = () => {
            upload.addEventListener("change", function(evt){
                var files = evt.target.files;
                if(files==null || files.length==0){
                    alert("No files wait for import");
                    return;
                }

                let name = files[0].name;
                let suffixArr = name.split("."), suffix = suffixArr[suffixArr.length-1];
                if(suffix!="xlsx"){
                    alert("Currently only supports the import of xlsx files");
                    return;
                }
                LuckyExcel.transformExcelToLucky(files[0], function(exportJson, luckysheetfile){
                    
                    if(exportJson.sheets==null || exportJson.sheets.length==0){
                        alert("Failed to read the content of the excel file, currently does not support xls files!");
                        return;
                    }
                    window.luckysheet.destroy();
                    
                    window.luckysheet.create({
                        container: 'luckysheet', //luckysheet is the container id
                        showinfobar:false,
                        data:exportJson.sheets,
                        title:exportJson.info.name,
                        userInfo:exportJson.info.name.creator
                    });
                });
            });

        }
    }
}