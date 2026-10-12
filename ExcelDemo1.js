const ExcelJs=require('exceljs');

async function excelTest()
{
    //creatig global object
     let output = {row:-1,column:-1};
    const workbook=new ExcelJs.Workbook();
    await workbook.xlsx.readFile("C:/Users/lalup/Downloads/ExceldownloadTest.xlsx");
    const worksheet=workbook.getWorksheet('Sheet1');
    worksheet.eachRow((row,rowNumber)=>{
     row.eachCell((cell,colNumber)=>{
        if(cell.value==="Banana")
        {
        output.row=rowNumber;
        output.column=colNumber;
        }

     });

    });
    //if (output.row !== -1 && output.column !== -1){

    const cell=worksheet.getCell(output.row,output.column); //upate thee apple value with iphone
    cell.value="Republic";
    await workbook.xlsx.writeFile("C:/Users/lalup/Downloads/ExceldownloadTest.xlsx");

console.log("Banana updated to Republic successfully");
}
   
  excelTest();