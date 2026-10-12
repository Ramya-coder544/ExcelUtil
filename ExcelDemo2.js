const ExcelJs=require('exceljs');

async function writeExcelTest(searchText,replacetext,change,filePath) {//Passing the arguments for the function call
const workbook= new ExcelJs.Workbook();
await workbook.xlsx.readFile(filePath);
const worksheet=workbook.getWorksheet('Sheet1');
const output=await readExcel(worksheet,searchText);//function call, we need to sent worksheet arug here from readExcel()
const cell=worksheet.getCell(output.row,output.column+change.colchange);//we have added 2 columns where
//we can find the price
cell.value=replacetext;
await workbook.xlsx.writeFile(filePath);
}
//Writing a separate function call

async function readExcel(worksheet,searchText)
{
    let output={row:-1,column:-1};
    worksheet.eachRow((row,rowNumber)=>
    {
       row.eachCell((cell,ColNumber)=>
        {
        if(cell.value===searchText)
            {
             output.row=rowNumber;
             output.column=ColNumber;
        }


       })
    })
    return output;

}
writeExcelTest("Mango","350",{rowchange:0,colchange:2},"C:/Users/lalup/Downloads/NewExcel.xlsx")