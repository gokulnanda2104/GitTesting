package excelReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadNewExcel {
    public void readExcel(String filePath,String fileName,String sheetName) throws IOException{
    File file =new File(filePath+"\\"+fileName);
    FileInputStream inputStream = new FileInputStream(file);
    Workbook ExcelFile = null; 
    String fileExtensionName = fileName.substring(fileName.indexOf("."));
    if(fileExtensionName.equals(".xlsx")){
    	ExcelFile = new XSSFWorkbook(inputStream);
    }
    else if(fileExtensionName.equals(".xls")){
    	ExcelFile = new HSSFWorkbook(inputStream);
    } 
    Sheet NewExcel = ExcelFile.getSheet(sheetName);
    int rowCount = NewExcel.getLastRowNum()-NewExcel.getFirstRowNum();
    for (int i = 0; i < rowCount+1; i++) {
        Row row = NewExcel.getRow(i);
        for (int j = 0; j < row.getLastCellNum(); j++) {
            System.out.print(row.getCell(j).getStringCellValue()+"   ");
        }
        System.out.println();
    } 
    }  
    public static void main(String...strings) throws IOException{
    ReadNewExcel objExcelFile = new ReadNewExcel();
    objExcelFile.readExcel(System.getProperty("user.dir")+"\\src\\excelReport","Excel.xlsx","ReadExcel");
    }
}
