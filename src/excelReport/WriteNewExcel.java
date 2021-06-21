package excelReport;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class WriteNewExcel {
    public void writeExcel(String filePath,String fileName,String sheetName,String[] dataToWrite) throws IOException{
        File file = new File(filePath+"\\"+fileName);
        FileInputStream inputStream = new FileInputStream(file);
        Workbook ExcelFile = null;
        String fileExtensionName = fileName.substring(fileName.indexOf("."));
        if(fileExtensionName.equals(".xlsx")){
        	ExcelFile = new XSSFWorkbook(inputStream);
        }
        else if(fileExtensionName.equals(".xls")){
        	ExcelFile = new HSSFWorkbook(inputStream);
        }    
    Sheet sheet = ExcelFile.getSheet(sheetName);
    int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
    Row row = sheet.getRow(0);
    Row newRow = sheet.createRow(rowCount+1);
    for(int j = 0; j < row.getLastCellNum(); j++){
        Cell cell = newRow.createCell(j);
        cell.setCellValue(dataToWrite[j]);
    }
    inputStream.close();
    FileOutputStream outputStream = new FileOutputStream(file);
    ExcelFile.write(outputStream);
    outputStream.close();
    }

    public static void main(String...strings) throws IOException{
        String[] valueToWrite = {"Madan","Chennai"};
        WriteNewExcel objExcelFile = new WriteNewExcel();
        objExcelFile.writeExcel(System.getProperty("user.dir")+"\\src\\excelReport","Excel.xlsx","ReadExcel",valueToWrite);
    }
}