package FactSheetDownload;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
//import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
public class Utility4 {
public static final String Path="C:\\Users\\ammanrr.CORP\\eclipse-workspace\\TrackerList.xlsx";
private static XSSFSheet ExcelWSheet;
private static XSSFWorkbook ExcelWBook;
private static XSSFCell Cell;
private static XSSFRow Row;
private static MissingCellPolicy xRow;
//This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method

public static void setExcelFile(String Path,String SheetName) throws Exception {

try {

// Open the Excel file

FileInputStream ExcelFile = new FileInputStream(Path);

//Access the required test data sheet

ExcelWBook = new XSSFWorkbook(ExcelFile);

ExcelWSheet = ExcelWBook.getSheet(SheetName);

} catch (Exception e){

throw (e);

}

}

public static int rowcount(String SheetName) {
	ExcelWSheet=ExcelWBook.getSheet(SheetName);
int totalNoOfRows=ExcelWSheet.getLastRowNum();

return totalNoOfRows;
}

//This method is to read the test data from the Excel cell, in this we are passing parameters as Row num and Col num

public static String getCellData1(String sSheetName,int RowNum, int ColNum) throws Exception{

try{

//String str= ExcelWSheet.getSheetName();


String CellData = Cell.getStringCellValue();
//int totalNoOfRows=ExcelWSheet.getLastRowNum();
return CellData;

}catch (Exception e){

return"";

}

}


public static String getCellData(String file, String sheetname, int RowNum, int ColNum) throws Exception {

    String CellData = null;
    FileInputStream excelfile = new FileInputStream(file);
    try {
       // String filename = file.substring(file.indexOf("."));
    //    if (filename.equalsIgnoreCase(".xls")) {
             // File file = new File("InputData.xls");
/*
             @SuppressWarnings("resource")
             HSSFWorkbook wb = new HSSFWorkbook(excelfile);
             Sheet sh = wb.getSheet(sheetname);
             DataFormat formatter = new DataFormatter();
             
             Cell cell = sh.getRow(RowNum).getCell(ColNum);
             Thread.sleep(4000);
             CellData = formatter.formatCellValue(cell);
             */
        
        //   int CellType = Cell.getCellType();
             
        //.out.println(CellType);
             //CellData = Cell.getStringCellValue();
        //   System.out.println(CellData);
       // } else if (filename.equalsIgnoreCase(".xlsx")) {
             // File file = new File("InputData.xls");

             @SuppressWarnings("resource")
             XSSFWorkbook wb = new XSSFWorkbook(excelfile);
             Sheet sh = wb.getSheet(sheetname);
             DataFormatter formatter = new DataFormatter();
             Cell cell = sh.getRow(RowNum).getCell(ColNum);

             CellData = formatter.formatCellValue(cell);
             
    
    } catch (Exception e) {
        System.out.printf("Unable to return the value of a cell", e.getMessage());
        throw (e);
    }

    return CellData;

}

//This method is to write in the Excel cell, Row num and Col num are the parameters

@SuppressWarnings("static-access")
public static void setCellData(String Result,  int RowNum, int ColNum) throws Exception	{

try{
	
	Row=ExcelWSheet.getRow(RowNum);
if(Row==null) 
Row  = ExcelWSheet.createRow(RowNum);

Cell=Row.getCell(ColNum);



Cell = Row.getCell(ColNum, xRow.RETURN_BLANK_AS_NULL);

if (Cell == null) {

Cell = Row.createCell(ColNum);

Cell.setCellValue(Result);

} else {

Cell.setCellValue(Result);

}

// Constant variables Test Data path and Test Data file name

FileOutputStream fileOut = new FileOutputStream(Path);

ExcelWBook.write(fileOut);

fileOut.flush();

fileOut.close();

}catch(Exception e){

throw (e);

}

}

}