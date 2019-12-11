package Lib;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Raj_Utility {
public static final String Path="";
private static XSSFSheet ExcelWSheet;
private static XSSFWorkbook ExcelWBook;
private static XSSFCell Cell;
private static XSSFRow Row;
private static MissingCellPolicy xRow;

//This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method

	public static void SetExcelFile(String Path,String SheetName) throws Exception{
		
		try{
			// Open the Excel file
			
		FileInputStream File= new FileInputStream(Path);
		
		     //Access the required test data sheet
	    ExcelWBook = new XSSFWorkbook(File);
	    ExcelWSheet = ExcelWBook.getSheet(SheetName);
		}
		catch(Exception e){
			throw (e);
			
		}
	}

    public  static int rowcount(String SheetName){
    	ExcelWSheet = ExcelWBook.getSheet(SheetName);
    	
    int TotalNofRows=ExcelWSheet.getLastRowNum();
		return TotalNofRows;
    	
    }
    public  static int columncount(String SheetName){
    	 int   row1,col1,maxrow = 0,maxcol = 0;
    	ExcelWSheet = ExcelWBook.getSheetAt(0);
    	XSSFRow r = ExcelWSheet.getRow(0);
    	int maxCell=  r.getLastCellNum();
    	
		
		return columncount(null);
		
    }

}


    