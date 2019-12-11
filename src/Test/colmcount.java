package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class colmcount {
	public static XSSFSheet ExcelWSheet;
	public static XSSFWorkbook ExcelWBook;
	public static XSSFCell Cell;
	public static XSSFRow Row;
	public MissingCellPolicy xRow;
		public static void main(String[] args) throws Exception {
		
	
			FileInputStream file = new FileInputStream(new File("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\ViaductDroppinglist2.xlsx"));
			
			//Get the workbook instance for XLS file 
			ExcelWBook  = new XSSFWorkbook(file);

			//Get first sheet from the workbook
		
			ExcelWSheet = ExcelWBook.getSheetAt(0);
			Row = ExcelWSheet.getRow(0);
			int colNum = Row.getLastCellNum();
			System.out.println(colNum);
		    int rowNum = ExcelWSheet.getLastRowNum();
		    System.out.println(rowNum);
					
			//Iterate through each rows from first sheet
			//ExcelWSheet=ExcelWBook.getSheet(SheetName);
		}
}
			