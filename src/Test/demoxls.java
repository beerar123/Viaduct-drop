package Test;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class demoxls {
	
public static void main(String args[]) throws IOException{
{
		FileInputStream fis=new FileInputStream("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\Credentials.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook(fis);
	XSSFSheet sheet= workbook.getSheet("Credentials");
	XSSFRow row = sheet.getRow(0);
	
	int col_num= -1;
	
	for (int i =0; i<row.getLastCellNum(); i++){
		if(row.getCell(i).getStringCellValue().trim().contains("UserName"))
			col_num = i;
		
		System.out.println("the value of the row " + i);

	}
	row = sheet.getRow(2);
	XSSFCell cell = row.getCell(col_num);
	
	String value= cell.getStringCellValue();
	System.out.println("Value of the Excell cell is - "+ value);
	
	
	
	
	} 
}
}
