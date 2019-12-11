package Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class demoxls1 {
public static void main(String args[]) throws IOException
{
	FileInputStream fis= new FileInputStream("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\Credentials.xlsx");
	XSSFWorkbook workbook= new XSSFWorkbook(fis);
	XSSFSheet sheet = workbook.getSheet("Credentials");
    XSSFRow row = sheet.getRow(0);
    
    int colNum = -1;
    
    for (int i =0; i <row.getLastCellNum(); i ++)
    {
    	if(row.getCell(i).getStringCellValue().trim().contains("Pass"))
    		colNum=i;
    
    }
    row = sheet.getRow(1);
	XSSFCell cell = row.getCell(colNum);
	String password = String.valueOf(cell.getStringCellValue());
	System.out.println("Value from the Excel sheet :"+ password);
}
}
