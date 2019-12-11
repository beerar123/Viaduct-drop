package Test;

import java.util.concurrent.TimeUnit;


import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
//import org.openqa.selenium.JavascriptExecutor;
//import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
//import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import org.openqa.selenium.support.ui.ExpectedConditions;
//import org.openqa.selenium.interactions.Actions;
//import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import Lib.Utility4;

public class MyViaductDrop {

	//declaring the public static variable
	 static String sFileName ="C:\\Users\\beerar\\Downloads\\Selenium_Automation\\ViaductDroppinglistMINE.xlsx" ;
	 static String sSheetName ="Sheet1";
	 static String sRunMode;
	 //static int count=0;
	 
	 //main method
public static void main(String[] args) throws Exception {
{    
	
	//browser properties
	System.setProperty("webdriver.chrome.driver", "C:\\Users\\beerar\\Downloads\\Selenium_Automation\\chromedriver.exe");
	
	
	WebDriver driver = new ChromeDriver(); 
	driver.manage().window().maximize();
	driver.manage().deleteAllCookies();
	driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

	//loading the excel file
	try {
	Utility4.setExcelFile("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\ViaductDroppinglistMINE.xlsx", "Sheet1");}
	catch (Exception e){System.out.println("invalid excel file");
	}

	//navigating to Viaduct UAT instance from excel file
	driver.get(Utility4.getCellData(sFileName, sSheetName, 1, 8));
	Thread.sleep(3000);

	//login details from excel file
	try {
	driver.findElement(By.name("SWEUserName")).sendKeys(Utility4.getCellData(sFileName, sSheetName, 1, 6));
	driver.findElement(By.name("SWEPassword")).sendKeys(Utility4.getCellData(sFileName, sSheetName, 1, 7));
	driver.findElement(By.id("s_swepi_22")).click();}catch(Exception e) {System.out.println("login failed");}
	Thread.sleep(3000);

	//browser refresh
	driver.navigate().refresh();
	Thread.sleep(5000);

	//click on communication tab
	try {
	WebDriverWait wait2 = new WebDriverWait(driver, 10);
	wait2.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-tabindex='tabScreen10']"))).click();;
	}catch(Exception e) {System.out.println("unable to click communication tab");}


	//Getting total project or communications from sheet1 that need to drop
	String sSheet1 = "Sheet1";
	int totalNoOfRows = Utility4.rowcount(sSheet1);
	System.out.println(+totalNoOfRows);
	int row;
	for ( row = 1; row <= totalNoOfRows; row++) 
	{
		//setting and verifying the Run mode
		int count=0;
		sRunMode=Utility4.getCellData(sFileName, sSheetName, row, 2);
		//checking RunMode status that need to drop

		if(sRunMode.equals("Yes")) {
	
			//click on query
	try {
		driver.findElement(By.xpath("//*[@id='s_1_1_8_0_Ctrl']")).click();}
		catch(Exception e) {System.out.println("unabble to click the query button");}

	//Entering the communication name from excel file
	String S="*";
	String Cname=Utility4.getCellData(sFileName, sSheetName, row, 0);

	//driver.findElement(By.xpath("//table[@id='s_1_l']/tbody/tr[2]/td[2]/input")).sendKeys(Utility4.getCellData(sFileName, sSheetName, row, 0));
	try {
		driver.findElement(By.xpath("//table[@id='s_1_l']/tbody/tr[2]/td[2]/input")).sendKeys(S+Cname+S);}
	catch(Exception e) {System.out.println("unable to enter the communication name");}

	//click on Go button
	driver.findElement(By.xpath("//*[@id='s_1_1_5_0_Ctrl']")).click();

	//click on show active
	//driver.findElement(By.xpath("//*[@id='s_3_1_3_0_Ctrl']")).click();

	//Getting the sheet name
	try {
	System.out.println(Utility4.getCellData(sFileName,sSheetName,row, 9));
	String str = Utility4.getCellData(sFileName,sSheetName,row, 9);
	System.out.println("account sheetname"+str);
	int totalNoOfRows1 = Utility4.rowcount(str);

	for(int Arow=9;Arow>=totalNoOfRows1;Arow++)
	
	{
		try {
			//click on act query
	driver.findElement(By.xpath("//button[@id='s_2_1_11_0_Ctrl']")).click();}
	catch(Exception e) {System.out.println("quer button is not responding");}

	try{//Enter the account number or communication name from excel file
	
	String S1="*";
	String ActName=Utility4.getCellData(sFileName,str,Arow, 9);
	driver.findElement(By.xpath("//input[@id='1_Account_Name']")).sendKeys(Keys.TAB);
	//driver.findElement(By.xpath("//input[@id='1_Account_Number']")).sendKeys(Utility4.getCellData(sFileName,str,Arow, 0));}
	driver.findElement(By.xpath("//input[@id='1_Account_Number']")).sendKeys(S1+S1+ActName);}
	catch(Exception e) {System.out.println("act number is not able to enter in specified location");}

	try{//click on Go button
	driver.findElement(By.xpath("//*[@id='s_2_1_8_0_Ctrl']")).click();}
	catch(Exception e) {System.out.println("go button is mot responding");}

	try {
		//click on communication instance
	WebDriverWait wait1=new WebDriverWait(driver,5);
	wait1.until(ExpectedConditions.elementToBeClickable(By.xpath("//table[@id='s_2_l']/tbody/tr[2]/td[9]/a"))).click();
	//driver.findElement(By.xpath("//table[@id='s_2_l']/tbody/tr[2]/td[9]/a")).click();
	Thread.sleep(5000);}
	catch(Exception e) {System.out.println("unable to click on communication isntance");}



	try {
		//click on Adhoc submit
	WebDriverWait wait1=new WebDriverWait(driver,5);
	wait1.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@aria-label='Communication Maintenance:Adhoc Submit']"))).click();
	//driver.findElement(By.xpath("//button[@aria-label='Communication Maintenance:Adhoc Submit']")).click();
	Thread.sleep(5000);}
	catch(Exception e) {System.out.println("unable to click on adhoc submit button");}


	try {
		//Handling the alert window
	
	
	Alert promptAlert = driver.switchTo().alert();

	//Enter the given Effective date i.e Quarterly or Monthly dates
	promptAlert.sendKeys(Utility4.getCellData(sFileName,sSheetName,row, 3));
	WebDriverWait wait3=new WebDriverWait(driver,15);
	wait3.until(ExpectedConditions.alertIsPresent());
	promptAlert.accept();
	Thread.sleep(10000);}
	catch(Exception e) {System.out.println("alert is not present");}
	try {
	Alert promptAlert1 = driver.switchTo().alert();
	//Thread.sleep(8000);
	
	WebDriverWait wait3=new WebDriverWait(driver,15);
	wait3.until(ExpectedConditions.alertIsPresent());
	promptAlert1.accept();
	Thread.sleep(8000);
	//Utility4.setExcelFile("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\ViaductDroppinglistMINE.xlsx", str);
	//Utility4.setCellData("Dropped", Arow, 1);
	
	}
	
	catch(Exception e) {System.out.println("alert2 is not working");
	Utility4.setExcelFile("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\ViaductDroppinglistMINE.xlsx", str);
	Utility4.setCellData("Failed", Arow, 1);
	count++;
	
	}
	//Return back to communication tab

	try{
		WebDriverWait wait1=new WebDriverWait(driver,5);
		wait1.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='siebui-threadbar']/li[1]/span/a"))).click();}
	//driver.findElement(By.xpath("//*[@id='siebui-threadbar']/li[1]/span/a")).click();}
	catch(Exception e) {System.out.println("unable to naviagte to back page");}
	Thread.sleep(5000);


	}
	}
	catch(Exception e) {
		System.out.println("unable to get the sheet name");
		}

	//Knowing the account numbers for that communication instance 

	if(count>=1) {
		Utility4.setExcelFile("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\ViaductDroppinglistMINE.xlsx", "Sheet1");
		Utility4.setCellData("Fail", row, 5);
		
	}else {
		Utility4.setExcelFile("C:\\Users\\beerar\\Downloads\\Selenium_Automation\\ViaductDroppinglistMINE.xlsx", "Sheet1");
		Utility4.setCellData("Pass", row, 5);
	}
		}
		//these communications which or depended on communication instances names no account numbers 
	
		}

}}}
