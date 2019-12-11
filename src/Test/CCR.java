package Test;




import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;



public class CCR {
	
	//declaring the public static variable
	
	 //static int count=0;
//	 
	 //main method
public static void main(String[] args) throws Exception {
{    
	
	//browser properties
	System.setProperty("webdriver.chrome.driver","C://Users//ammanrr.CORP//Downloads//chromedriver.exe");
	
	
	WebDriver driver = new ChromeDriver(); 
	driver.manage().window().maximize();
	driver.manage().deleteAllCookies();
	driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	
driver.get("http://ccrmgtconsole.dev.invesco.net/default.aspx");
////*[@id="ctl00_ContentPalceHolder2_dgReportReady"]/tbody/tr//input[@type='submit']
WebElement e1=driver.findElement(By.xpath("//*[@id='ctl00_ContentPalceHolder2_dgReportReady']/tbody/tr[2]/td[2]//input[@type='submit']"));
String text=e1.getAttribute("value");
System.out.println("value is" +text);
}}}



