package Package1.MavenEx;

import java.io.FileInputStream;
import java.io.IOException;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class TestNGExcel 
{
	WebDriver d1;
	/*public boolean isAlertPresent()
	{
		try
		{
			d1.switchTo().alert();
			return true;
		}
		catch(Exception e)
		{
			return false;
		}
	}*/
	@BeforeClass
	public void driversetup()
	{
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Mohit Agarwal\\Documents\\selenium\\chromedriver.exe");
		//System.setProperty("webdriver.firefox.marionette", "C:\\Users\\Mohit Agarwal\\Documents\\selenium\\geckodriver.exe");
		//WebDriver d1 = new FirefoxDriver();
		d1 = new ChromeDriver();
		d1.manage().window().maximize();   // To Maximize Browser
	}
	
	@Test(priority=1)
	public void webpage()
	{
		d1.get("http://www.demo.guru99.com/V4/");
	}
	
	Sheet s1;
	@Test(priority=2)
	public void openexcel() throws BiffException, IOException
	{
		FileInputStream f1=new FileInputStream("C:\\Users\\Mohit Agarwal\\Documents\\selenium\\guru.xls");
		Workbook w1= Workbook.getWorkbook(f1);
		s1=w1.getSheet("Sheet1");
	}
	
	@Test(priority=3)
	public void reportcols()
	{
		/*Reporter.log("Username");
		Reporter.log("\t");
		Reporter.log("Password");
		Reporter.log("\t");
		Reporter.log("Result");
		Reporter.log("\n");
		*/
	}
	
	@Test(priority=4)
	public void readexcel() throws InterruptedException
	{
		//get row count
		int rows = s1.getRows();
		int cols = s1.getColumns();
		
		/*System.out.print("Username");
		System.out.print("\t");
		System.out.print("Password");
		System.out.print("\t");
		System.out.print("Result");
		System.out.print("\n");
		*/
		
		
		for(int x=1;x<rows;x++)
		{
			for(int y=0;y<cols;y++)
			{
				String z=s1.getCell(y,x).getContents();    
				System.out.print(z);
				Reporter.log(z);
				System.out.print("\t");
				Reporter.log("\t");
				
				if(y==0)
				d1.findElement(By.name("uid")).sendKeys(z);
			
				if(y==1)
					d1.findElement(By.name("password")).sendKeys(z);
			}
			d1.findElement(By.name("btnLogin")).click();
			Thread.sleep(5000);
			//String t = d1.getTitle();
			//System.out.println(t);
			//Alert c = d1.switchTo().alert();
			try
			{
				Alert a = d1.switchTo().alert();
				System.out.print(a.getText());
				Reporter.log(a.getText());
				a.accept();
				System.out.print("\n");
			}
			catch(NoAlertPresentException na)
			{
				Thread.sleep(5000);
				String t = d1.getTitle();
				System.out.print(t);
				Reporter.log(t);
				System.out.print("\n");
				Reporter.log("\n");
				
				
				d1.findElement(By.linkText("Log out")).click();
				Alert b = d1.switchTo().alert();
				System.out.print(b.getText());
				Reporter.log(b.getText());
				b.accept();
				
				String lo = d1.getTitle();
				System.out.println("Logout Complete");
				Reporter.log("Logout Complete");
				System.out.println(lo);
				Reporter.log(lo);
			}
			
			catch(Exception e)
			{
				d1.findElement(By.xpath("//marquee[@class='heading3']")).getText();
			}
		}	
		Thread.sleep(5000);	
	}
		
	@AfterClass
	public void Last()
	{
		d1.close();  // Or d1.quit()  To Close Browser
	}
}
