package FactSheetDownload;

import java.io.FileInputStream;
import java.util.Properties;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import FactSheetDownload.Utility4;

public class BulkFactSheetDownload {
	
	   static String sFileName ="C:\\Users\\ammanrr\\eclipse-workspace\\TickerList.xlsx" ;
	   static String sSheetName ="Sheet1";
		public static void main(String[] args) throws Exception {
		
		// launching the browser
		try {
			System.setProperty("webdriver.chrome.driver", "C://Users//ammanrr//Downloads//chromedriver.exe");
		} 
		catch (Exception e)
		{
			System.out.println("unable to find the chrome driver");
		}

		WebDriver driver = new ChromeDriver();
		// maximizing the browser
		driver.manage().window().maximize();

		// searching for user data file
		try {
			Utility4.setExcelFile("C:\\Users\\ammanrr\\eclipse-workspace\\TickerList.xlsx", "Sheet1");
			} 
		catch (Exception e) {
			System.out.println("Given Updated user input file is not availble in specified location");
			}
			
		//loading the properties file
		Properties obj = new Properties();					
	    FileInputStream objfile = new FileInputStream(System.getProperty("user.dir")+"\\dataFile.properties");
	    obj.load(objfile);
		
	    // Entering the user login data from excel file
		try {
			driver.get(Utility4.getCellData(sFileName,sSheetName,1, 1));

			driver.findElement(By.xpath(obj.getProperty("username"))).sendKeys(Utility4.getCellData(sFileName,sSheetName,1, 2));

			driver.findElement(By.xpath(obj.getProperty("password"))).sendKeys(Utility4.getCellData(sFileName,sSheetName,1, 3));

			driver.findElement(By.xpath(obj.getProperty("loginbutton"))).sendKeys(Keys.ENTER);
			} catch (Exception e) {
			System.out.println("User is unable to login into aplication");
			}
		
		//Clicking on the obligation and task tab
		Thread.sleep(3000);
		driver.findElement(By.xpath(obj.getProperty("obligation_task"))).click();
		
		//Clicking, Entering the ticker name from excel and downloading the factsheet from the obligation list
		Thread.sleep(5000);
		
		String sSheet1 = "Sheet1";
		int totalNoOfRows = Utility4.rowcount(sSheet1);
		System.out.println(totalNoOfRows);
		for (int row = 1; row <= totalNoOfRows; row++) 
		{
		
		WebElement filter1=driver.findElement(By.xpath(obj.getProperty("filter")));
		filter1.clear();
		
		filter1.sendKeys(Utility4.getCellData(sFileName, sSheetName, row, 0));
		
		filter1.sendKeys(Keys.ENTER);
		Thread.sleep(1000);
		driver.findElement(By.xpath(obj.getProperty("CheckBox"))).click();
		
		Thread.sleep(1000);
		driver.findElement(By.xpath(obj.getProperty("DownloadButton"))).click();
		
		try {
		WebElement notfound=driver.findElement(By.xpath(obj.getProperty("ReportNotFound")));
		
		if(notfound!=null) { 
			
			System.out.println("FactSheet not Downloaded Successfully - "+Utility4.getCellData(sFileName,sSheetName,row, 0));
			driver.findElement(By.xpath(obj.getProperty("ReportButtonClose"))).click();
		}else {
			
			System.out.println("FactSheet Downloaded Successfully - "+Utility4.getCellData(sFileName,sSheetName,row, 0));
		}
		}catch(Exception e) {
			System.out.println("FactSheet Downloaded Successfully - "+Utility4.getCellData(sFileName,sSheetName,row, 0));
		}
		Thread.sleep(1000);
		driver.findElement(By.xpath(obj.getProperty("CheckBox"))).click();
		
		}
		}
		}