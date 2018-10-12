package shopdetails;

import org.testng.annotations.Test;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Assert;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class AllShopDetails {
	
	WebDriver driver;
	
	//elements
	String get_size="//div[contains(@data-async-context,'query:')]/div/div/div/div/div/div/div/div/div[4]/div";
	String store_name="//div[@class='kp-header']/div/div/div/div[@role='heading']/div/span[1]";
	String address = "//a[text()='Address']/parent::span/following-sibling::span";
	String phone = "//a[text()='Phone']/parent::span/following-sibling::span/span/span";
	String category ="//span[@class='YhemCb'][1]";
	String next="//span[text()='Next']";
	String rating="//div[@class='kp-header']/div/div[2]/div/div/div/span[@class='rtng']";
	String website="//a[text()='Website']";
	String dropdown_arrow="//div[@jsaction='oh.handleHoursAction']/span/span/span/span[2]";
	String timings="//div[@data-attrid='kc:/location/location:hours']";
	String input="";
	
	@BeforeTest
	public void launchBrowser() throws IOException {
		FileReader reader = new FileReader("./src/test/resources/input.properties");
		Properties p=new Properties();
		p.load(reader);
		input = p.getProperty("input");
		 ChromeOptions chromeOptions = new ChromeOptions();
		chromeOptions.addArguments("--headless");
		System.setProperty("webdriver.chrome.driver", "/Users/swetha/Documents/softwares/automation drivers/chromedriver");
		driver = new ChromeDriver(chromeOptions);
	}
	
	@Test
	public void getDetails() throws InterruptedException, IOException, StaleElementReferenceException {
		
		
		driver.navigate().to("https://google.com");
		driver.findElement(By.xpath("//input[@title='Search']")).sendKeys(input);
		Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@title='Search']")).sendKeys(Keys.TAB);
		driver.findElement(By.xpath("//input[@name='btnK']")).click();
		Thread.sleep(2000);
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("window.scrollBy(0,250)", "");
		Thread.sleep(2000);
		//driver.findElement(By.xpath("//span[text()='More places']")).click();
		driver.findElement(By.xpath("//div[@class='Mpn0ac']/img")).click();
		Thread.sleep(2000);
		
		Workbook workbook = new XSSFWorkbook();
		CreationHelper createHelper = workbook.getCreationHelper();
		Sheet sheet = workbook.createSheet(input);
		
		String[] columns = {"category","Store Name", "Address", "phone","Website","Timings","ratings"};
		
		// Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLUE.getIndex());
        
     // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }
        
      
       boolean condn=true;
        
       int rowNum = 1;
        try {
        	while(condn) {	       	
        	for (int i=1;i<driver.findElements(By.xpath(get_size)).size();i++) {
			driver.findElement(By.xpath("//div[contains(@data-async-context,'query:')]/div/div/div/div/div/div/div/div/div[4]/div["+i+"]")).click();
			Thread.sleep(300);
			Row row = sheet.createRow(rowNum++);
			for(int m=0;m<=2;m++) {
				try{
			if(driver.findElements(By.xpath(category)).size()!=0)
				row.createCell(0).setCellValue(driver.findElement(By.xpath(category)).getText());
			Thread.sleep(100);			
			
			if(driver.findElements(By.xpath(store_name)).size()!=0)
				row.createCell(1).setCellValue(driver.findElement(By.xpath(store_name)).getText());
			Thread.sleep(300);
			
			if(driver.findElements(By.xpath(address)).size()!=0)
				row.createCell(2).setCellValue(driver.findElement(By.xpath(address)).getText());
			Thread.sleep(100);
			
			if(driver.findElements(By.xpath(phone)).size()!=0)
				row.createCell(3).setCellValue(driver.findElement(By.xpath(phone)).getText());
			Thread.sleep(100);
			
			if(driver.findElements(By.xpath("//a[text()='Website']")).size()!=0)
				row.createCell(4).setCellValue(driver.findElement(By.xpath("//a[text()='Website']")).getAttribute("href"));
			Thread.sleep(100);
			
			if(driver.findElements(By.xpath(dropdown_arrow)).size()!=0) {
				Thread.sleep(1000);
				if(driver.findElements(By.xpath(dropdown_arrow)).size()!=0) {
				driver.findElement(By.xpath(dropdown_arrow)).click();
				Thread.sleep(100);
				row.createCell(5).setCellValue(driver.findElement(By.xpath(timings)).getText());}
			}
			Thread.sleep(100);
			row.createCell(6).setCellValue(driver.findElement(By.xpath(rating)).getText());
			
			break;
			}catch(Exception e){
			     System.out.println(e.getMessage());
			  }
				}
        		}
        	if(driver.findElements(By.xpath(next)).size()!=0)
    		{	
    			driver.findElement(By.xpath(next)).click();
    			Thread.sleep(3000);
    		}else {
        		condn=false;
        	}
		/*if(driver.findElements(By.xpath(next)).size()!=0)
		{	
			driver.findElement(By.xpath(next)).click();
			Thread.sleep(3000);
		}*/
        	}
	
        }catch(Exception e) {
        	e.printStackTrace();
        //	driver.navigate().refresh();
        }
	finally {
	FileOutputStream fileOut = null;
	// Write the output to a file
		try {
			 fileOut = new FileOutputStream("/Users/swetha/"+input+".xlsx");
		}catch(Exception e) {
			e.printStackTrace();
		}
	
		workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
	}
	}

}
