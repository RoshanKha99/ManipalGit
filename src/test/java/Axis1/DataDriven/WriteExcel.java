package Axis1.DataDriven;


	import java.io.File;
	import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
	import java.util.concurrent.TimeUnit;
	 
	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	import org.openqa.selenium.By;
	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.chrome.ChromeDriver;
	import org.testng.annotations.Test;
	 
	public class WriteExcel {
		WebDriver driver;
		XSSFWorkbook workbook;
		XSSFSheet sheet;
		XSSFCell cell;
		private String title;

		@SuppressWarnings("deprecation")
		@Test
		public void FBlogin() throws IOException {

				System.setProperty("webdriver.chrome.driver",
						"C:\\Users\\Roshan Khapekar\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
				WebDriver driver = new ChromeDriver();
				
				driver.get("http://www.facebook.com/");
				driver.manage().window().maximize();
				
			// Import excel sheet
			File src = new File("C:\\Users\\Roshan Khapekar\\Downloads\\apache-maven-3.9.6-bin\\DataDriven\\TestData.xlsx");
			// load the file
			FileInputStream fis = new FileInputStream(src);
			// load the work book

			workbook = new XSSFWorkbook(fis);
			// access the sheet from the work book
			sheet = workbook.getSheetAt(0);
			for (int i = 1; i<=sheet.getLastRowNum(); i++) {
				// import the data from email
				cell = sheet.getRow(i).getCell(0);
				driver.findElement(By.xpath("//input[@name = 'email']")).clear();
				driver.findElement(By.xpath("//input[@name = 'email']")).sendKeys(cell.getStringCellValue());

				// import the data for the password 

				cell = sheet.getRow(i).getCell(1);
				driver.findElement(By.xpath("//input[@id = 'pass']")).clear();
				driver.findElement(By.xpath("//input[@id = 'pass']")).sendKeys(cell.getStringCellValue());

				// To write in excel
				
				FileOutputStream fos = new FileOutputStream(src);
				
				
				sheet.getRow(i).createCell(2).setCellValue(title);
				
				workbook.write(fos);
				fos.close();
			}}
	}


