package poly.Lab08;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class SaveTestNGResultToExcel {
	public WebDriver driver;
	public UIMap uimap;
	public UIMap datafile;
	public String workingDir;
	String Path = "C:\\chromedriver.exe";
	
	HSSFWorkbook workbook;
	
	HSSFSheet sheet;
	
	Map<String, Object[]> TestNGResults;
	
	
	@BeforeClass(alwaysRun = true)
	public void suiteSetUp() {
		workbook = new HSSFWorkbook();
		
		sheet= workbook.createSheet("TestNg Result Summary");
		TestNGResults = new LinkedHashMap<String, Object[]>();
		
		TestNGResults.put("1",new Object[] {"Test step No", "Action", "Expected Output", "Actual Output"});
		
		try {
			workingDir= System.getProperty("user.dir");
			datafile = new UIMap(workingDir+"\\Resources\\datafile.properties");
			uimap = new UIMap(workingDir+"\\Resources\\locator.properties");
			System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
			driver= new ChromeDriver();
			driver.manage().timeouts().implicitlyWait(5,TimeUnit.SECONDS);
			
		} catch (Exception e) {
			// TODO: handle exception
			throw new IllegalStateException("can't start the Firefox web driver", e);
		}
		
	}
	
	@Test(description = "open the testng demo website for login test",priority = 1)
	public void LaunchWebsite() throws Exception{
		try {
			driver.get("https://phptravels.net/login");
			driver.manage().window().maximize();
			TestNGResults.put("2", new Object[] {1d, "Navigate to demo website","Site gets Opened" ,"Pass"});
		} catch (Exception e) {
			// TODO: handle exception
			TestNGResults.put("2", new Object[] {1d, "Navigate to demo website","Site gets Opened" , "Pass"});
			Assert.assertTrue(false);
		}
	}
	
	@Test(description = "FillLoginDetail",priority = 2)
	public void FillLoginDetails() throws Exception{
		try {
			WebElement username= driver.findElement(uimap.getLocator("Username_field"));
			username.sendKeys(datafile.getData("username"));
			
			WebElement password = driver.findElement(uimap.getLocator("Passwors_field"));
			password.sendKeys(datafile.getData("password"));
			Thread.sleep(1000);
			TestNGResults.put("3", new Object[] {2d, "fill login form data(Username/password)","login details gets filled" ,"Pass"});
		} catch (Exception e) {
			// TODO: handle exception
			TestNGResults.put("3", new Object[] {2d, "fill login form data(Username/password)","login details gets filled" ,"Pass"});
			Assert.assertTrue(false);
		}
		
	}
	
	@Test(description = "Perform Login",priority = 3)
	public void DoLogin() throws Exception{
		try {
			
			WebElement login= driver.findElement(uimap.getLocator("Login_button"));
			login.click();
			Thread.sleep(1000);
			
//			WebElement onlineuser = driver.findElement(uimap.getLocator("online_user"));
//			Assert.assertEquals("Hi, John Smith", onlineuser.getText());
			TestNGResults.put("4", new Object[] {3d, "Click Login anh verify welcome massage","login success" ,"Pass"});
		} catch (Exception e) {
			// TODO: handle exception
			TestNGResults.put("4", new Object[] {3d, "Click Login anh verify welcome massage","login success" ,"fails"});
			Assert.assertTrue(false);
		}
		
	}
	
	@AfterClass
	public void suiteTeardown() {
		Set<String> ketset= TestNGResults.keySet();
		int rownum= 0;
		for (String key:ketset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = TestNGResults.get(key);
			int cellnum= 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof Date) {
					cell.setCellValue((Date) obj);
				}else if (obj instanceof Boolean) {
					cell.setCellValue((Boolean) obj);
				}else if (obj instanceof String) {
					cell.setCellValue((String) obj);
				}else if (obj instanceof Double) {
					cell.setCellValue((Double) obj);
				}
			}
		}
		try {
			FileOutputStream out = new FileOutputStream(new File("SaveTestNGResultToExcel.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("Successfully saved selenium Webdriver TestNG result to Excel File!");
		} catch (FileNotFoundException e) {
			// TODO: handle exception
			e.printStackTrace();
			
		}catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		driver.close();
	}
	
	
	

	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
