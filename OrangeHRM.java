package DDF;

import java.io.IOException;
import java.time.Duration;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class OrangeHRM {

	WebDriver driver ;
	
	@BeforeTest
	public void setUp()
	{
		driver = new ChromeDriver();
		
		driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");
	
	//	driver.manage().window().maximize();
		
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
	}
	
	
	@AfterTest
	public void tearDown() {
		driver.close();
	}
	
	@Test(dataProvider = "D")
	public void setData(String un, String pwd) throws InterruptedException {
		Thread.sleep(3000);
		
		driver.findElement(By.name("username")).clear();
		driver.findElement(By.name("username")).sendKeys(un);
		
		driver.findElement(By.name("password")).clear();
		driver.findElement(By.name("password")).sendKeys(pwd);
		
		driver.findElement(By.xpath("//button[normalize-space(text()='Login')]")).click();
	}
	
	@DataProvider(name="D")
	public Object[][] getData() throws IOException, InvalidFormatException {
		
		ExcelUtil_2 util2 = new ExcelUtil_2("D:\\DDF sheet\\Demo.xlsx");
		
		int rowCount = util2.rowCount("Sheet3");
		System.out.println("Rowcount of sheet3 :"+rowCount);
		
		short columCount = util2.columCount("Sheet3");
		System.out.println("Column count of sheet 3 :"+ columCount);
		
		Object[][] obj = new Object[rowCount][columCount];
		
		for (int i = 0; i <rowCount; i++) {
			for (int j = 0; j <columCount; j++) {				
				obj[i][j] =util2.getExcelData("Sheet3", i, j);
			}
		}
		return obj;
	}
	
	
}
