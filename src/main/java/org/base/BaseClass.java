package org.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	public static WebDriver driver;

	public static WebDriver browserLaunch(String browserName) {
		switch (browserName) {
		case "Chrome":
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			break;
		case "Firefox":
			System.setProperty("webdriver.gecko.driver",
					"C:\\Users\\ManoKutty\\eclipse-workspace\\MavenProject02.00PM_June21\\Drivers\\geckodriver.exe");
			driver = new FirefoxDriver();
			break;
		case "Ie":
			System.setProperty("webdriver.ie.driver",
					"C:\\Users\\ManoKutty\\eclipse-workspace\\MavenProject02.00PM_June21\\Drivers\\IEDriverServer.exe");
			driver = new InternetExplorerDriver();
			break;
		case "Edge":
			System.setProperty("webdriver.edge.driver",
					"C:\\Users\\ManoKutty\\eclipse-workspace\\MavenProject02.00PM_June21\\Drivers\\msedgedriver.exe");
			driver = new EdgeDriver();
			break;

		default:
			System.err.println("Invalid Browser");
			break;
		}

		return driver;
	}

	public static void implicitwait(long sec) {
		driver.manage().timeouts().implicitlyWait(sec, TimeUnit.SECONDS);

	}

	public static void launchUrl(String url) {
		driver.get(url);
	}

	public static void fillTextBox(WebElement element, String value) {
		element.sendKeys(value);

	}

	public static void btnClick(WebElement e) {
		e.click();
	}

	public static void browserQuit() {
		driver.quit();

	}

	public static String getCurrentUrl() {
		return driver.getCurrentUrl();
	}

	public static String getAttribute(WebElement e) {
		String sat = e.getAttribute("value");
		return sat;
	}

	public static void moveToElement(WebElement e) {
		Actions a = new Actions(driver);
		a.moveToElement(e).perform();
	}

	public static void dragANdDrop(WebElement src, WebElement des) {
		Actions a = new Actions(driver);
		a.dragAndDrop(src, des).perform();
	}

	public static void selectByIndex(WebElement element, int index) {
		Select s = new Select(element);
		s.selectByIndex(index);
	}

	public static WebElement findElement(String locatorType, String value) {
		WebElement element = null;
		if (locatorType.equals("id")) {

			element = driver.findElement(By.id(value));
		} else if (locatorType.equals("name")) {

			element = driver.findElement(By.name(value));
		} else if (locatorType.equals("xpath")) {

			element = driver.findElement(By.xpath(value));
		}
		return element;

	}

	public static void jsSendKeys(WebElement e, String input) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].setAttribute('value','" + input + "')", e);
	}

	public static void screenShot(String image) throws IOException {
		TakesScreenshot tk = (TakesScreenshot) driver;
		File src = tk.getScreenshotAs(OutputType.FILE);
		File des = new File(
				"C:\\Users\\ManoKutty\\eclipse-workspace\\MavenProject02.00PM_June21\\src\\test\\resources\\ScreenShot\\"
						+ image + ".png");
		FileUtils.copyFile(src, des);

	}

	public static void windowsHandling() {
		String parentId = driver.getWindowHandle();
		Set<String> allId = driver.getWindowHandles();
		for (String eachId : allId) {
			if (!parentId.equals(eachId)) {
				driver.switchTo().window(eachId);
			}

		}

	}

	public static String getData(String sheetName, int rowNo, int cellNo) throws IOException {
		File loc = new File(
				"C:\\Users\\ManoKutty\\eclipse-workspace\\MavenProject02.00PM_June21\\src\\test\\resources\\Excel\\Data.xlsx");
		FileInputStream st = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(st);
		Sheet s = w.getSheet(sheetName);
		Row row = s.getRow(rowNo);
		Cell cell = row.getCell(cellNo);
		int type = cell.getCellType();
		String value = null;
		if (type == 1) {
			value = cell.getStringCellValue();
		} else {
			if (DateUtil.isCellDateFormatted(cell)) {
				value = new SimpleDateFormat("dd-MMM-yyyy").format(cell.getDateCellValue());

			} else {
				value = String.valueOf((long) cell.getNumericCellValue());
			}
		}
		return value;

	}

	public static void main(String[] args) throws IOException {
		String data = getData("new", 3, 1);
		System.out.println(data);
	}

}
