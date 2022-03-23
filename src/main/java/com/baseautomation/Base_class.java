package com.baseautomation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.invoke.SwitchPoint;
import java.time.Duration;
import java.util.concurrent.TimeUnit;
import java.util.logging.FileHandler;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Base_class {

	public static WebDriver driver;
	public static String value;

	public static WebDriver browser_configuration(String type) {

		if (type.equalsIgnoreCase("chrome")) {

			System.setProperty("webdriver.chrome.driver",
					System.getProperty("user.dir") + "\\Driver\\chromedriver.exe");

			driver = new ChromeDriver();

		} else if (type.equalsIgnoreCase("edge")) {

			System.setProperty("webdriver.edge.driver", System.getProperty("user.dir") + "\\Driver\\msedgedriver.exe");

			driver = new EdgeDriver();

		}

		driver.manage().window().maximize();
		return driver;

	}

	public static void geturl(String url) {

		driver.get(url);

	}

	public static void clickOnelement(WebElement element) {

		element.click();

	}

	public static void inputValueelement(WebElement element, String data) {

		element.sendKeys(data);

	}

	public static void clearing(WebElement element) {
		element.clear();
	}

	public static void sleep(int milliseconds) throws Throwable {

		Thread.sleep(milliseconds);
	}

	public static void Action(String type, WebElement src, WebElement dest) throws Exception {

		Actions act = new Actions(driver);
		if (type.equalsIgnoreCase("move")) {

			act.moveToElement(src).build().perform();

		} else if (type.equalsIgnoreCase("drag")) {

			act.dragAndDrop(src, dest).build().perform();
		}

	}

	public static void javascript(String parameter, WebElement element) {
		JavascriptExecutor execute = (JavascriptExecutor) driver;

		execute.executeScript(parameter, element);

	}

	public static void Screenshot(String type, String path) throws Throwable {
		TakesScreenshot screenshot = (TakesScreenshot) driver;
		File source = screenshot.getScreenshotAs(OutputType.FILE);
		File destination = new File(System.getProperty("user.dir") + path);
		if (type.equalsIgnoreCase("utils")) {
			FileUtils.copyFile(source, destination);
		} else if (type.equalsIgnoreCase("handler")) {
			org.openqa.selenium.io.FileHandler.copy(source, destination);
		} else {
			System.out.println("please pass correct screenshot type");
		}

	}

	public static void frames(WebElement element, String type, String value) {

		switch (type) {
		case "index":
			driver.switchTo().frame(Integer.parseInt(value));
			break;
		case "id":
			driver.switchTo().frame(value);
			break;
		case "element":
			driver.switchTo().frame(value);
			break;
		default:
			System.out.println("please pass correct frames type");
			break;
		}
	}

	public static void defaultcontent() {

		driver.switchTo().defaultContent();
	}

	public static void navigate(String navigatetype, String url) {
		switch (navigatetype) {
		case "to":
			driver.navigate().to(url);
			break;
		case "forward":
			driver.navigate().forward();
			break;
		case "back":
			driver.navigate().back();
			break;
		case "refresh":
			driver.navigate().refresh();
			break;

		default:
			System.out.println("please pass correct navigate type...");
			break;
		}

	}

	public static void Alerts(String type, String text) {

		Alert a = driver.switchTo().alert();

		try {
			switch (type) {
			case "ok":
				a.accept();
				break;
			case "dismiss":
				a.dismiss();
				break;
			case "text":
				a.sendKeys(text);
				break;
			default:
				System.out.println("please pass correct alert type");
				break;

			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void selectdropdownvalue(WebElement element, String type, String value) {
		Select select = new Select(element);
		switch (type) {
		case "index":
			select.selectByIndex(Integer.parseInt(value));
			break;
		case "value":
			select.selectByValue(value);
			break;
		case "visibletext":
			select.selectByVisibleText(value);
			break;
		default:
			System.out.println("please pass correct dropdown value");
			break;
		}

	}

	public static String Particular_Cell_Data(String path, int row_index, int cell_index) throws IOException {

		File f = new File(path);

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet s = wb.getSheetAt(0);

		Row r = s.getRow(row_index);

		Cell c = r.getCell(cell_index);

		CellType type = c.getCellType();

		if (type.equals(CellType.STRING)) {

			value = c.getStringCellValue();

		} else if (type.equals(CellType.NUMERIC)) {

			double cellValue = c.getNumericCellValue();

			int val = (int) cellValue;

			String v = Integer.toString(val);

			value = String.valueOf(v);

		}
		return value;

	}

	public static void implicitwait(int time, TimeUnit format) {
		driver.manage().timeouts().implicitlyWait(time, format);

	}

	public static void explicitwait(int time) {
		WebDriverWait wait = new WebDriverWait(driver, time);

	}

	public static void close() {
		driver.close();

	}

}
