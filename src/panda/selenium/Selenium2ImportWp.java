package panda.selenium;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Selenium2ImportWp {

	private static final String FILE_NAME = "files/listFilm.xlsx";

	private static final String USER_NAME = "gennim";

	private static final String USER_PASS = "Minneg2908";

	@SuppressWarnings("resource")
	public static void main(String[] args) throws AWTException, InterruptedException, IOException {

		System.out.println("Start Selenium");

		System.setProperty("webdriver.chrome.driver", "driver/chromedriver.exe");

		WebDriver driver = new ChromeDriver();

		//// Advance image search with size large 640x480
		driver.get("http://genncinnema.byethost32.com/wp-admin");

		Thread.sleep(2000);

		//// login
		WebElement inputLogin = driver.findElement(By.id("user_login"));
		inputLogin.sendKeys(USER_NAME);

		WebElement inputPass = driver.findElement(By.id("user_pass"));
		inputPass.sendKeys(USER_PASS);

		WebElement btnLogin = driver.findElement(By.id("wp-submit"));
		btnLogin.click();

		Thread.sleep(2000);

		FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet sheet = workbook.getSheetAt(0);

		int colTitle = 6;
		int colImgName = 7;

		// for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
		for (int rowIndex = sheet.getLastRowNum(); rowIndex >= 1325; rowIndex--) {
			Row row = sheet.getRow(rowIndex);
			if (row != null) {
				Cell cellFilmName = row.getCell(colTitle);
				Cell cellImgName = row.getCell(colImgName);

				if (cellFilmName != null) {
					// Found column and there is value in the cell.
					String nameFilm = cellFilmName.getStringCellValue() == null ? "fox baby" : cellFilmName.getStringCellValue();

					SimpleDateFormat myFormat = new SimpleDateFormat("yyyy-MM-dd");
					String nameImg = cellImgName.getStringCellValue() == null ? myFormat.format(new Date())
							: cellImgName.getStringCellValue();

					//// Call selenium
					System.out.println(nameFilm);

					//// Post
					WebElement liPost = rowIndex == sheet.getLastRowNum() ? driver.findElement(By.className("wp-menu-name")) : driver.findElement(By.className("wp-menu-name"));
					liPost.click();
					Thread.sleep(2000);

					WebElement inputSearchPost = driver.findElement(By.id("post-search-input"));
					inputSearchPost.sendKeys(nameFilm);
					inputSearchPost.sendKeys(Keys.RETURN);
					Thread.sleep(1000);

					WebElement aRowTitle = driver.findElement(By.className("row-title"));
					aRowTitle.click();
					Thread.sleep(1000);

					WebElement aRemoveThumb = driver.findElement(By.id("remove-post-thumbnail"));
					if (aRemoveThumb != null) {
						aRemoveThumb.click();
					}

					WebElement aSetThumb = driver.findElement(By.id("set-post-thumbnail"));
					aSetThumb.click();

					Robot robot = new Robot();
					robot.keyPress(KeyEvent.VK_TAB);
					robot.keyPress(KeyEvent.VK_TAB);
					robot.keyRelease(KeyEvent.VK_TAB);

					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);

					//// Button upload
					WebElement btnUpload = driver.findElement(By.id("__wp-uploader-id-1"));
					btnUpload.click();
					Thread.sleep(500);

					StringSelection selection = new StringSelection(nameImg);
					Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
					clipboard.setContents(selection, selection);

					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					Thread.sleep(300);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_V);

					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(5000);

					WebElement btnSelect = driver.findElement(By.className("media-button-select"));
					btnSelect.click();
					Thread.sleep(1000);

					WebElement inputPublish = driver.findElement(By.id("publish"));
					inputPublish.sendKeys(Keys.ENTER);

					Thread.sleep(2000);
				}
			}
		}

		driver.quit();

		/*
		 * for (String str : arr) {
		 * 
		 * WebElement searchBox = driver.findElement(By.name("q")); searchBox.clear();
		 * searchBox.sendKeys(str + SUFFIX_SEARCH); searchBox.sendKeys(Keys.RETURN);
		 * 
		 * // Point coordinates = searchBox.getLocation(); Robot robot = new Robot();
		 * robot.mouseMove(75, 400); robot.mousePress(InputEvent.BUTTON1_MASK);
		 * Thread.sleep(1000);
		 * 
		 * robot.mousePress(InputEvent.BUTTON1_MASK);
		 * robot.mouseRelease(InputEvent.BUTTON1_MASK); Thread.sleep(1000);
		 * 
		 * robot.mousePress(InputEvent.BUTTON1_MASK);
		 * robot.mouseRelease(InputEvent.BUTTON1_MASK); Thread.sleep(1000);
		 * 
		 * robot.mousePress(InputEvent.BUTTON1_MASK);
		 * robot.mouseRelease(InputEvent.BUTTON1_MASK); Thread.sleep(1000);
		 * 
		 * robot.mousePress(InputEvent.BUTTON1_MASK);
		 * robot.mouseRelease(InputEvent.BUTTON1_MASK); Thread.sleep(500);
		 * 
		 * robot.mouseMove(200, 500); Thread.sleep(1000);
		 * 
		 * robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
		 * robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK); Thread.sleep(1000);
		 * 
		 * // 8 lan - Select Save as... robot.keyPress(KeyEvent.VK_DOWN);
		 * Thread.sleep(200); robot.keyPress(KeyEvent.VK_DOWN); Thread.sleep(200);
		 * robot.keyPress(KeyEvent.VK_DOWN); Thread.sleep(200);
		 * robot.keyPress(KeyEvent.VK_DOWN); Thread.sleep(200);
		 * robot.keyPress(KeyEvent.VK_DOWN); Thread.sleep(200);
		 * robot.keyPress(KeyEvent.VK_DOWN); Thread.sleep(200);
		 * robot.keyPress(KeyEvent.VK_DOWN); Thread.sleep(200);
		 * robot.keyRelease(KeyEvent.VK_DOWN);
		 * 
		 * robot.keyPress(KeyEvent.VK_ENTER); Thread.sleep(200);
		 * 
		 * StringSelection selection = new StringSelection(str.replace(" ", "_"));
		 * Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
		 * clipboard.setContents(selection, selection);
		 * 
		 * robot.keyPress(KeyEvent.VK_CONTROL); robot.keyPress(KeyEvent.VK_V);
		 * Thread.sleep(300); robot.keyRelease(KeyEvent.VK_CONTROL);
		 * robot.keyRelease(KeyEvent.VK_V);
		 * 
		 * Thread.sleep(1000); robot.keyPress(KeyEvent.VK_ENTER);
		 * robot.keyRelease(KeyEvent.VK_ENTER);
		 * 
		 * // driver.close(); }
		 * 
		 * 
		 * driver.quit();
		 */
	}

}
