package panda.selenium;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.InputEvent;
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

public class Selenium2Example {

	private static final String FILE_NAME = "files/listFilm.xlsx";

	private static final String SUFFIX_SEARCH = " poster film";

	@SuppressWarnings("resource")
	public static void main(String[] args) throws AWTException, InterruptedException, IOException {

		System.out.println(" =============> Start Dowload Image (Selenium 2) <=============");

		System.setProperty("webdriver.chrome.driver", "driver/chromedriver.exe");

		WebDriver driver = new ChromeDriver();

		//// Advance image search with size large 640x480

		driver.get(
				"https://www.google.com.vn/search?q=image&rlz=1C1CHZL_viVN749VN749&tbs=isz:lt,islt:vga&tbm=isch&source=lnt&sa=X&ved=0ahUKEwjijOvphfzVAhVMOI8KHeJrCH0QpwUIGQ&biw=1273&bih=659&dpr=1");

		FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet sheet = workbook.getSheetAt(0);

		int colTitle = 6;
		int colImgName = 7;

//		for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
		for (int rowIndex = 1294; rowIndex >= 1245; rowIndex--) {
			Row row = sheet.getRow(rowIndex);
			if (row != null) {
				Cell cell = row.getCell(colTitle);
				Cell cellImgName = row.getCell(colImgName);

				if (cell != null) {
					// Found column and there is value in the cell.
					String nameFilm = cell.getStringCellValue() == null ? "fox baby" : cell.getStringCellValue();

					SimpleDateFormat myFormat = new SimpleDateFormat("yyyy-MM-dd");
					String nameImg = cellImgName.getStringCellValue() == null ? myFormat.format(new Date())
							: cellImgName.getStringCellValue();

					//// Call selenium
					System.out.println(nameFilm);

					WebElement searchBox = driver.findElement(By.name("q"));
					searchBox.clear();
					searchBox.sendKeys(nameFilm + SUFFIX_SEARCH);
					searchBox.sendKeys(Keys.RETURN);

					// Point coordinates = searchBox.getLocation();
					Robot robot = new Robot();
					robot.mouseMove(75, 400);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					Thread.sleep(800);

					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					Thread.sleep(800);

					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					Thread.sleep(800);

					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					Thread.sleep(800);

					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					Thread.sleep(500);

					robot.mouseMove(200, 500);
					Thread.sleep(500);

					robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
					robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
					Thread.sleep(500);

					// 8 lan - Select Save as...
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(100);
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(100);
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(100);
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(100);
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(100);
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(100);
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(100);
					robot.keyRelease(KeyEvent.VK_DOWN);

					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(200);

					StringSelection selection = new StringSelection(nameImg);
					Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
					clipboard.setContents(selection, selection);

					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					Thread.sleep(300);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_V);

					Thread.sleep(800);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					
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
