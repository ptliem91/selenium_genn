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

public class Selenium2DownloadFromImdb {

	private static final String FILE_NAME = "files/listFilm.xlsx";

	private static final String SUFFIX_SEARCH = "";

	@SuppressWarnings("resource")
	public static void main(String[] args) throws AWTException, InterruptedException, IOException {

		System.out.println(" =============> Start Dowload Image (Selenium 2) <=============");

		System.setProperty("webdriver.chrome.driver", "driver/chromedriver.exe");

		WebDriver driver = new ChromeDriver();

		driver.get("http://www.imdb.com/");
		
		FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet sheet = workbook.getSheetAt(0);

		int colTitle = 6;
		int colImgName = 7;
		
		int count = 0;
//		for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
		for (int rowIndex = 1248; rowIndex >= 1200; rowIndex--) {
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
					
					driver.navigate().to("http://www.imdb.com/?ref_=nv_home");

					WebElement searchBox = driver.findElement(By.name("q"));
					searchBox.clear();
					searchBox.sendKeys(nameFilm + SUFFIX_SEARCH);
					searchBox.sendKeys(Keys.RETURN);

					//// Point coordinates = searchBox.getLocation();
					Robot robot = new Robot();
					robot.mouseMove(75, 380);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					Thread.sleep(3000);
					
					robot.mouseMove(75, 420);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					Thread.sleep(1000);

					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					Thread.sleep(800);
										
					robot.mouseMove(350, 380);
					Thread.sleep(1000);
					
					robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
					robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
					Thread.sleep(1000);
					
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(300);
					robot.keyPress(KeyEvent.VK_DOWN);
					Thread.sleep(300);
					
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(300);
//
					StringSelection selection = new StringSelection(nameImg);
					Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
					clipboard.setContents(selection, selection);

					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					Thread.sleep(300);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_V);

					Thread.sleep(500);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					
					robot.keyPress(KeyEvent.VK_BACK_SPACE);
					robot.keyRelease(KeyEvent.VK_BACK_SPACE);
					
					Thread.sleep(3000);
					
					count++;
				}
			}
		}

		driver.quit();
		
		System.out.println(" =============> End Dowload Image (Selenium 2) <=============");
		System.out.println(" =============> " + count);
	}

}
