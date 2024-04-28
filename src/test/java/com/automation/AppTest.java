package com.automation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class AppTest {

    private final String CHATGPT_URL = "https://chat.openai.com/";

    // Update the path to your Chrome profile directory
    private final String USER_DATA_DIR = "C:\\Users\\Lenovo\\AppData\\Local\\Google\\Chrome\\User Data\\";
    private final String PROFILE_DIRECTORY = "Profile 5";
    private final String EXECUTABLE_PATH = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";

    private final String QUESTION_SHEET_PATH = "./assets/sheets/questions.xlsx";
    private final String REPORT_PATH = "./out/reports/index.html";
    private final String LOGGER_PATH = "./out/logs/app.log";
    private final String SCREENSHOT_PATH = "./out/screenshots/";

    WebDriver driver;
    Actions actions;
    ExtentReports reports;
    Wait<WebDriver> wait;
    XSSFWorkbook workbook;
    XSSFSheet sheet;


    Question questions;
    List<Question> questionsList;

    public class Question {

        String question,additionalInfo;
        int marks,sno;

        Question(Row row) {
            Cell cellsno = row.getCell(0);
            Cell cellQuestion = row.getCell(1); 
            Cell cellMarks = row.getCell(2);
    
            if (cellsno != null && cellsno.getCellType() == CellType.NUMERIC && 
                cellQuestion != null && cellQuestion.getCellType() == CellType.STRING &&
                cellMarks != null && cellMarks.getCellType() == CellType.NUMERIC) {
    
                this.sno = (int)cellsno.getNumericCellValue();
                this.question = cellQuestion.getStringCellValue();
                this.marks = (int) cellMarks.getNumericCellValue();
            } 
            else {
                this.sno = 0;
                this.question = "Default Question";
                this.marks = 0;
            }
        }

        @Override
        public String toString(){

            return "S.No         :"+sno+"\n"+
                   "Question     :"+question+"\n" +
                   "Marks        :"+marks+"\n" +
                   "Additional Information :"+additionalInfo+"\n";
        }
    }


    @BeforeTest
    public void setupDriver() {
        ChromeOptions options = new ChromeOptions();
        options.setBinary(EXECUTABLE_PATH);
        options.addArguments("--user-data-dir=" + USER_DATA_DIR);
        options.addArguments("--profile-directory=" + PROFILE_DIRECTORY);

        this.driver = new ChromeDriver();
        this.actions = new Actions(driver);
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        driver.get(CHATGPT_URL);
    }

    @BeforeTest
    public void setUpExcel() throws IOException {
        
        workbook = new XSSFWorkbook(QUESTION_SHEET_PATH);
        sheet = workbook.getSheetAt(0);

        questionsList = new ArrayList<>();

        int rowCount=sheet.getLastRowNum();
        questionsList = new ArrayList<>();

        for(int i=1;i<=rowCount;i++)
        {
            Row row = sheet.getRow(i);
            if(row!=null){
                questionsList.add(new Question(row));
            }
        }
    
        workbook.close();
    }


    @Test
    public void shouldAnswerWithTrue() {
        System.out.println("Test is running...");
    }

    @AfterTest
    public void wrapUp() {
        driver.quit();
        reports.flush();
    }

    // All your private function goes here lexigraphically
    private void takeScreenshot(String name) throws IOException {
        File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        FileUtils.copyFile(screenshotFile, new File(SCREENSHOT_PATH + name + "_" + timestamp + ".png"));
    }

}
