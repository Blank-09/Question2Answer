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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
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
    private final String EXECUTABLE_PATH = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
    private final String USER_DATA_DIR = "C:\\Users\\<username>\\AppData\\Local\\Google\\Chrome\\User Data";
    private final String PROFILE_DIRECTORY = "Profile 1";

    private final String QUESTION_SHEET_PATH = "./assets/sheets/questions.xlsx";
    private final String REPORT_PATH = "./out/reports/index.html";
    private final String LOGGER_PATH = "./out/logs/app.log";
    private final String SCREENSHOT_PATH = "./out/screenshots/";

    WebDriver driver;
    Actions actions;
    ExtentReports reports;
    Wait<WebDriver> wait;
    List<Question> questions;

    public class Question {

        String question, additionalInfo;
        int marks, sno;

        Question(Row row) {
            this.sno = (int) row.getCell(0).getNumericCellValue();
            this.question = row.getCell(1).getStringCellValue();
            this.marks = (int) row.getCell(2).getNumericCellValue();

            Cell additionalInfo = row.getCell(3);
            if (additionalInfo != null) {
                this.additionalInfo = additionalInfo.getStringCellValue();
            }
        }

        @Override
        public String toString() {
            return "S.No         :" + sno + "\n" +
                    "Question     :" + question + "\n" +
                    "Marks        :" + marks + "\n" +
                    "Additional Information :" + additionalInfo;
        }
    }

    @BeforeTest
    public void setupDriver() {
        ChromeOptions options = new ChromeOptions();
        options.setBinary(EXECUTABLE_PATH);
        options.addArguments("--user-data-dir=" + USER_DATA_DIR);
        options.addArguments("--profile-directory=" + PROFILE_DIRECTORY);

        this.driver = new ChromeDriver(options);
        this.actions = new Actions(driver);
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(30));

    }

    @BeforeTest
    public void setupExcel() throws IOException {

        Workbook workbook = new XSSFWorkbook(QUESTION_SHEET_PATH);
        Sheet sheet = workbook.getSheetAt(0);

        int rowCount = sheet.getLastRowNum();
        questions = new ArrayList<>();

        for (int i = 1; i <= rowCount; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                questions.add(new Question(row));
            }
        }

        workbook.close();
    }

    @Test(priority = 1)
    public void getAnswersFromChat() throws InterruptedException {

        driver.get(CHATGPT_URL);

        // Assuming 'driver' and 'questions' are properly instantiated

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(200));
        By textareaLocator = By.id("prompt-textarea");
        By submitButtonLocator = By.cssSelector("button[data-testid=send-button]");

        for (Question question : questions) {
            String questionText = question.question;
            String marks = ". Answer the question as " + question.marks + " marks";
            String additionalInfo = "";

            if (question.additionalInfo != null && !question.additionalInfo.isEmpty())
                additionalInfo = " and add the following information: " + question.additionalInfo;

            String prompt = questionText + marks + additionalInfo;

            // Entering question text into the text area
            driver.findElement(textareaLocator).sendKeys(prompt);

            // Clicking the button to submit the question
            driver.findElement(submitButtonLocator).click();

            // Waiting for the button to disappear and then reappear
            wait.until(ExpectedConditions.invisibilityOfElementLocated(submitButtonLocator));
            wait.until(ExpectedConditions.presenceOfElementLocated(submitButtonLocator));
            Thread.sleep(5000);
        }
    }

    @AfterTest
    public void wrapUp() {
        driver.quit();
        // reports.flush();
    }

    // All your private function goes here lexigraphically
    private void takeScreenshot(String name) throws IOException {
        File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        FileUtils.copyFile(screenshotFile, new File(SCREENSHOT_PATH + name + "_" + timestamp + ".png"));
    }

}
