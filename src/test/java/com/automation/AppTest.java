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
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
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
    private final String USER_DATA_DIR = "C:\\Users\\<UserName>\\AppData\\Local\\Google\\Chrome\\User Data";
    private final String PROFILE_DIRECTORY = "Profile 1";

    private final String QUESTION_SHEET_PATH = "./assets/sheets/questions.xlsx";
    private final String REPORT_PATH = "./out/reports/index.html";
    private final String LOGGER_PATH = "./out/logs/app.log";
    private final String SCREENSHOT_PATH = "./out/screenshots/";
    Logger logger = LogManager.getLogger(getClass());

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
        logger.info("Setting up Chrome driver...");
        logger.info("Creating ChromeOptions object...");
        ChromeOptions options = new ChromeOptions();
        logger.info("Setting executable path...");
        options.setBinary(EXECUTABLE_PATH);
        logger.info("Adding user data directory argument...");
        options.addArguments("--user-data-dir=" + USER_DATA_DIR);
        logger.info("Adding profile directory argument...");
        options.addArguments("--profile-directory=" + PROFILE_DIRECTORY);

        logger.info("Creating ChromeDriver object...");
        this.driver = new ChromeDriver(options);
        logger.info("Creating Actions object...");
        this.actions = new Actions(driver);
        logger.info("Creating WebDriverWait object...");
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        logger.info("Chrome driver setup complete.");

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

        logger.info("Getting answers from chat...");
    try {
        logger.info("Navigating to ChatGPT URL...");
        driver.get(CHATGPT_URL);

        logger.info("Creating WebDriverWait object...");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(200));

        logger.info("Creating textareaLocator object...");
        By textareaLocator = By.id("prompt-textarea");

        logger.info("Creating submitButtonLocator object...");
        By submitButtonLocator = By.cssSelector("button[data-testid=send-button]");

        for (Question question : questions) {
            logger.info("Getting question text...");
            String questionText = question.question;

            logger.info("Creating marks string...");
            String marks = ". Answer the question as " + question.marks + " marks";

            String additionalInfo = "";

            if (question.additionalInfo!= null &&!question.additionalInfo.isEmpty()) {
                logger.info("Adding additional information...");
                additionalInfo = " and add the following information: " + question.additionalInfo;
            }

            logger.info("Creating prompt string...");
            String prompt = questionText + marks + additionalInfo;

            // Entering question text into the text area
            logger.info("Entering question text into text area...");
            driver.findElement(textareaLocator).sendKeys(prompt);

            // Clicking the button to submit the question
            logger.info("Clicking submit button...");
            driver.findElement(submitButtonLocator).click();

            // Waiting for the button to disappear and then reappear
            logger.info("Waiting for submit button to disappear...");
            wait.until(ExpectedConditions.invisibilityOfElementLocated(submitButtonLocator));
            logger.info("Waiting for submit button to reappear...");
            wait.until(ExpectedConditions.presenceOfElementLocated(submitButtonLocator));
            logger.info("Waiting for 5 seconds...");
            Thread.sleep(5000);
        }
    } catch (Exception e) {
        logger.error("An error occurred while getting answers from chat: ", e);
    }
    logger.info("Getting answers from chat complete.");
    }

    @AfterTest
    public void wrapUp() {
        driver.quit();
        logger.info("Quitting Driver");
        // reports.flush();
    }

    // All your private function goes here lexigraphically
    private void takeScreenshot(String name) throws IOException {
        File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        FileUtils.copyFile(screenshotFile, new File(SCREENSHOT_PATH + name + "_" + timestamp + ".png"));
    }

}
