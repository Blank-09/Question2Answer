package com.automation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
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
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Pdf;
import org.openqa.selenium.PrintsPage;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.print.PrintOptions;
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
    private final String USER_DATA_DIR = "C:\\Users\\<username>\\AppData\\Local\\Google\\Chrome\\User Data\\";
    private final String PROFILE_DIRECTORY = "Profile 1";

    private final String QUESTION_SHEET_PATH = "./assets/sheets/questions.xlsx";
    private final String REPORT_PATH = "./out/reports/index.html";
    private final String LOGGER_PATH = "./out/logs/app.log";
    private final String SCREENSHOT_PATH = "./out/screenshots/";
    private final String PDF_PATH = "./out/result/";

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
        ChromeOptions options = new ChromeOptions();
        logger.info("Creating ChromeOptions object...");
        options.setBinary(EXECUTABLE_PATH);
        logger.info("Setting executable path...");
        options.addArguments("--user-data-dir=" + USER_DATA_DIR);
        logger.info("Adding user data directory argument...");
        options.addArguments("--profile-directory=" + PROFILE_DIRECTORY);
        logger.info("Adding profile directory argument...");

        this.driver = new ChromeDriver(options);
        logger.info("Creating ChromeDriver object...");
        this.actions = new Actions(driver);
        logger.info("Creating Actions object...");
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        logger.info("Creating WebDriverWait object...");
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
            driver.get(CHATGPT_URL);
            logger.info("Navigating to ChatGPT URL...");

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(200));
            logger.info("Creating WebDriverWait object...");

            By textareaLocator = By.id("prompt-textarea");
            logger.info("Creating textareaLocator object...");

            By submitButtonLocator = By.cssSelector("button[data-testid=send-button]");
            logger.info("Creating submitButtonLocator object...");

            for (Question question : questions) {
                String questionText = question.question;
                logger.info("Getting question text...");

                String marks = ". Answer the question as " + question.marks + " marks";
                logger.info("Creating marks string...");

                String additionalInfo = "";

                if (question.additionalInfo != null && !question.additionalInfo.isEmpty()) {
                    additionalInfo = " and add the following information: " + question.additionalInfo;
                    logger.info("Adding additional information...");
                }

                String prompt = questionText + marks + additionalInfo;
                logger.info("Creating prompt string...");

                // Entering question text into the text area
                driver.findElement(textareaLocator).sendKeys(prompt);
                logger.info("Entering question text into text area...");

                // Clicking the button to submit the question
                driver.findElement(submitButtonLocator).click();
                logger.info("Clicking submit button...");

                // Waiting for the button to disappear and then reappear
                wait.until(ExpectedConditions.invisibilityOfElementLocated(submitButtonLocator));
                logger.info("Waiting for submit button to disappear...");
                wait.until(ExpectedConditions.presenceOfElementLocated(submitButtonLocator));
                logger.info("Waiting for submit button to reappear...");
                Thread.sleep(5000);
                logger.info("Waiting for 5 seconds...");
            }
        } catch (Exception e) {
            logger.error("An error occurred while getting answers from chat: ", e);
        }

        logger.info("Getting answers from chat complete.");
    }

    @Test(priority = 2)
    public void generateChatToPDF() throws IOException {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript(
                "document.querySelector('#__next>div').classList.remove('h-full', 'overflow-hidden');" +
                "document.querySelector('#__next>div>div').classList.remove('overflow-hidden');" +
                "document.querySelector('#__next main').classList.remove('overflow-auto');" +
                "document.querySelector('#__next main')?.parentElement.classList.remove('overflow-hidden');" +
                "document.querySelector('#__next main>div>div').classList.remove('overflow-hidden');" +
                "document.querySelector('#__next main>div>div.w-full').classList.add('hidden');" +
                "document.querySelector('#__next header')?.classList.add('hidden');" +
                "document.body.classList.remove('dark');");

        Pdf pdf = ((PrintsPage) driver).print(new PrintOptions());
        Files.write(Paths.get(PDF_PATH + "answers.pdf"), OutputType.BYTES.convertFromBase64Png(pdf.getContent()));
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
