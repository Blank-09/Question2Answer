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
    private final String USER_DATA_DIR = "C:\\Users\\<username>\\AppData\\Local\\Google\\Chrome\\User Data";
    private final String PROFILE_DIRECTORY = "Profile 1";

    private final String QUESTION_SHEET_PATH = "./assets/sheets/questions.xlsx";
    private final String REPORT_PATH = "./out/reports/index.html";
    private final String LOGGER_PATH = "./out/logs/app.log";
    private final String SCREENSHOT_PATH = "./out/screenshots/";
    private final String PDF_PATH = "./out/result/";

    private WebDriver driver;
    private Actions actions;
    private ExtentReports reports;
    private Wait<WebDriver> wait;
    private List<Question> questions;

    private Logger logger = LogManager.getLogger(getClass());

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

        logger.info("Setting executable path to " + EXECUTABLE_PATH);
        options.setBinary(EXECUTABLE_PATH);

        logger.info("Setting user data directory to " + USER_DATA_DIR);
        options.addArguments("--user-data-dir=" + USER_DATA_DIR);

        logger.info("Setting profile directory to " + PROFILE_DIRECTORY);
        options.addArguments("--profile-directory=" + PROFILE_DIRECTORY);

        logger.info("Creating ChromeDriver object.");
        this.driver = new ChromeDriver(options);
        this.actions = new Actions(driver);
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

        logger.info("Initializing Testcase 1");
        logger.info("Preparing to initialize ChatGPT");

        try {
            logger.info("Navigating to " + CHATGPT_URL);
            driver.get(CHATGPT_URL);

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(200));

            By textareaLocator = By.id("prompt-textarea");
            By submitButtonLocator = By.cssSelector("button[data-testid=send-button]");

            for (Question question : questions) {
                logger.info("Querying question " + question.sno + " (" + question.marks + " marks)");

                String questionText = question.question;
                String marks = ". Answer the question as " + question.marks + " marks";
                String additionalInfo = "";

                if (question.additionalInfo != null && !question.additionalInfo.isEmpty()) {
                    logger.info("Adding additional information...");
                    additionalInfo = " and add the following information: " + question.additionalInfo;
                } else {
                    logger.warn("Additional Information not provided. Please provide it for better results.");
                }

                String prompt = questionText + marks + additionalInfo;
                logger.info("Prompt: " + prompt);

                // Entering question text into the text area
                logger.info("Entering question text into textarea...");
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
            logger.error("An error occurred while getting answers from Chat: ", e);
        }

        logger.info("ChatGPT Automation complete.");
        logger.info("Testcase 1 completed.");
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
        logger.info("Wrapping up...");
        logger.info("Quitting WebDriver");
        driver.quit();
        // reports.flush();
        logger.info("Wrap-up complete");
    }

    // All your private function goes here lexigraphically
    private void takeScreenshot(String name) throws IOException {
        logger.info("Taking Screenshot...");

        String screenshotPath = SCREENSHOT_PATH + name + "_" + timestamp + ".png";
        File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());

        FileUtils.copyFile(screenshotFile, new File(screenshotPath));
        logger.info("Screenshot saved at " + screenshotPath);
    }

}
