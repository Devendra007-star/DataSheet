package cares.cwds.salesforce.web.utilities;

import org.openqa.selenium.*;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.Assert;

import reportutilities.common.ReportCommon;
import reportutilities.common.ScreenshotCommon;
import reportutilities.extentmodel.ExtentUtilities;
import reportutilities.extentmodel.PageDetails;
import reportutilities.model.TestCaseParam;

import java.time.Duration;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.Set;


public class Webkeywords
{
	private static final String TITLE_ATTRIBUTE = "title";
	private static final String SUCCESSFULLY_ENTERED_TEXT ="Successfully Entered Text {} to {}";

	private static final String VERIFY_ELEMENT_TEXT = "Verify Element Text";
	private static final String VERIFY_ELEMENT_NOT_DISPLAYED = "Verify Element Not Displayed--";
	private static final String ENTERED_TEXT ="Entered Text -> ";

	private static final String VERIFY_ELEMENT_TITLE = "Verify Element Title";
	private static final String NAVIGATE = "Navigate -> ";
	private static final String CONTAINS_EXPECTED_TEXT = "Contains Expected Text:-";
	private static final String EXPECTED = ":Expected:-";
	private static final String FAILED_FORMAT = "Failed ==> {} {} ";
	private static final String ACTUAL_VALUE_DISPLAYED = "Actual Value Displayed-->";
	private static final String ACTUAL_TEXT_DISPLAYED = "Actual Text Displayed-->";
	private static final String LOG_FAILED_FORMAT = "Failed ==> {} {}";
	private static final String TESTDATA_NOT_APPLICABLE = "n\\\\a";

	private static final String CLICK_ACTION_PREFIX = "Click --> ";
	private static final String FIND_ACTION_PREFIX = "Find Element --> ";

	private static final String VALUE = "value";
	private static final String ERROR_LOG_FORMAT = "Failed ==>{} {} ";
	private static final String LOG_SWITCHED_TO = "Switched to: {}";

	private static final String EXPECTED_VALUE_DELIMITER = "<--> Expected Value-->";
	private static final int TIME_OUT = 20;



	private static Webkeywords instance = new Webkeywords();
	private static final Logger logger =LoggerFactory.getLogger(Webkeywords.class.getName());
	private static final String STATUSPASS="PASS";
	private static final String STATUSFAIL="FAIL";
	private static final String STATUSDONE="DONE";
	private static final String  VRIFYELEMENT= "Verify Element Displayed--";
	private static final String EMPTYTEXTMSG="Textbox is empty. Retrying...";
	private static final String TABLECOLUMNMSG= "Table Columns Text Displayed-->";
	ExtentUtilities extentUtilities = new ExtentUtilities();
	ScreenshotCommon scm = new ScreenshotCommon();

	ReportCommon testStepDetails = new ReportCommon();
	static ReportCommon exceptionDetails = new ReportCommon();
	
	public static Webkeywords instance()
	{
		return instance;

	}

	public void navigate(WebDriver driver, String url, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String actionNv = NAVIGATE+url;
		String actionDescriptionNv = "";
		LocalDateTime startTime=  LocalDateTime.now();
		WebDriverWait wait = new WebDriverWait(driver, TIME_OUT);

		try
		{

				driver.navigate().to(url);
				waitTillPageLoad(wait);
			testStepDetails.logTestStepDetails(driver, testCaseParam, actionNv, actionDescriptionNv,pageDetails, startTime, STATUSDONE);
			logger.info(SUCCESSFULLY_ENTERED_TEXT,url);

		}catch (InterruptedException e){
			logger.error("Failed == {} ", pageDetails.getPageActionDescription());
			testStepDetails.logExceptionDetails(driver, testCaseParam, pageDetails.getPageActionName(), pageDetails.getPageActionDescription(), startTime,e);
			Thread.currentThread().interrupt();
			throw new CustomException("Error Occured",e); 
        }
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT, actionNv, actionDescriptionNv);
			testStepDetails.logExceptionDetails(driver, testCaseParam, actionNv, actionDescriptionNv, startTime,e);
			testStepDetails.logTestStepDetails(driver, testCaseParam, actionNv, actionDescriptionNv,pageDetails, startTime, STATUSFAIL);

			 throw new CustomException(e.getMessage());
		}


	}

	public void openUrl(WebDriver drivernew, String url, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = "Navigate -->"+url ;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
		try
		{

			if (!(url.startsWith("http://") || url.startsWith("https://"))) {
				drivernew.navigate().to(url);
				waitTillPageLoad(wait);
			}
		}
		catch (Exception e)
		{
			logger.error(FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			Thread.currentThread().interrupt();
			throw new CustomException(e.getMessage());
		}
	}


	public boolean checkOptions(String options) {
		return options.equals("N//A") || options.equals("N/A") || options.equals("n//a") || options.equals("n/a");
	}

	public void trymultipleclick(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		Webkeywords.instance().fluentWait(drivernew, element);
		String action="";
		action = getaction(drivernew, element, testCaseParam, pageDetails);

		String actionDescription = element.toString();
		LocalDateTime startTime=  LocalDateTime.now();
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);


		try
		{
			Actions actionnew = new Actions(drivernew);
			actionnew.moveToElement(element).build().perform();
			element.click();
			waitTillPageLoad(wait);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);


		}
		catch (Exception e)
		{
			Thread.currentThread().interrupt();
			try
			{

				Actions actionnew = new Actions(drivernew);
				actionnew.moveToElement(element).build().perform();
				element.click();
				waitTillPageLoad(wait);
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);


			}

			catch(Exception ex)
			{
				Thread.currentThread().interrupt();
				try
				{

					Actions actionnew = new Actions(drivernew);
					actionnew.moveToElement(element).build().perform();
					element.click();
					waitTillPageLoad(wait);
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);


				}

				catch(Exception exx)
				{
					Thread.currentThread().interrupt();
					try
					{

						Actions actionnew = new Actions(drivernew);
						actionnew.moveToElement(element).build().perform();
						element.click();
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);



					}

					catch(Exception exxx)
					{
						try
						{
							Actions actionnew = new Actions(drivernew);
							actionnew.moveToElement(element).build().perform();
							element.click();
							testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);



						}

						catch(Exception exxxx)
						{
							handleActionExc(drivernew, element, testCaseParam, pageDetails, action, actionDescription,
									startTime);

						}

					}

				}

			}
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

			 throw new CustomException(e.getMessage());
		}
	}

	private void handleActionExc(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,
								 PageDetails pageDetails, String action, String actionDescription, LocalDateTime startTime) throws CustomException
	{
		try
		{
			Actions actionnew = new Actions(drivernew);
			actionnew.moveToElement(element).build().perform();
			element.click();
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);



		}

		catch(Exception exp)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,exp);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
			 throw new CustomException(exp.getMessage());


		}
	}

	private String getaction(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,
							 PageDetails pageDetails) throws CustomException {
		String action;
		try
		{
			action = CLICK_ACTION_PREFIX +getLocatorFromWebElement(element,drivernew,testCaseParam,pageDetails);

		}
		catch(Exception e)
		{
			action = CLICK_ACTION_PREFIX + getLocatorFromWebElement(element,drivernew,testCaseParam,pageDetails);
		}
		return action;
	}




	public void tryAction(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,
						   PageDetails pageDetails, String action, String actionDescription, LocalDateTime startTime) throws CustomException, InterruptedException {
			Actions actionnew = new Actions(drivernew);
			WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
			actionnew.moveToElement(element).build().perform();

			element.click();
			waitTillPageLoad(wait);

			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			Thread.sleep(900);

	}





		public void waitElementToBeVisibleNew(WebDriver drivernew, WebElement element, int timeOut) throws CustomException, InterruptedException {
		int count = 0;
		int maxTries = 5;
		while (count < maxTries) {
			try {
				By locatorvalue = getLocatorvalue(element);
				WebDriverWait wait = new WebDriverWait(drivernew, timeOut);
				wait.until(ExpectedConditions.presenceOfElementLocated(locatorvalue));
				break;
			} catch (Exception e) {
				Thread.sleep(10000);
				count++;
				if (++count == maxTries) {
					 throw new CustomException(e.getMessage());
				}
			}
		}
	}
		
	public void waitElementToBeVisible(WebDriver drivernew, WebElement element) {
					WebDriverWait wait = new WebDriverWait(drivernew, 20);
					wait.until(ExpectedConditions.visibilityOf(element));
	}
	
	public void waitElementToBeVisible(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException{
		String actionDescription = element.toString();
		LocalDateTime startTime = LocalDateTime.now();
		String action = FIND_ACTION_PREFIX + getLocatorFromWebElement(element, drivernew, testCaseParam, pageDetails);
		
		try {
			WebDriverWait wait = new WebDriverWait(drivernew, 20);
			wait.pollingEvery(Duration.ofMillis(500));
			wait.ignoring(StaleElementReferenceException.class);
			wait.until(ExpectedConditions.visibilityOf(element));
		}catch(Exception e) {
			logger.error(LOG_FAILED_FORMAT, action, actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,  strigToException(getActionDescription(action,actionDescription)));
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription, pageDetails, startTime, STATUSFAIL);
			throw new CustomException(e.getMessage());	
		}
	}

	public void waitForElementToBeVisible(WebDriver drivernew, String elementId) throws CustomException, InterruptedException  {
		int count = 0;
		int maxTries = 5;
		int timeOut=3000;
		while (count < maxTries) {
			try {
				WebElement element=drivernew.findElement(By.xpath(elementId));
				By locatorvalue = getLocatorvalue(element);
				WebDriverWait wait = new WebDriverWait(drivernew, timeOut);
				wait.until(ExpectedConditions.presenceOfElementLocated(locatorvalue));
				break;
			} catch (Exception e) {
				Thread.sleep(10000);
				count++;
				if (++count == maxTries) {
					 throw new CustomException(e.getMessage());
				}
			}
		}
	}


	public String getLocatorFromWebElement(WebElement element, WebDriver drivernew, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {
		   
	    final int MAX_ELEMENT_STRING_LENGTH = 10000;
	    final int MAX_SPLIT_PARTS = 2;

	    String elementString = element.toString();
	    if (elementString == null || elementString.isEmpty()) {
	        logger.warn("Received null or empty WebElement string");
	        return "";
	    }

	    if (elementString.length() > MAX_ELEMENT_STRING_LENGTH) {
	        logger.warn("WebElement string exceeds maximum allowed length");
	        throw new IllegalArgumentException("WebElement string exceeds maximum allowed length");
	    }

	    String[] parts = elementString.split("->", MAX_SPLIT_PARTS);
	    if (parts.length < MAX_SPLIT_PARTS) {
	        logger.warn("WebElement string does not contain expected delimiter");
	        return "";
	    }

	    String locatorOutput = parts[1];
	    int lastBracketIndex = locatorOutput.lastIndexOf(']');
	    if (lastBracketIndex != -1) {
	        locatorOutput = locatorOutput.substring(0, lastBracketIndex);
	    }

	    String[] splitLocator = locatorOutput.split(": ", MAX_SPLIT_PARTS);
	    if (splitLocator.length < MAX_SPLIT_PARTS) {
	        logger.warn("Locator string does not contain expected delimiter");
	        return "";
	    }

	    splitLocator[0] = splitLocator[0].trim();
	    splitLocator[1] = splitLocator[1].trim();

	    if (splitLocator[0].isEmpty() || splitLocator[1].isEmpty()) {
	        logger.warn("Invalid locator parts");
	        return "";
	    }

	    return getElementText(drivernew, splitLocator[0], splitLocator[1], testCaseParam, pageDetails);
	
	}

	public void switchtoFrame(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try
		{

			drivernew.switchTo().frame(element);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);

		}
		catch (Exception e)
		{
			logger.error(FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}
	}

	public void switchToWindowByTitle(WebDriver drivernew, String title, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException, InterruptedException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();

		Thread.sleep(2000);
		Set<String> windows=drivernew.getWindowHandles();


		for (String handle: windows)
		{

			String myTitle = drivernew.switchTo().window(handle).getTitle();
			if (myTitle.equals(title))
			{

				drivernew.switchTo().window(handle);
				logger.info("switched to Window--> {}",title);
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
				Thread.sleep(0);
				break;
			}
			else
			{
				switchToWindowHandle(drivernew, handle);
			}


		}

	}

	private void switchToWindowHandle(WebDriver drivernew, String handle) {
		try
		{
			drivernew.switchTo().window(handle);
			logger.info("switched to main Window");
		}
		catch (Exception e)
		{
			logger.info("Unable to switch to Main Window");
		}
	}


	public void refresh(WebDriver drivernew, TestCaseParam testCaseParam,PageDetails pageDetails)
	{
		String action = "Refresh the page ";
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
		try
		{
			drivernew.navigate().refresh();
			waitTillPageLoad(wait);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);

		}
		catch (Exception e)
		{
			Thread.currentThread().interrupt();
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);


		}
	}
	public void jsClick(WebDriver drivernew, WebElement element, String testData, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {
		// Use default value if timeOut is not provided
		if(!(testData.equalsIgnoreCase("n/a"))||!(testData.equalsIgnoreCase("n\\a"))||!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			String action = NAVIGATE;
			String actionDescription = "";
			LocalDateTime startTime = LocalDateTime.now();
			WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
	
			try {
				JavascriptExecutor js = (JavascriptExecutor) drivernew;
				js.executeScript("arguments[0].click();", element);
				waitTillPageLoad(wait);
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription, pageDetails, startTime, STATUSDONE);
			} catch (Exception e) {
				Thread.currentThread().interrupt();
				logger.error(FAILED_FORMAT, action, actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime, e);
				 throw new CustomException(e.getMessage());
			}
		}
	}



	public void zoombackToOriginal(WebDriver drivernew, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException,InterruptedException
	{
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try
		{


			JavascriptExecutor js = (JavascriptExecutor)drivernew;

			js.executeScript("document.body.style.zoom='100%'");
			waitTillPageLoad(wait);
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action ,actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			throw e;
		}
	}

	public void clickwithaction(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,PageDetails pageDetails,String testData) throws CustomException, InterruptedException
	{

		String action="";
		String actionDescription="";
		LocalDateTime startTime=  LocalDateTime.now();
		if(!(testData.equalsIgnoreCase("n/a"))||!(testData.equalsIgnoreCase("n\\a"))||!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{
			try {


				Webkeywords.instance().fluentWait(drivernew, element);
				Webkeywords.instance().waitElementforelementclickable(drivernew, element, 1000);
				scrollIntoViewElement(drivernew, element);
				action = CLICK_ACTION_PREFIX+getLocatorFromWebElement(element,drivernew,testCaseParam,pageDetails);
				actionDescription = CLICK_ACTION_PREFIX+getLocatorFromWebElement(element,drivernew,testCaseParam,pageDetails);

				tryAction(drivernew, element, testCaseParam, pageDetails, action, actionDescription, startTime);

			}
			catch(Exception e)
			{
				Thread.currentThread().interrupt();
				logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			}
		}
	}



	public void clickRadioButton(WebDriver driver, TestCaseParam testCaseParam, WebElement element) throws CustomException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try
		{

			for (int i = 0; i < 10; i++)
			{
				element.click();

				if (element.isSelected())
				{
					break;
				}


			}

		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action, actionDescription);
			testStepDetails.logExceptionDetails(driver, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}
	}




	private boolean checktextboxtestExp(String textBoxValue, String textBoxtest) {

		return textBoxValue.matches("\\d")||textBoxtest.matches("\\d");
	}

	private boolean checktextboxExpression(String textBoxValue, String textBoxtest) {
		return textBoxValue.matches("^[a-zA-Z0-9]+$")||textBoxtest.matches("^[a-zA-Z0-9]+$");
	}


	public boolean checkIfTrue(String text) {
		return text.equalsIgnoreCase("N/A")|| text.equalsIgnoreCase("N//A");
	}

	public void setText(WebDriver drivernew, WebElement element1, String text1,TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		if (!(text1.equalsIgnoreCase("n/a")) && !(text1.equalsIgnoreCase("n\\a")) && !(text1.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			String action = ENTERED_TEXT + text1;
			String actionDescription = ENTERED_TEXT+ text1;
			String password = "*".repeat(text1.length());
			LocalDateTime startTime=  LocalDateTime.now();
	
			try
			{
	
				if(!textCheck(text1))
				{
					Webkeywords.instance().fluentWait(drivernew, element1);
					WebDriverWait wait1 = new WebDriverWait(drivernew, 1000);
					wait1.until(ExpectedConditions.elementToBeClickable(element1));
					element1.clear();
					element1.sendKeys(text1);
					boolean isDataFilled1 = false;
					while (!isDataFilled1) {
	
	
						String textBoxValue=element1.getAttribute(VALUE);
						String textBoxtest=element1.getText();
						isDataFilled1 = validateTextBoxValue(drivernew, element1, text1, isDataFilled1, textBoxValue,
								textBoxtest);
					}
					if(element1.getAttribute("type").equals("password")) {
						logger.info("Successfully entered Password : {}", password);
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, "Entered Password -> "+password, "Entered Password -> "+password,pageDetails, startTime, STATUSDONE);
					}
					else {
						logger.info(SUCCESSFULLY_ENTERED_TEXT ,text1 , element1);
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, getActionDescription(action,actionDescription),pageDetails, startTime, STATUSDONE);
					}
	
					Thread.sleep(0);
				}
			}
			catch (Exception e)
			{
				Thread.currentThread().interrupt();
				logger.error(FAILED_FORMAT ,action ,actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime, strigToException(getActionDescription(action,actionDescription)));
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
				throw new CustomException(e.getMessage());
			}
		}
	}

	public void setDateText(WebDriver drivernew, WebElement element1, String text1, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = ENTERED_TEXT + text1;
		String actionDescription = ENTERED_TEXT+ text1;

		LocalDateTime startTime=  LocalDateTime.now();

		try
		{

			if(!textCheck(text1))
			{
				Webkeywords.instance().fluentWait(drivernew, element1);
				WebDriverWait wait1 = new WebDriverWait(drivernew, 1000);
				wait1.until(ExpectedConditions.elementToBeClickable(element1));
				element1.clear();
				element1.sendKeys(text1);
				logger.info(SUCCESSFULLY_ENTERED_TEXT ,text1 , element1);

				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
				Thread.sleep(0);
			}
		}
		catch (Exception e)
		{
			Thread.currentThread().interrupt();
			logger.error(FAILED_FORMAT ,action ,actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
			 throw new CustomException(e.getMessage());
		}

	}
	
	private boolean validateTextBoxValue(WebDriver drivernew, WebElement element, String text, boolean isDataFilled,
										 String textBoxValue, String textBoxtest) {
		if (checktextboxExpression(textBoxValue, textBoxtest))
		{
			isDataFilled = true;
		}

		else if (checkTestboxValue(textBoxValue, textBoxtest))
		{
			isDataFilled = true;
		}
		else if (textBoxValue.matches("^[a-z A-Z]+$")||textBoxtest.matches("^[a-z A-Z]+$"))
		{
			isDataFilled = true;
		}

		else if (checktextboxtestExp(textBoxValue, textBoxtest))
		{
			isDataFilled = true;
		}

		else if (checkTextBoxRegularExp(textBoxValue, textBoxtest)) {
			isDataFilled = true;
		}

		else if (textBoxValue.matches("^(?=.*[@#$!%^&+=-]).+$")||textBoxtest.matches("^(?=.*[@#$!%^&+=-]).+$")) {
			isDataFilled = true;
		}


		else
		{

			logger.info(EMPTYTEXTMSG);
			element.clear();


			JavascriptExecutor jsExecutor = (JavascriptExecutor) drivernew;
			jsExecutor.executeScript("arguments[0].value = arguments[1]", element, text);
		}
		return isDataFilled;
	}

	public void setTextwithoutverification(WebDriver drivernew, WebElement element, String text, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException, InterruptedException
	{
		String action = ENTERED_TEXT+ text;
		String actionDescription = ENTERED_TEXT+ text;

		LocalDateTime startTime=  LocalDateTime.now();
		
		if(!textCheck(text))
		{
			Webkeywords.instance().fluentWait(drivernew, element);
			WebDriverWait wait = new WebDriverWait(drivernew, 1000);
			wait.until(ExpectedConditions.elementToBeClickable(element));
			element.clear();
			element.sendKeys(text);

			boolean isDataFilled = false;
			while (!isDataFilled) {
				
				String textBoxValue=element.getAttribute(VALUE);
				String textBoxtest=element.getText();
				if (checktextboxExpression(textBoxValue, textBoxtest))
				{
					isDataFilled = true;
				}

				else if (checkTestboxValue(textBoxValue, textBoxtest))
				{
					isDataFilled = true;
				}

				else if (checktextboxtestExp(textBoxValue, textBoxtest))
				{
					isDataFilled = true;
				}
				else if (textBoxValue.matches("^(?=.*[a-zA-Z])(?=.*\\d)(?=.*[@#$%^&+=/_-]).+$") || textBoxtest.matches("^(?=.*[a-zA-Z])(?=.*\\d)(?=.*[@#$%^&+=/_-]).+$")) {

					isDataFilled = true;				        }

				else if (textBoxValue.matches("^(?=.*\\d)(?=.*[@#$%^&+=/_-]).+$") || textBoxtest.matches("^(?=.*\\d)(?=.*[@#$%^&+=/_-]).+$")) {
					isDataFilled = true;				        }

				else if (checkTextBoxRegularExp(textBoxValue, textBoxtest)) {
					isDataFilled = true;
				}



				else
				{

					logger.info(EMPTYTEXTMSG);
					element.clear();


					JavascriptExecutor jsExecutor = (JavascriptExecutor) drivernew;
					jsExecutor.executeScript("arguments[0].value = arguments[1]", element, text);
				}
			}





			logger.info(SUCCESSFULLY_ENTERED_TEXT," {} to {}" ,text ,element);

			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);

			Thread.sleep(0);
		}


	}

	private boolean checkTextBoxRegularExp(String textBoxValue, String textBoxtest) {
		return textBoxValue.matches("^(?=.*[a-zA-Z])(?=.*[@#$%^&+=/_-]).+$") || textBoxtest.matches("^(?=.*[a-zA-Z])(?=.*[@#$%^&+=/_-]).+$");

	}

	private boolean checkTestboxValue(String textBoxValue, String textBoxtest) {
		return textBoxValue.matches("^[a-zA-Z]+$")||textBoxtest.matches("^[a-zA-Z]+$");
	}

	private boolean textCheck(String text) {
		return text.equalsIgnoreCase("N/A")|| text.equalsIgnoreCase("N//A");
	}


	public void jsSetText(WebDriver drivernew, WebElement element, String text, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{

		LocalDateTime startTime=  LocalDateTime.now();
		String action= "";
		String actionDescription = "";

		try
		{
			waitwebElementVisible(element);

			JavascriptExecutor js = (JavascriptExecutor)drivernew;

			js.executeScript("arguments[0].setAttribute('value','" + text + "');", element);
		}


		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			 throw new CustomException(e.getMessage());
		}

	}

	public String jsGetText(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();

		String webText = "";

		try
		{
			webText = element.getText().trim();

		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			throw new CustomException(e.getMessage());


		}
		return webText;

	}

	public void submit(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException,InterruptedException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
		try
		{
			waitwebElementVisible(element);
			element.submit();
			waitTillPageLoad(wait);
			logger.info("Successfully Clicked Button ==> {}" ,element);
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			throw e;
		}

	}

	public void setDate(WebDriver drivernew, WebElement element, String text, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException,InterruptedException
	{
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try


		{
			String id = element.getAttribute("ID");
			((JavascriptExecutor)drivernew).executeScript("document.getElementById('" + id + "').removeAttribute('readonly',0);");


			element.click();
			waitTillPageLoad(wait);

			((JavascriptExecutor)drivernew).executeScript("document.getElementById('" + id + "').setAttribute('value', '" + text + "')");

			drivernew.findElement(By.xpath("//div[@id=\"fare_" + text + "\"]")).click();


		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			throw e;
		}

	}

	public void click(WebDriver drivernew, WebElement element, String testData, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {
		if (!(testData.equalsIgnoreCase("n/a")) && !(testData.equalsIgnoreCase("n\\a")) && !(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			String actionDescription = element.toString();
			LocalDateTime startTime = LocalDateTime.now();
			String action = CLICK_ACTION_PREFIX + getLocatorFromWebElement(element, drivernew, testCaseParam, pageDetails);
			WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);

			try {
				Actions actionnew = new Actions(drivernew);
				actionnew.moveToElement(element).build().perform();
				element.click();
				waitTillPageLoad(wait);
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription, pageDetails, startTime, STATUSDONE);
				Thread.sleep(900);

			} catch (InterruptedException e) {

				Thread.currentThread().interrupt();
			} catch (Exception e) {
				try {
					Webkeywords.instance().waitElementforelementclickable(drivernew, element, 1000);
					JavascriptExecutor executor = (JavascriptExecutor) drivernew;
					executor.executeScript("arguments[0].scrollIntoView(true);", element);
					Thread.sleep(300);
					executor.executeScript("arguments[0].click();", element);
					waitTillPageLoad(wait);
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription, pageDetails, startTime, STATUSDONE);

				} catch (InterruptedException f) {
					// Re-interrupt the thread to preserve the interrupted status
					Thread.currentThread().interrupt();
				} catch (Exception f) {
					logger.error(LOG_FAILED_FORMAT, action, actionDescription);
					testStepDetails.logExceptionDetails(drivernew, testCaseParam, action,actionDescription, startTime, strigToException(getActionDescription(action,actionDescription)));
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action,actionDescription, pageDetails, startTime, STATUSFAIL);
					throw new CustomException(f.getMessage());
				}
			}
		}
	}



	public void waitElementforelementclickable(WebDriver drivernew,WebElement element, int timeOut) throws  InterruptedException
	{

		try
		{
			By locatorvalue=getLocatorvalue(element);

			WebDriverWait wait = new WebDriverWait(drivernew, timeOut);
			wait.until(ExpectedConditions.elementToBeClickable(locatorvalue));

		}
		catch (WebDriverException e)
		{
			Thread.sleep(1500);
			By locatorvalue=getLocatorvalue(element);

			WebDriverWait wait = new WebDriverWait(drivernew, timeOut);
			wait.until(ExpectedConditions.elementToBeClickable(locatorvalue));


		}
	}



	public static void waitForPageLoad(WebDriver driver) {

		throw new UnsupportedOperationException("This method is not supported.");
	}

	public By getLocatorvalue(WebElement element) {

		try {
			String elementString = element.toString();
			int arrowIndex = elementString.indexOf("->");
			int bracketIndex = elementString.lastIndexOf(']');
			String locatorString = elementString.substring(arrowIndex + 2, bracketIndex).trim();
			int colonIndex = locatorString.indexOf(':');
			String locatorType = locatorString.substring(0, colonIndex).trim();
			String locatorValue = locatorString.substring(colonIndex + 1).trim();

			if (elementString.isEmpty() || elementString.length() > 1000 ||
					arrowIndex == -1 || bracketIndex == -1 || arrowIndex >= bracketIndex ||
					colonIndex == -1 || locatorType.isEmpty() || locatorValue.isEmpty()) {
				throw new IllegalArgumentException("WebElement is not displayed on the screen " + element);
			}
			return getlocatorvalueforwait(locatorType, locatorValue);

		}catch(IllegalArgumentException e) {
			throw new IllegalArgumentException(e);
		}
	}
	
	public By getlocatorvalueforwait(String locatorType,String locatorValue)
	{
		String action = NAVIGATE;
		String actionDescription = "";

		try
		{
			By locatorvalue = null;
			locatorType =locatorType.strip().trim();
			locatorValue =locatorValue.strip().trim();

			switch (locatorType.toLowerCase())
			{
				case "id":
					locatorvalue = By.id(locatorValue.trim());
					break;
				case "name":
					locatorvalue = By.name(locatorValue.trim());
					break;
				case "xpath":
					locatorvalue = By.xpath(locatorValue.trim());
					break;
				case "tag":
					locatorvalue = By.name(locatorValue.trim());
					break;
				case "link text":
					locatorvalue = By.linkText(locatorValue.trim());
					break;
				case "css":
					locatorvalue = By.cssSelector(locatorValue.trim());
					break;
				case "class":
					locatorvalue = By.className(locatorValue.trim());
					break;
				default:

			}
			return locatorvalue;
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action ,actionDescription);


			throw e;
		}

	}


	public void waitforinvisibilityofelement(WebDriver drivernew, By locator)
	{
		String action = NAVIGATE;
		String actionDescription = "";
		try
		{
			WebDriverWait wait = new WebDriverWait(drivernew, 3000);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(locator));
		}
		catch(Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action ,actionDescription);
			throw e;
		}

	}

	public String wpGetPageTitle(WebDriver drivernew)
	{
		String action = NAVIGATE;
		String actionDescription = "";
		try {
			String wpPageTitle="";
			int len=drivernew.findElements(By.xpath("//*[@alt='Print' and @tabindex='0' and @class='vertical-align-middle cursor-hand height-20px printDoc']/.././../div/h2/span/label")).size();
			if(len>0)
			{
				wpPageTitle=drivernew.findElement(By.xpath("//*[@alt='Print' and @tabindex='0' and @class='vertical-align-middle cursor-hand height-20px printDoc']/.././../div/h2/span/label")).getText();
			}
			else
			{
				wpPageTitle="";
			}
			return wpPageTitle;
		}
		catch(Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action ,actionDescription);

			throw e;
		}
	}




	public void fluentWait(WebDriver drivernew, WebElement element) {
		try {
			Wait<WebDriver> wait2 = new FluentWait<WebDriver>(drivernew)
					.ignoring(Exception.class)
					.ignoring(NoSuchElementException.class)
					.ignoring(StaleElementReferenceException.class)
					.ignoring(ElementNotVisibleException.class);

			wait2.until(driver -> {
				if (element.isDisplayed()) {
					return element;
				} else {
					return null;
				}
			});
		} catch (TimeoutException e) {
			logger.info(Arrays.toString(e.getStackTrace()));
		}
	}


	public void waitElementClickable(WebDriver drivernew, WebElement element)
	{
		String action = NAVIGATE;
		String actionDescription = "";
		try
		{
			logger.info("Waiting for Element = {}" ,element);
			WebDriverWait wait = new WebDriverWait(drivernew, 5000);
			wait.until(ExpectedConditions.elementToBeClickable(element));
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action ,actionDescription);

			throw e;
		}
	}


	public void waitElementClickable(WebDriver drivernew, WebElement locatorValue, int timeOut)
	{

		try
		{
			logger.info("Waiting for Element = {} ", locatorValue);
			WebDriverWait wait = new WebDriverWait(drivernew, timeOut);
			wait.until(ExpectedConditions.elementToBeClickable(locatorValue));

		}
		catch (TimeoutException e)
		{
			logger.error("Get {} , {} is not visible",e ,locatorValue);

		}
	}
	public void waitElementEnabled(WebElement element, int timeOut)
	{
		for (int i = 0; i < timeOut; i++)
		{

			if (element.isEnabled())
			{
				break;
			}




		}
	}

	public void switchFocusToOtherWindow(WebDriver drivernew, By pagetitlelocator)
	{
		String action = NAVIGATE;
		String actionDescription = "";
		try {

			Set<String> allWindows =  drivernew.getWindowHandles();
			for(String curWindow : allWindows){
				drivernew.switchTo().window(curWindow);
				if(!(drivernew.findElements(pagetitlelocator).isEmpty()))
				{
					String pageTitle=drivernew.findElement(pagetitlelocator).getText();
					logger.info(LOG_SWITCHED_TO,pageTitle);
				}
				else
				{
					String pageTitle=drivernew.getTitle();
					logger.info(LOG_SWITCHED_TO,pageTitle);

				}
			}

		}
		catch(Exception e)
		{
			logger.error(FAILED_FORMAT , action , actionDescription);
			throw e;
		}
	}

	public void switchFocusToMainWindow(WebDriver drivernew, By pagetitlelocator)
	{
		String action = NAVIGATE;
		String actionDescription = "";
		try {

			String pageTitle=drivernew.findElement(pagetitlelocator).getText();
			logger.info(LOG_SWITCHED_TO,pageTitle);
			Set<String> allWindows =  drivernew.getWindowHandles();

			for(String curWindow : allWindows){
				drivernew.switchTo().window(curWindow);
				if(!drivernew.findElements(pagetitlelocator).isEmpty())
				{
					pageTitle=drivernew.findElement(pagetitlelocator).getText();
					logger.info(LOG_SWITCHED_TO,pageTitle);
					break;
				}
				else
				{
					pageTitle=drivernew.getTitle();
					logger.info(LOG_SWITCHED_TO,pageTitle);

				}
			}
		}
		catch(Exception e)
		{
			logger.error(FAILED_FORMAT , action , actionDescription);
			throw e;
		}
	}

	public void waitTitleContains(WebDriver drivernew, String title)
	{
		String action = NAVIGATE;
		String actionDescription = "";
		try
		{
			int timeOut=10000;
			WebDriverWait wait = new WebDriverWait(drivernew, timeOut);
			wait.until(ExpectedConditions.titleContains(title));
		}
		catch (WebDriverException e)
		{
			logger.error(FAILED_FORMAT , action , actionDescription);
			throw e;
		}
	}

	public String getAttribute(WebElement element, String attribute)
	{
		return element.getAttribute(attribute);
	}

	public void verifyAttributeValue(WebDriver drivernew, WebElement element, String attribute,String expectedText,TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{

		String action = "Verify "+element.getAttribute(attribute)+" Attribute value-->"+expectedText;
		String actionDescription = "Verify "+element.getAttribute(attribute)+" Attribute value"+expectedText;
		LocalDateTime startTime = LocalDateTime.now();
		if(element.getAttribute(attribute).equalsIgnoreCase(expectedText))

		{
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
		}
		else
		{
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

		}
	}

	public boolean getTitle(WebDriver drivernew,String expectedTitle, TestCaseParam testCaseParam) throws CustomException
	{

		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try
		{

			String pageTitle = drivernew.getTitle();
			boolean validTitle = false;
			if (expectedTitle.contains(pageTitle))
			{
				validTitle = true;
			}
			else
			{
				validTitle = false;
			}

			logger.info(SUCCESSFULLY_ENTERED_TEXT ,pageTitle);
			return validTitle;


		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			throw new CustomException(e.getMessage());
		}
	}

	public String getCssValue( WebElement element, String value)
	{
		return element.getCssValue(value);

	}

	public String getPageSource(WebDriver drivernew)
	{
		return drivernew.getPageSource();
	}



	public void clearText(WebDriver drivernew, WebElement element, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try
		{

			element.clear();
		}
		catch (Exception e)
		{
			logger.error(FAILED_FORMAT ,action ,actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			 throw new CustomException(e.getMessage());
		}
	}

	public JavascriptExecutor javaScript(WebDriver driver)
	{
		return (JavascriptExecutor)driver;
	}

	public WebElement findElement(WebDriver drivernew, String value, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try
		{
			WebElement element = null;
			String locatorType = value.split(";")[0];
			String locatorValue = value.split(";")[1];

			switch (locatorType.toLowerCase())
			{
				case "id":
					element = drivernew.findElement(By.id(locatorValue));
					break;
				case "name":
					element = drivernew.findElement(By.name(locatorValue));
					break;
				case "xpath":
					element = drivernew.findElement(By.xpath(locatorValue));
					break;
				case "tag":
					element = drivernew.findElement(By.name(locatorValue));
					break;
				case "link":
					element = drivernew.findElement(By.linkText(locatorValue));
					break;
				case "css":
					element = drivernew.findElement(By.cssSelector(locatorValue));
					break;
				case "class":
					element = drivernew.findElement(By.className(locatorValue));
					break;
				default:

			}
			return element;
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action ,actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			 throw new CustomException(e.getMessage());
		}

	}

	public String getElementText(WebDriver drivernew, String locatorType,String locatorValue, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = NAVIGATE;
		String actionDescription = "";
		LocalDateTime startTime=  LocalDateTime.now();
		try
		{


			String text=null;
			WebElement element = null;
			locatorType =locatorType.trim();
			locatorValue =locatorValue.trim();

			switch (locatorType.toLowerCase())
			{
				case "id":
					element = drivernew.findElement(By.id(locatorValue));
					text=element.getAttribute(VALUE);
					text = checkTextNull(text, element);
					break;
				case "name":
				case "tag":
					element = drivernew.findElement(By.name(locatorValue));
					text=element.getAttribute(VALUE);
					text = checkTextNull(text, element);
					break;
				case "xpath":
					element = drivernew.findElement(By.xpath(locatorValue));
					text=element.getAttribute(VALUE);
					text = checkTextNull(text, element);
					break;

				case "link text":
					element = drivernew.findElement(By.linkText(locatorValue));
					text=element.getAttribute(VALUE);
					text = checkTextNull(text, element);
					break;
				case "css":
					element = drivernew.findElement(By.cssSelector(locatorValue));
					text=element.getAttribute(VALUE);
					text = checkTextNull(text, element);
					break;
				case "class":
					element = drivernew.findElement(By.className(locatorValue));
					text=element.getAttribute(VALUE);
					text = checkTextNull(text, element);
					break;
				default:
			}
			return text;
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			throw new CustomException(e.getMessage());
			
		}

	}

	private String checkTextNull(String text, WebElement element) {
		if(text==null)
		{
			text=element.getText();
		}
		return text;
	}

	public List<WebElement> findElements(WebDriver drivernew, String value)
	{
		List<WebElement> elements = null;
		String locatorType = value.split(";")[0];
		String locatorValue = value.split(";")[1];

		switch (locatorType.toLowerCase())
		{
			case "id":
				elements = drivernew.findElements(By.id(locatorValue));
				break;
			case "name":
				elements = drivernew.findElements(By.name(locatorValue));
				break;
			case "xpath":
				elements = drivernew.findElements(By.xpath(locatorValue));
				break;
			case "tag":
				elements = drivernew.findElements(By.name(locatorValue));
				break;
			case "link":
				elements = drivernew.findElements(By.linkText(locatorValue));
				break;
			case "css":
				elements = drivernew.findElements(By.cssSelector(locatorValue));
				break;
			case "class":
				elements = drivernew.findElements(By.className(locatorValue));
				break;
			default:

		}
		return  elements;
	}

	public boolean isElementVisible(WebElement element)
	{
		return element.isDisplayed();
	}

	public void waitwebElementVisible(WebElement locatorValue)
	{

		try
		{
			logger.info("Waiting for Element = {}" , locatorValue);

			boolean visible = isElementVisible(locatorValue);
			while (true)
			{
				if(visible)
				{
					break;
				}
				else
				{

					visible = false;
				}
			}

		}
		catch (Exception e)
		{
			logger.error("Error: {}, Locator value: {} is not visible.", e, locatorValue);
		}
	}




	public void verifyElementPresent(WebDriver driver, WebElement element,TestCaseParam testCaseParam) throws CustomException
	{
		String testStepName = "VerifyElementPresent";
		String testStepDescription = "VerifyElementPresent";
		LocalDateTime startTime = LocalDateTime.now();
		Boolean isDisplayed = false;

		try
		{

			waitwebElementVisible(element);
			isDisplayed = isElementVisible(element);
			logger.info("Element present in application: {}" , isDisplayed);
			if (isDisplayed.equals(true))
			{
				logger.info("The Element is displayed in the application");

			}
			else
			{
				logger.info("The Element is not displayed in the application");

			}
		}
		catch (Exception e)
		{
			logger.error("Failed ==> {}" , testStepDescription);
			exceptionDetails.logExceptionDetails(driver, testCaseParam, testStepName, testStepDescription, startTime);
			 throw new CustomException(e.getMessage());
		}
	}


	public void documentUpload(WebDriver drivernew, WebElement element,String path,String docName, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException,InterruptedException
	{
		String action = "Upload Document";
		String actionDescription = "Upload Document";
		LocalDateTime startTime = LocalDateTime.now();
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);

		try
		{
			Webkeywords.instance().fluentWait(drivernew, element);


			element.sendKeys(path+docName);
			waitTillPageLoad(wait);
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);


		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			throw e;
		}
	}


	//*********************************SCREEN VALIDATIONS**************************************************


	public void verifyDropdownText(WebDriver drivernew, WebElement element, String options, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		
		if (!(options.equalsIgnoreCase("n/a")) && !(options.equalsIgnoreCase("n\\a")) && !(options.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
		
			
		String action = "Verify Dropdown Text-->Actual:"+options+EXPECTED+element.getText();
		String actionDescription = "Verify Dropdown Text-->Actual:"+options+EXPECTED+element.getText();
		LocalDateTime startTime = LocalDateTime.now();
		Boolean valuefound = false;

		try
		{

				Webkeywords.instance().fluentWait(drivernew, element);
				Select select = new Select(element);
				List<WebElement> allOptions =select.getAllSelectedOptions();
				for(int i=0; i<allOptions.size();i++)
				{
					if(allOptions.get(i).getText().contains(options))
					{
						valuefound=true;
						action = "Verify Dropdown-->Actual:"+options+EXPECTED+allOptions.get(i).getText();
						actionDescription = "Verify Dropdown-->Actual:"+options+EXPECTED+allOptions.get(i).getText();
						break;
					}
				}

				if(valuefound.equals(true))
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}
			}
		
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}
	}
		
	}

	public void verifytextentered(WebDriver drivernew, WebElement element, String text, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException
	{
		if (!(text.equalsIgnoreCase("n/a")) && !(text.equalsIgnoreCase("n\\a")) && !(text.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			
			String action = "Verify Dropdown Selection-->"+text;
		String actionDescription = "Verify Dropdown Selection"+text;
		LocalDateTime startTime = LocalDateTime.now();
		Boolean verified = false;

		try
		{


				Webkeywords.instance().fluentWait(drivernew, element);
				String enteredText = element.getText();

				if(enteredText.equals(text))
				{
					verified=true;
				}


				if(verified.equals(true))
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSPASS);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
				}
			}
		catch (Exception e)
		{
			logger.error(FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}
	}
	}

	public void verifyElementDisplayed(WebDriver drivernew, WebElement element,String testdata, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
			if (!(testdata.equalsIgnoreCase("n/a")) && !(testdata.equalsIgnoreCase("n\\a")) && !(testdata.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
				
				String actionv1 = VRIFYELEMENT+element.getAttribute(VALUE);
				String actionDescription =VRIFYELEMENT+element.getAttribute(VALUE);
				LocalDateTime startTime = LocalDateTime.now();

			try
			{
				scrollIntoViewElement(drivernew, element);
				Actions action = new Actions(drivernew);
				action.moveToElement(element).perform();
				if(element.isDisplayed())
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, actionv1, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, actionv1, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT ,actionv1 , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, actionv1, actionDescription, startTime, strigToException(getActionDescription(actionv1,actionDescription)));
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyElementDisplayedtextattribute(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{


		if(!(testData.equalsIgnoreCase("n/a"))&&!(testData.equalsIgnoreCase("n\\a"))&&!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{
			String actionNew =VRIFYELEMENT+element.getText();
			String actionDescription =VRIFYELEMENT+element.getText();
			LocalDateTime startTime = LocalDateTime.now();

			try
			{
				Actions action = new Actions(drivernew);
				action.moveToElement(element).perform();
				if(element.isDisplayed())
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, actionNew, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, actionNew, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(FAILED_FORMAT , actionNew , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, actionNew, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyelementnotdisplayed(WebDriver drivernew, List<WebElement> element, String testData, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException
	{
		String action = VERIFY_ELEMENT_NOT_DISPLAYED+ testData;
		String actionDescription = VERIFY_ELEMENT_NOT_DISPLAYED+ testData;
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))&& !(testData.equalsIgnoreCase("n\\a"))&& !(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{


			try
			{

				if(!element.isEmpty())
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSPASS);
				}

			}
			catch (Exception e)
			{
				logger.error(ERROR_LOG_FORMAT, action, actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}


	public void verifyelementnotdisplayed(WebDriver drivernew, String xpath, String testData, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException
	{
		String action = VERIFY_ELEMENT_NOT_DISPLAYED+ testData;
		String actionDescription = VERIFY_ELEMENT_NOT_DISPLAYED+ testData;
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))&&!(testData.equalsIgnoreCase("n\\a"))&&!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{

			try
			{

				int size=drivernew.findElements(By.xpath(xpath)).size();
				if(size>0)
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSPASS);
				}

			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT , action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}




	public void verifyElementEnabled(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String actionnew = "Verify Element Enabled";
		String actionDescription = "Verify Element Enabled";
		LocalDateTime startTime = LocalDateTime.now();

		if (!testData.equalsIgnoreCase("n/a") && !testData.equalsIgnoreCase("n\\a") && !testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))

		{

			try
			{
				Actions action=new Actions(drivernew);
				action.moveToElement(element);

				if(element.isEnabled())
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, actionnew, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, actionnew, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(ERROR_LOG_FORMAT, actionnew , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, actionnew, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyElementSelected(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = "Verify Element Selected";
		String actionDescription = "Verify Element Selected";
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))&&!(testData.equalsIgnoreCase("n\\a"))&&!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{


			try
			{

				if(element.isSelected())
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT ,action ,actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyElementDisabled(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = "Verify Element Disabled";
		String actionDescription = "Verify Element Disabled";
		LocalDateTime startTime = LocalDateTime.now();
		if(!(testData.equalsIgnoreCase("n/a"))&&!(testData.equalsIgnoreCase("n\\a"))&&!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{


			try
			{


				if(!element.getAttribute("disabled").equals("true"))
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT , action ,actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyCheckBoxChecked(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = "Verify CheckBox Checked";
		String actionDescription = "Verify CheckBox Checked";
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))&&!(testData.equalsIgnoreCase("n\\a"))&&!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{

			try
			{

				if(element.getAttribute("checked").equals("true"))
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT , action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyElementTitle(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails)throws CustomException
	{
		String action = VERIFY_ELEMENT_TITLE;
		String actionDescription = VERIFY_ELEMENT_TITLE;
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))&&!(testData.equalsIgnoreCase("n\\a"))&&!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{


			try
			{

				if(element.getAttribute(TITLE_ATTRIBUTE).equals(testData))
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT ,action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyTextBoxValue(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails)throws CustomException
	{
		String action = "Verify TextBox Value";
		String actionDescription = VERIFY_ELEMENT_TITLE;
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))&&!(testData.equalsIgnoreCase("n\\a"))&&!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{

			try
			{
				Webkeywords.instance().fluentWait(drivernew, element);
				if(element.getAttribute(VALUE).equals(testData))
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, "Pass");
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}

			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT , action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public boolean verifyTextDisplayed(WebDriver drivernew, WebElement element,String testdata, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = "";
		String actionDescription = "";
		LocalDateTime startTime = LocalDateTime.now();

		boolean isresult=false;

		if(!(testdata.equalsIgnoreCase("n/a"))&&!(testdata.equalsIgnoreCase("n\\a"))&&!(testdata.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{
			try
			{
				if(!(testdata.equals("N//A")||testdata==null))
				{

					scrollIntoViewElement(drivernew, element);
					zoomWebPage(drivernew,"50%",testCaseParam,pageDetails);
					action=ACTUAL_TEXT_DISPLAYED +element.getText()+CONTAINS_EXPECTED_TEXT+testdata;
					actionDescription = ACTUAL_TEXT_DISPLAYED +element.getText()+CONTAINS_EXPECTED_TEXT+testdata;
					Webkeywords.instance().fluentWait(drivernew, element);
					if(testdata.contains("\""))
					{
						testdata=testdata.replace("\"", "");
					}
					if(element.getText().toLowerCase().equalsIgnoreCase(testdata.toLowerCase()))
					{
						testStepDetails.logVerificationDetails(drivernew, testCaseParam, action,  startTime, STATUSDONE, element.getText(), testdata);
						isresult=true;
					}
					else
					{
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

						isresult=false;
					}
				}
				zoomWebPage(drivernew,"100%",testCaseParam,pageDetails);
			}
			catch (Exception e)
			{
				logger.error(FAILED_FORMAT ,action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
		return isresult;
	}



	public boolean verifyText(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = "";
		String actionDescription = "";
		LocalDateTime startTime = LocalDateTime.now();


		try
		{

			if(!(testData.equals("N//A")||testData==null))
			{

				scrollIntoViewElement(drivernew, element);
				zoomWebPage(drivernew,"50%",testCaseParam,pageDetails);
				action=ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+testData;
				actionDescription = ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+testData;
				Webkeywords.instance().fluentWait(drivernew, element);
				if(testData.contains("\""))
				{
					testData=testData.replace("\"", "");
				}
				if(element.getText().toLowerCase().equalsIgnoreCase(testData.toLowerCase()))
				{
					testStepDetails.logVerificationDetails(drivernew, testCaseParam, action,  startTime, STATUSDONE, element.getText(), testData);

					return true;
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
					return false;
				}
			}
			zoomWebPage(drivernew,"100%",testCaseParam,pageDetails);
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}
		return false;
	}



	public void verifyTextDisplayedValueAttribute(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+testData;
		String actionDescription = ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+testData;
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))||!(testData.equalsIgnoreCase("n\\a"))||!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{


			try
			{

				if(!(testData.equals("N//A")||testData==null))
				{

					Webkeywords.instance().fluentWait(drivernew, element);
					if(testData.contains("\""))
					{
						testData=testData.replace("\"", "");
					}
					if(element.getAttribute(VALUE).toLowerCase().equalsIgnoreCase(testData))
					{
						testStepDetails.logVerificationDetails(drivernew, testCaseParam, action,  startTime, STATUSDONE, element.getText(), testData);
					}
					else
					{
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

					}
				}
			}
			catch (Exception e)
			{
				logger.error(FAILED_FORMAT ,action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyTextDisplayedTitleAttribute(WebDriver drivernew, WebElement element,String testData, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+testData;
		String actionDescription = ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+testData;
		LocalDateTime startTime = LocalDateTime.now();

		if(!(testData.equalsIgnoreCase("n/a"))||!(testData.equalsIgnoreCase("n\\a"))||!(testData.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE)))
		{

			try
			{


				if(!(testData.equals("N//A")||testData ==null))
				{

					Webkeywords.instance().fluentWait(drivernew, element);
					if(testData.contains("\""))
					{
						testData=testData.replace("\"", "");
					}
					if(element.getAttribute(TITLE_ATTRIBUTE).toLowerCase().equalsIgnoreCase(testData))
					{
						testStepDetails.logVerificationDetails(drivernew, testCaseParam, action,  startTime, STATUSDONE, element.getText(), testData);

					}
					else
					{
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

					}
				}
			}
			catch (Exception e)
			{
				logger.error(LOG_FAILED_FORMAT , action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}



	public void verifyTextDisplayedNew(WebDriver drivernew, WebElement element,List<String> text, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+text;
		String actionDescription = ACTUAL_TEXT_DISPLAYED+element.getText()+CONTAINS_EXPECTED_TEXT+text;
		LocalDateTime startTime = LocalDateTime.now();


		try
		{
			if(!(text.contains("N//A")||text==null))

			{
				Webkeywords.instance().fluentWait(drivernew, element);
				if(element.getText().contains(text.get(0)))
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}
			}
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}
	}

	public boolean verifyTableData(WebDriver drivernew, List<WebElement> element,String text, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = TABLECOLUMNMSG;
		String actionDescription = TABLECOLUMNMSG;
		LocalDateTime startTime = LocalDateTime.now();
		boolean verifydata=false;

		try
		{
			if(!(text.equals("N//A")||text==null))
			{


				List<WebElement> elementnew;
				elementnew=element;
				ArrayList<String> dataList = new ArrayList<>();

				String[] dataCount = new String[element.size()];
				int l = 0;
				for (WebElement data : elementnew)
				{
					action = ACTUAL_TEXT_DISPLAYED+data.getText()+"Expected:-"+text;
					actionDescription = ACTUAL_TEXT_DISPLAYED+data.getText()+"Expected:-"+text;
					dataCount[l++] = data.getText();
					dataList.add(data.getText());

					if(data.getText().contains(text))
					{
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action,actionDescription,pageDetails, startTime, STATUSPASS);
						verifydata= true;
					}
					else
					{
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action,actionDescription,pageDetails, startTime, STATUSFAIL);
						verifydata= false;
					}

				}
			}


			return 	verifydata;


		}
		catch (Exception e)
		{
			logger.error(ERROR_LOG_FORMAT, action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			return false;
		}
	}

	public void verifyTableDataifNull(WebDriver drivernew, List<WebElement> element, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = TABLECOLUMNMSG;
		String actionDescription = TABLECOLUMNMSG;
		LocalDateTime startTime = LocalDateTime.now();


		try
		{

			ArrayList<String> dataList = new ArrayList<>();

			String[] dataCount = new String[element.size()];
			int l = 0;
			for (WebElement data : element)
			{
				action = TABLECOLUMNMSG+data.getText();
				actionDescription = TABLECOLUMNMSG+data.getText();
				dataCount[l++] = data.getText();
				dataList.add(data.getText());
				if(data.getText()!=null)
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}
			}
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}
	}

	public void verifyPartialTextDisplayed(WebDriver drivernew, WebElement element,String text, TestCaseParam testCaseParam,PageDetails pageDetails) throws CustomException
	{
		String action = "Verify Element Partial Text";
		String actionDescription = "Verify Element Partial Text";
		LocalDateTime startTime = LocalDateTime.now();


		try
		{

			if(element.getText().toLowerCase().contains(text.toLowerCase()))
			{
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			}
			else
			{
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
			}

		}
		catch (Exception e)
		{
			logger.error(FAILED_FORMAT , action ,actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());

		}
	}




	public void verifyValueEntered(WebDriver drivernew, WebElement element,String text, TestCaseParam testCaseParam,PageDetails pageDetails)
	{
		String action = VERIFY_ELEMENT_TEXT;
		String actionDescription = VERIFY_ELEMENT_TEXT;
		LocalDateTime startTime = LocalDateTime.now();


		try
		{
			if(!(text.equalsIgnoreCase("N//A")||text.equalsIgnoreCase("N/A")))
			{


				String actualText=element.getAttribute(VALUE);
				if(text.contains("\""))
				{
					text=text.replace("\"", "");
				}
				if(actualText.equals("___-__-____"))
				{
					actualText=actualText.replace("___-__-____", "");
				}
				if(actualText.contains("_-"))
				{
					actualText=actualText.replace("_", "");
					actualText=actualText.replace("-", "");
				}


				action = "Actual Value Displayed" + EXPECTED_VALUE_DELIMITER + actualText + EXPECTED_VALUE_DELIMITER + text;

				actionDescription = "Actual Value Displayed" + EXPECTED_VALUE_DELIMITER + actualText + EXPECTED_VALUE_DELIMITER + text;

				if(actualText.equalsIgnoreCase(text))
				{

					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
				}
			}
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action , actionDescription);


		}
	}

	public void verifyValueSelected(WebDriver drivernew, WebElement element,String options, TestCaseParam testCaseParam,PageDetails pageDetails)
	{
		String action = VERIFY_ELEMENT_TEXT;
		String actionDescription = VERIFY_ELEMENT_TEXT;
		LocalDateTime startTime = LocalDateTime.now();


		try
		{
			if(!isOptionsTrue(options))
			{
				Select select = new Select(element);
				element=select.getFirstSelectedOption();

				String actualText=element.getText();


				action = ACTUAL_VALUE_DISPLAYED+actualText+EXPECTED_VALUE_DELIMITER+options;
				actionDescription = ACTUAL_VALUE_DISPLAYED+actualText+"<-->Expected Value-->"+options;

				if(actualText.equals(options))
				{

					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
				}
			}

		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT, action, actionDescription);

		}
	}




	public String isReadyEmpty(String isReadOnly) {
		if(isReadOnly==null)
		{
			isReadOnly="";
		}
		return isReadOnly;
	}

	private boolean isDisableBlank(String isDisabled) {
		return isDisabled.equals("")||isDisabled==null;
	}

	public boolean isDisableNull(String isDisabled, String isReadOnly) {
		return isReadOnly.equals("true")&&isDisableBlank(isDisabled);
	}

	public boolean isReadblnkisDisabledtrue(String isDisabled, String isReadOnly) {
		return isReadOnly.equals("") && isDisabled.equals("true");
	}

	public boolean isReadDisableTrue(String isDisabled, String isReadOnly) {
		return isDisabled.equals("true") && isReadOnly.equals("true");
	}

	public void verifypropertyofelement(WebDriver drivernew, WebElement element, String dataValue, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException
	{
		String action = "Verify Property of Element-->"+element.getAttribute(VALUE);
		String actionDescription = "Verify Property of Element-->"+element.getAttribute(VALUE);
		LocalDateTime startTime = LocalDateTime.now();
		if(pageDetails != null)
		{


			try
			{
				Boolean displayed = true;
				displayed = element.isDisplayed();
				if (displayed.equals(true))
				{
					String prop = "";
					String[] data = null;

					data = dataValue.split("&");
					prop=element.getAttribute(data[0]) ;

					if (prop.equals(data[1]))
					{
						action = "The " + data[0] + " Property of the element is as expected";
						actionDescription = "The " + data[0] + " Property of the element is as expected";
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);

					}
					else
					{
						action = "The " + data[0] + " Property of the element is not as expected";
						actionDescription = "The " + data[0] + " Property of the element is not as expected";
						testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

					}
				}
			}
			catch (Exception e)
			{
				logger.error(FAILED_FORMAT , action , actionDescription);
				testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
				 throw new CustomException(e.getMessage());
			}
		}
	}

	public void verifyCheckBoxIsNotChecked(WebDriver drivernew,WebElement element,TestCaseParam testCaseParam, PageDetails pageDetails ) throws CustomException
	{
		String action="";
		String actionDescription="";
		LocalDateTime startTime = LocalDateTime.now();
		try
		{

			if (!element.isSelected())
			{

				action = "The element is Not checked";
				actionDescription = "The element is Not checked";
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);

			}
			else
			{
				action = "The element is  checked";
				actionDescription = "The element is  checked";
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

			}
		}
		catch (Exception e)
		{

			logger.error(ERROR_LOG_FORMAT, action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
			 throw new CustomException(e.getMessage());
		}

	}



	public void enterTextByFunctionKeys(WebDriver drivernew,WebElement element,Keys key, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException
	{


		String action="";
		String actionDescription="";
		LocalDateTime startTime = LocalDateTime.now();


		try
		{
			Webkeywords.instance().fluentWait(drivernew, element);
			WebDriverWait wait = new WebDriverWait(drivernew, 1000);
			wait.until(ExpectedConditions.elementToBeClickable(element));

			element.sendKeys(key);

			logger.info(SUCCESSFULLY_ENTERED_TEXT ," {} to {}",key ,element);
		}
		catch (Exception e)
		{
			action = "Failed to Dismiss Alert Message";
			actionDescription = "Failed to Dismiss Alert Message";
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
			 throw new CustomException(e.getMessage());
		}


	}

	public void zoomWebPage(WebDriver drivernew,String zoomValue, TestCaseParam testCaseParam, PageDetails pageDetails)
	{


		String action="";
		String actionDescription="";
		LocalDateTime startTime = LocalDateTime.now();



		try
		{

			JavascriptExecutor executor = (JavascriptExecutor)drivernew;
			executor.executeScript("document.body.style.zoom = '"+zoomValue+"'");
			action = "Sucessfully Zoomed the page to -->"+zoomValue;
			actionDescription = "Sucessfully Zoomed the page to -->"+zoomValue;
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);


		}
		catch (Exception e)
		{
			action = "Failed to Zoom the page";
			actionDescription = "Failed to Zoom the page";
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);

		}


	}

	public void mouseHover(WebDriver drivernew,WebElement element, TestCaseParam testCaseParam, PageDetails pageDetails)
	{


		String actionnew="Mouse Hover";
		String actionDescription="Mouse Hover";
		LocalDateTime startTime = LocalDateTime.now();



		try
		{
			Webkeywords.instance().fluentWait(drivernew, element);
			Actions action = new Actions(drivernew);

			action.moveToElement(element).build().perform();
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, actionnew, actionDescription,pageDetails, startTime, STATUSDONE);


		}
		catch (Exception e)
		{
			actionnew = "Failed do mouse hover";
			actionDescription = "Failed do mouse hover";
			logger.error(LOG_FAILED_FORMAT , actionnew , actionDescription);
		}


	}



	public void countZeroValidation(WebDriver drivernew, TestCaseParam testCaseParam, PageDetails pageDetails,
									 String action, String actionDescription, LocalDateTime startTime, int count) throws CustomException
	{
		if(count==0)
		{
			testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

		}
	}

	private boolean isOptionsTrue(String options) {
		return options.equals("N//A") || options.equals("N/A") || options.equals("n//a") || options.equals("n/a")||options==null;
	}

	public void scrollUpPageToTheTop(WebDriver driver)
	{

		String action = "Scroll Up Page To The Top";
		String actionDescription = "Scroll Up Page To The Top";



		try
		{

			((JavascriptExecutor) driver).executeScript("window.scrollTo(document.body.scrollHeight, 0)");


		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT,action,actionDescription);

		}



	}

	public void scrollIntoViewElement(WebDriver driver, WebElement element)
	{

		Webkeywords.instance().fluentWait(driver, element);

		JavascriptExecutor js = (JavascriptExecutor) driver;

		js.executeScript("arguments[0].scrollIntoView(true);",element);

	}



	public void verifydropdownoptionnotavailable(WebDriver drivernew, WebElement element, String options, TestCaseParam testCaseParam, PageDetails pageDetails)throws CustomException
	{
		String action = "Verifying Dropdown Options -->"+element.getText()+" : does not contain :-"+options;
		String actionDescription = "Verifying Dropdown Options -->"+element.getText()+" : does not contain :-"+options;
		LocalDateTime startTime = LocalDateTime.now();
		Boolean valuefound = true;

		try
		{
			if (!("N//A".equals(options) || options == null))
			{


				Webkeywords.instance().fluentWait(drivernew, element);
				Select select = new Select(element);
				List<WebElement> allOptions = select.getOptions();
				for(int i=0; i<allOptions.size(); i++)
				{

					if(allOptions.get(i).getText().contains(options))
					{
						valuefound=false;
						action = "Verify Dropdown-->contains "+allOptions.get(i).getText()+"which is not as Expected";
						actionDescription = "Verify Dropdown-->contains "+allOptions.get(i).getText()+"which is not as Expected";

						break;
					}
				}
				if(valuefound.equals(true))
				{
					testStepDetails.logVerificationDetails(drivernew, testCaseParam, action,  startTime, STATUSDONE, element.getText(), options);
				}
				else
				{
					testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);

				}
			}
		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT , action , actionDescription);
			testStepDetails.logExceptionDetails(drivernew, testCaseParam, action, actionDescription, startTime,e);
		
		}
	}





	public void verifyDropdownValueSelected(WebDriver drivernew, WebElement element,String options, TestCaseParam testCaseParam,PageDetails pageDetails)
	{
		String action = VERIFY_ELEMENT_TEXT;
		String actionDescription = VERIFY_ELEMENT_TEXT;
		LocalDateTime startTime = LocalDateTime.now();


		try
		{
			String actualText=element.getAttribute(TITLE_ATTRIBUTE);


			action = ACTUAL_VALUE_DISPLAYED+actualText+EXPECTED_VALUE_DELIMITER+options;
			actionDescription = ACTUAL_VALUE_DISPLAYED+actualText+"<-->Expected Value-->"+options;

			if(actualText.equals(options))
			{

				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSDONE);
			}
			else
			{
				testStepDetails.logTestStepDetails(drivernew, testCaseParam, action, actionDescription,pageDetails, startTime, STATUSFAIL);
			}



		}
		catch (Exception e)
		{
			logger.error(LOG_FAILED_FORMAT ,action,actionDescription);


		}
	}

	public static final void waitTillPageLoad(WebDriverWait wait) throws InterruptedException {
		ExpectedCondition<Boolean> pageLoadCondition = driver -> ((JavascriptExecutor) driver)
				.executeScript("return document.readyState").equals("complete");
		wait.until(pageLoadCondition);
		Thread.sleep(2000);
}

	public void selectDropdownValueByElement(WebDriver driver, WebElement element, List<WebElement> listElement,String listValues, String elementName, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {
		if (!(listValues.equalsIgnoreCase("n/a")) && !(listValues.equalsIgnoreCase("n\\a")) && !(listValues.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			GenericLocators genericLocators =  new GenericLocators(driver);

			logger.info("user is selecting {} values in {} dropdown" , listValues, elementName);

			String[] valuesToSelect;
			if (listValues.contains(";")) {
				valuesToSelect = listValues.split(";");
			} else {
				valuesToSelect = new String[] { listValues };
			}
			Webkeywords.instance().click(driver,element, "", testCaseParam, pageDetails);
			for (String value : valuesToSelect) {
				Webkeywords.instance().waitTillElementsSizeIsMoreThanZero(listElement);

				if (listElement.isEmpty()) {
					Webkeywords.instance().click(driver,element, "", testCaseParam, pageDetails);
					listElement = genericLocators.dropdownList(driver,elementName,listValues);
				}
				boolean flag = false;
				flag = extracted(driver, listElement, testCaseParam, pageDetails, value, flag);

				if (!flag) { 
					selectIndexPositionFromInputDropdown(driver,listValues, elementName, testCaseParam, pageDetails);
					break;
				}

			}
		}
	}

	private boolean extracted(WebDriver driver, List<WebElement> listElement, TestCaseParam testCaseParam,
			PageDetails pageDetails, String value, boolean flag) throws CustomException {
		for (WebElement listValue : listElement) {
			if (listValue.getText().equalsIgnoreCase(value)) {
				Webkeywords.instance().click(driver,listValue, "", testCaseParam, pageDetails);
				flag = true;
				break;
			}
		}
		return flag;
	}
	
	public void selectValueInputDropdown(WebDriver driver, String listValues, String elementName, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {

		if (!(listValues.equalsIgnoreCase("n/a")) && !(listValues.equalsIgnoreCase("n\\a")) && !(listValues.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			GenericLocators genericLocators =  new GenericLocators(driver);
	
			logger.info("user is selecting {} values in {} dropdown", listValues,elementName);

			String[] valuesToSelect;
			if (listValues.contains(";")) {
				valuesToSelect = listValues.split(";");
			} else {
				valuesToSelect = new String[] { listValues };
			}
			WebElement element = genericLocators.dropdown(driver, elementName, listValues);
			Webkeywords.instance().click(driver,element, "", testCaseParam, pageDetails);
			for (String value : valuesToSelect) {
				this.waitTillElementsSizeIsMoreThanZero(genericLocators.dropdownList(driver,elementName,listValues));
				List<WebElement> listElement = genericLocators.dropdownList(driver,elementName, listValues);
	
				if (listElement == null || listElement.isEmpty()) {
					this.click(driver,element, "", testCaseParam, pageDetails);
					listElement = genericLocators.dropdownList(driver,elementName,listValues);
				}
				boolean flag = false;
				flag = extracted(driver, listElement, testCaseParam, pageDetails, value, flag);
	
				if (!flag) { 
					selectIndexPositionFromInputDropdown(driver,listValues, elementName, testCaseParam, pageDetails);
					break;
				}
	
			}
		}
	}


	public void selectIndexPositionFromInputDropdown(WebDriver driver, String valueToSelect, String labelName,TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {
		if (!(valueToSelect.equalsIgnoreCase("n/a")) && !(valueToSelect.equalsIgnoreCase("n\\a")) && !(valueToSelect.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			WebElement inputDropdown = null;
			List<WebElement> inputDropdowns = driver.findElement(By.xpath("//button[@id=string(//label[text()='" + labelName + "']/@for)]"
							+ "| //input[@id=string(//label[text()='" + labelName + "']/@for)]"
							+ "| //input[@id=string(//span[text()='" + labelName + "']/../@for)]"
							+ "| //input[@id=string((//span[text()='" + labelName + "']/../@for)[2])]"));


			if (Objects.isNull(inputDropdowns)) {
				throw new NullPointerException("Exception occurred. Element name:"
						+ labelName + " dropdown web element not found");
			} 

			this.waitTillElementsSizeIsMoreThanZero(inputDropdowns);
			for (WebElement ele : inputDropdowns) {
				if (ele.isDisplayed()) {
					inputDropdown = ele;
					break;
				}
			}	
			if (Objects.isNull(inputDropdown)) {
				throw new NullPointerException("Exception occurred. Element"
						+ inputDropdown + " dropdown web element not found");
			} 

			selectValueIndexPosition(driver, valueToSelect, inputDropdown, testCaseParam, pageDetails);
		}
	}


	public void selectValueIndexPosition(WebDriver driver, String valueToSelect,WebElement inputDropdown, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {
		if (!(valueToSelect.equalsIgnoreCase("n/a")) && !(valueToSelect.equalsIgnoreCase("n\\a")) && !(valueToSelect.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			this.waitElementClickable(driver, inputDropdown);
			int index=0; 
			extracted(driver, valueToSelect, inputDropdown, testCaseParam, pageDetails, index);
		}
	}

	private void extracted(WebDriver driver, String valueToSelect, WebElement inputDropdown,
			TestCaseParam testCaseParam, PageDetails pageDetails, int index) throws CustomException {
		for (int indexLoop = 0; indexLoop < 100; indexLoop++) {
			if(indexLoop==0) {
				this.click(driver, inputDropdown, "", testCaseParam, pageDetails);
				inputDropdown.sendKeys(Keys.HOME);
				inputDropdown.sendKeys(Keys.RETURN);	
			}
			this.click(driver, inputDropdown, "", testCaseParam, pageDetails);
			inputDropdown.sendKeys(Keys.DOWN);
			inputDropdown.sendKeys(Keys.RETURN);
			String selectedText = inputDropdown.getAttribute("data-value");
			if (selectedText.equalsIgnoreCase(valueToSelect)) {
				break;
			}
			if(String.valueOf(index).equals(valueToSelect)) {
				break;
			}else {
				index++;
			}
			try {
				Thread.sleep(2000);
			} catch (InterruptedException e) {
				logger.error(e.getMessage());
				Thread.currentThread().interrupt();
			}
		}
	}
	
	/**
	 * This method wait till the size of the list is more than 0
	 * 
	 * @param wait
	 * @param elements
	 */
	public final void waitTillElementsSizeIsMoreThanZero(List<WebElement> elements) {
		try {
			int counter = 0;
			while (elements.isEmpty()) {
				Thread.sleep(1000);
				if (counter > 30) {
					break;
				}
				counter++;
			}
		} catch (InterruptedException ex) {
			logger.info("Interrupted Exception Occured");
			ex.getClass().getSimpleName();
			Thread.currentThread().interrupt();
		} catch (NullPointerException ne) {
			logger.error("Nullpointer exception occurred");
			ne.getMessage();
		}
	}

	/**
	 * This method closes all open tabs
	 * 
	 * @param wait
	 * @param elements
	 * @throws InterruptedException 
	 */
	public final void closeAllOpenTabs(WebDriver drivernew, WebElement closeButton){
		WebDriverWait wait = new WebDriverWait(drivernew, TIME_OUT);
		try {
			waitTillPageLoad(wait);
			while(closeButton.isDisplayed()) {
				closeButton.click();
				waitTillPageLoad(wait);
			}
		}
		catch(Exception e){
			Thread.currentThread().interrupt();
			logger.info("No close button : ");
		}
	}
	
	public static final void waitTillElementIsVisibleLocator(WebDriverWait wait, By by) {
		wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(by));
	}
	
	public WebElement waitAndReturnWebElement(WebDriver driver,By bylocator) {
		WebDriverWait wait = new WebDriverWait(driver, TIME_OUT);
		try {
			waitTillElementIsVisibleLocator(wait, bylocator);
		} catch (TimeoutException e) {
			wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(bylocator));
		}
		return driver.findElement(bylocator);
	}
	
	public String getActionDescription(String action,String actionDesc) {
		String [] res;
		String desc;
		if(action.equalsIgnoreCase(CLICK_ACTION_PREFIX)) {
			res = actionDesc.split("Proxy element for:");
			desc = res[1].trim();
			return "Can not click on the object " + desc + " is not displayed on the current screen";
		}else if (action.equalsIgnoreCase(ENTERED_TEXT)) {
			desc = actionDesc;
			return "Can not find the object " + desc + " is not displayed on the current screen";
		}else if(action.contains(VRIFYELEMENT)) {
			res = actionDesc.split("Verify Element Displayed--");
			desc = res[1].trim();
			return "Can not verify the object " + desc + " is not displayed on the current screen";
		}else if(action.equalsIgnoreCase(FIND_ACTION_PREFIX)) {
			res = actionDesc.split("Proxy element for:");
			desc = res[1].trim();
			return "Can not find the object " + desc + " is not displayed on the current screen";
		}
		else {
			return actionDesc;
		}
	}
	
	
	public  void selectValueInSearchbox(WebDriver driver,String value, String elementName, String testdata) throws NoSuchElementException, CustomException {

		Actions actions = new Actions(driver);
		JavascriptExecutor executor = (JavascriptExecutor) driver;

		if(value == null) {
			throw new CustomException("Value is null" + value);
		}

		waitTillElementsSizeIsMoreThanZero(GenericLocators.searchboxList(driver,elementName,testdata));

		List<WebElement> parentlistElement = GenericLocators.searchboxList(driver, elementName,testdata);
		if (parentlistElement != null) {
			for (WebElement dropdownValue : parentlistElement) {
				if (dropdownValue.getText().trim().equalsIgnoreCase(value.trim())) {
					try {
						actions.moveToElement(dropdownValue).perform();
						dropdownValue.click();
					} catch (Exception e) {
						executor.executeScript("arguments[0].click()", dropdownValue);
					}
					break;
				}
			}
		}
		else {
			clickAction(driver,waitAndReturnWebElement(driver,
					By.xpath("" + "//*[text()='Show All Results for \"" + value + "\"']")), false);

			clickAction(driver,
					waitAndReturnWebElement(driver,
							By.xpath("//th[@title='Full Name']//ancestor::table//tr//a[text()='" + value + "']")),
					false);
		}
	}
	

	
	public static void clickAction(WebDriver driver, WebElement element, boolean useAction) {
		Actions actions = new Actions(driver);
		if (Objects.isNull(element)) {
			throw new NullPointerException(
					"Exception occurred while clicking element: " + element + " web element not found");
		}

		try {
			if (element.isEnabled()) {
				element.click();
			}

		} catch (Exception e) {
			try {
				if (useAction) {
					actions.moveToElement(element).perform();
					element.click();
				} else {
					throw new JavascriptException("Action class not used.");
				}
			} catch (Exception exp) {
				JavascriptExecutor executor = (JavascriptExecutor) driver;
				executor.executeScript("arguments[0].click()", element);
			}
		}

	}
	
	public Exception strigToException(String message) {
		return new Exception(message);
	}
	
	
	
	public void scrollTillElement(WebDriver driver,WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("arguments[0].scrollIntoView();", element);
	}
	
	public void jsClick(WebDriver driver,WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("arguments[0].click();", element);
	}
	
	public void verifyDropdownPickList(WebDriver driver, String listValues, String elementName, TestCaseParam testCaseParam, PageDetails pageDetails) throws CustomException {

		if (!(listValues.equalsIgnoreCase("n/a")) && !(listValues.equalsIgnoreCase("n\\a")) && !(listValues.equalsIgnoreCase(TESTDATA_NOT_APPLICABLE))) {
			GenericLocators genericLocators =  new GenericLocators(driver);
	
			logger.info("user is selecting {} values in {} dropdown", listValues,elementName);

			String[] valuesToSelect;
			
			if (listValues.contains(",")) {
				valuesToSelect = listValues.split(",");
			} else {
				valuesToSelect = new String[] { listValues };
			}
			WebElement element = genericLocators.dropdown(driver, elementName, listValues);
			Webkeywords.instance().click(driver,element, "", testCaseParam, pageDetails);
			for (String value : valuesToSelect) {
				this.waitTillElementsSizeIsMoreThanZero(genericLocators.dropdownList(driver,elementName,listValues));
				List<WebElement> listElement = genericLocators.dropdownList(driver,elementName, listValues);
	
				if (listElement == null || listElement.isEmpty()) {
					this.click(driver,element, "", testCaseParam, pageDetails);
					listElement = genericLocators.dropdownList(driver,elementName,listValues);
				}			
				extracted(value, listElement);
				logger.info("DropDown list is displayed");
			}
		}
	}

	private void extracted(String value, List<WebElement> listElement) {
		for (WebElement listValue : listElement) {
			if (listValue.getText().equalsIgnoreCase(value)) {
				Assert.assertTrue(true);
			}
		}
	}
	
}