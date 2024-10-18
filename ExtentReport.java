package reportutilities.extentreport;


import reportutilities.extentmodel.TestCaseDetails;
import reportutilities.extentmodel.TestStepDetails;
import reportutilities.model.InterfaceTCDetails;
import reportutilities.model.InterfaceTestRun;
import reportutilities.model.TestCaseParam;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.CodeLanguage;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.model.ExceptionInfo;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExtentReport
{

	private static ExtentHtmlReporter htmlReporter;
	private static ExtentReports extent;

	private static Map<Integer, ExtentTest> dictExtentTestScenario = new HashMap<>();
	private static Map<Integer, String> dictExtentTestCase = new HashMap<>();
	private static Map<Integer, ArrayList<ExtentTestSteps>> dictExtentTestSteps = new HashMap<>();
	static final  Logger logger =LoggerFactory.getLogger(ExtentReport.class.getName()); 

	static String userNameFieldName = "Username";
	static String hostNameFieldName = "Host Name";
	static String environmentFieldName = "Environment";
	static String screenLoggerHeaderLine = "\"=================Screen: \"";
	static String moduleLoggerHeaderLine = "\"=================Module: \"";
	static String loggerHeaderLine = "\"==============\"";

	ExtentTest multiTCNode = null;
	ExtentTest moduleNode = null;
	ExtentTest screenNode = null;
	ExtentTest newParentNode=null;

	private static Map<String, Status> testCaseStatusMap = new HashMap<>();
	private static Map<String, ExtentTest> parentTestMap = new HashMap<>();

	//added for live report
	private static Map<String, ExtentTest> tcExtentMapping = new HashMap<>();	
	public static void startReport(String reportPath, String hostName, String environment, String userName)
	{
		try
		{
			reportPath = reportPath + "/ExtentReport_RunTime.html";



			htmlReporter = new ExtentHtmlReporter(reportPath);
			extent = new ExtentReports();
			extent.attachReporter(htmlReporter);
			extent.setSystemInfo(hostNameFieldName, hostName);
			extent.setSystemInfo(environmentFieldName, environment);
			extent.setSystemInfo(userNameFieldName, userName);
		}
		catch(Exception e)
		{
			logger.info("Unable to initialize the Extent Report");
			logger.info(Arrays.toString(e.getStackTrace()));
			logger.info(e.getMessage());



		}


	}

	public static void startReport(String reportPath, String environment)
	{
		try
		{
			reportPath = reportPath + "/ExtentReport_RunTime.html";

			htmlReporter = new ExtentHtmlReporter(reportPath);
			extent = new ExtentReports();
			extent.attachReporter(htmlReporter);
			extent.setSystemInfo(environmentFieldName, environment);
		}
		catch(Exception e)
		{
			logger.info("Unable to initialize the Extent Report");
			logger.info(Arrays.toString(e.getStackTrace()));
			logger.info(e.getMessage());



		}


	}

	public void startTest(String testCase, String moduleName, String browser, String testcaseDescription) throws InterruptedException {
		try {
			ExtentTest tc = extent.createTest(testCase, testcaseDescription);
			ExtentTest test = tc.createNode("Iteration==>1");
			test.assignCategory(browser);
			test.assignCategory(moduleName);
			Thread.sleep(0);
			tcExtentMapping.put(testCase + "_" + moduleName + "_" + browser, test);
			parentTestMap.put(testCase + "_" + moduleName + "_" + browser, tc); // Track the parent node
		} catch (InterruptedException e) {
			Thread.currentThread().interrupt(); // Preserve the interrupted status
			throw new InterruptedException("Thread was interrupted"); // Rethrow as a RuntimeException or handle appropriately
		} catch (Exception e) {
			logger.info("Unable to initialize the Extent Test Case ==> {}", testCase);
			logger.info(Arrays.toString(e.getStackTrace()));
			logger.info(e.getMessage());
			// Log the failure status
			ExtentTest tc = extent.createTest(testCase, testcaseDescription);
			tc.fail("Initialization failed: " + e.getMessage());
		}
	}

	public void startTest(String testCase,String moduleName, String browser, int iteration,String testcaseDescription)
	{

		try
		{
			ExtentTest test;

			if(iteration>1)
			{
				ExtentTest tc =tcExtentMapping.get(testCase+"_"+ moduleName+ "_" + browser);

				test =tc.createNode("Iteration==>"+iteration);
				tcExtentMapping.put(testCase+"_"+ moduleName+ "_" + browser, test);
			}
			else
			{
				ExtentTest tc = extent.createTest(testCase, testcaseDescription);
				test =tc.createNode("Iteration==>1");	

				test.assignCategory(browser);
				test.assignCategory(moduleName);

			}

			tcExtentMapping.put(testCase+"_"+ moduleName+ "_" + browser, test);
		}
		catch(Exception e)
		{
			logger.info("Unable to initialize the Extent Test Case ==> {} ", testCase );
			logger.info(Arrays.toString(e.getStackTrace()));
			logger.info(e.getMessage());



		}

	}

	public void addTestStepsLogs(TestCaseParam testCaseParam, TestStepDetails tsDetails) {
		try {
			String testCaseKey = testCaseParam.getTestCaseName() + "_" + testCaseParam.getModuleName() + "_" + testCaseParam.getBrowser();
			ExtentTest test = tcExtentMapping.get(testCaseKey);

			if (tsDetails.testStepType == TestStepDetails.TestStepType.TEST_STEP ||
					tsDetails.testStepType == TestStepDetails.TestStepType.VERIFICATION ||
					tsDetails.testStepType == TestStepDetails.TestStepType.EXCEPTION) {

				if (tsDetails.screenShotData.equals("")) {
					test.log(tsDetails.extentStatus, tsDetails.stepName);
				} else {
					if(tsDetails.extentStatus == Status.FAIL && !tsDetails.getErrorMessage().isEmpty() && !tsDetails.getErrorDetails().isEmpty()) {
							test.log(tsDetails.extentStatus, writeImageToReport(tsDetails.screenShotData, "STEP ACTION ==> " + tsDetails.stepName + getAccurateDescription(tsDetails.getStepDescription()) + "<br>" + " ERROR ==> " + tsDetails.getErrorMessage()
							+ "<br>" + " DETAIL LOGS ==> " + trimStackTrace(tsDetails.getErrorDetails())));
					}
				}

				// Update the test case status
				if (tsDetails.extentStatus == Status.FAIL) {
					testCaseStatusMap.put(testCaseKey, Status.FAIL);
				} else if (!testCaseStatusMap.containsKey(testCaseKey)) {
					testCaseStatusMap.put(testCaseKey, Status.PASS);
				}
			}

			extracted(tsDetails, test);
		} catch (Exception e) {
			logger.info("Unable to update the Extent Test Step ==> {} ", tsDetails.stepName);
			logger.info(Arrays.toString(e.getStackTrace()));
			logger.info(e.getMessage());
		}
	}

	private void extracted(TestStepDetails tsDetails, ExtentTest test) {
		if (tsDetails.testStepType == TestStepDetails.TestStepType.MODULE) {
			test.log(Status.INFO, moduleLoggerHeaderLine + tsDetails.moduleName + loggerHeaderLine);
		}

		if (tsDetails.testStepType == TestStepDetails.TestStepType.SCREEN) {
			test.log(Status.INFO, screenLoggerHeaderLine + tsDetails.screenName + loggerHeaderLine);
		}

		if (tsDetails.testStepType == TestStepDetails.TestStepType.MODULE_SCREEN) {
			test.log(Status.INFO, moduleLoggerHeaderLine + tsDetails.moduleName + loggerHeaderLine);
			test.log(Status.INFO, screenLoggerHeaderLine + tsDetails.screenName + loggerHeaderLine);
		}
	}


	public static void endReportTear() {
		for (Map.Entry<String, Status> entry : testCaseStatusMap.entrySet()) {
			String testCaseKey = entry.getKey();
			Status status = entry.getValue();
			ExtentTest parentTest = parentTestMap.get(testCaseKey);

			if (status == Status.FAIL) {
				parentTest.fail("Overall test case failed due to one or more failed steps.");
			} else {
				parentTest.pass("Overall test case passed.");
			}
		}

	}
	//live report end


	public static void startReport(String reportPath, String hostName, String environment, String userName, int... node)
	{
		if (node == null) {
			node = new int[1]; // Initialize with appropriate size
		}
		node[0] = 0;
		reportPath = reportPath + "../TestReports/ExtentReport.html";
		htmlReporter = new ExtentHtmlReporter(reportPath);
		extent = new ExtentReports();
		extent.attachReporter(htmlReporter);
		extent.setSystemInfo(hostNameFieldName, hostName);
		extent.setSystemInfo(environmentFieldName, environment);
		extent.setSystemInfo(userNameFieldName, userName);
	}


	public void startTest(String testScenarioName, String testScenarioDescription, int... node)
	{
		if (node == null) {
			node = new int[1]; // Initialize with appropriate size
		}
		node[0] = 0;
		extent.createTest(testScenarioName, testScenarioDescription);
	}

	public void endTest(String testcaseName, String testcaseDesc, String moduleName, int... node)
	{
		int key = (node == null || node.length == 0) ? 0 : node[0];

		int failCounter = 0;

		if (dictExtentTestScenario.containsKey(key))
		{
			ExtentTest testScenario = dictExtentTestScenario.get(key);
			ExtentTest  extentNode = testScenario.createNode(testcaseName, testcaseDesc);

			ArrayList<ExtentTestSteps> arrayListTestSteps = dictExtentTestSteps.get(key);

			for (ExtentTestSteps ETS : arrayListTestSteps)
			{
				if (ETS.getTestStepStatus() == ExtentConstants.TestStepStatus.PASS)
				{
					extentNode.createNode(ETS.getTestStepName(), ETS.getTestStepDesc()).pass(ETS.getTestStepName() + "--Passed");
				}

				else if (ETS.getTestStepStatus() == ExtentConstants.TestStepStatus.FAIL)
				{
					extentNode.createNode(ETS.getTestStepName(), ETS.getTestStepDesc()).fail(ETS.getTestStepName() + "--Failed");
					failCounter++;
				}

				else
				{
					extentNode.createNode(ETS.getTestStepName(), ETS.getTestStepDesc()).skip(ETS.getTestStepName() + "--Skipped");
				}
			}

			if (failCounter > 0)
			{
				extentNode.fail(testcaseName + "--Fail");
			}
			else
			{
				extentNode.pass(testcaseName + "--Pass");
			}


			extentNode.assignCategory(moduleName);

			if (key > 0)
			{
				extentNode.assignDevice("Node -- " + node[0]);
			}

		}
	}



	public void endReport()

	{

		//End Report

		extent.flush();

	}

	public void addTestStepsLogs(String testStepName, String testStepDesc, ExtentConstants.TestStepStatus testStepStatus, int... node)
	{
		int key = (node == null || node.length == 0) ? 0 : node[0];

		ExtentTestSteps extentTestSteps = new ExtentTestSteps();
		if (dictExtentTestSteps.containsKey(key))
		{
			ArrayList<ExtentTestSteps> arrayListTestSteps = dictExtentTestSteps.get(key);
			extentTestSteps.setTestStepName(testStepName);
			extentTestSteps.setTestStepDesc(testStepDesc);
			extentTestSteps.setTestStepStatus(testStepStatus);

			arrayListTestSteps.add(extentTestSteps);
			dictExtentTestSteps.putIfAbsent( key , arrayListTestSteps);
		}
		else
		{
			ArrayList<ExtentTestSteps> arrayListTestSteps = new ArrayList<>();
			extentTestSteps.setTestStepName(testStepName);
			extentTestSteps.setTestStepDesc(testStepDesc);
			extentTestSteps.setTestStepStatus(testStepStatus);

			arrayListTestSteps.add(extentTestSteps);
			dictExtentTestSteps.putIfAbsent(key, arrayListTestSteps);
		}

	}

	public void createExtentReportFromModel(String reportPath, InterfaceTestRun model,
			String testName, String environment, 
			String requestPath, String responsePath) {

		reportPath = reportPath + "/ExtentReport_Interface.html";
		ExtentHtmlReporter htmlReporterLcl = new ExtentHtmlReporter(reportPath);
		ExtentReports extentLcl = new ExtentReports();
		extentLcl.attachReporter(htmlReporterLcl);
		extentLcl.setSystemInfo(hostNameFieldName, testName);
		extentLcl.setSystemInfo(environmentFieldName, environment);
		extentLcl.setSystemInfo(userNameFieldName, "Automation");

		if (model != null) {
			for (InterfaceTCDetails node : model.getInterfaceTCDetails()) {
				processTestNode(node, extentLcl, requestPath, responsePath);
			}
		}
		extentLcl.flush();
	}

	private void processTestNode(InterfaceTCDetails node, ExtentReports extent, 
			String requestPath, String responsePath) {

		String[] names = generateNames(node);
		String testcaseName = names[0];
		String requestFileName = names[1];
		String responseFileName = names[2];

		ExtentTest test = extent.createTest(node.getModule() + "==>" + node.getInterfaceName());
		test.assignCategory(node.getModule());

		switch (node.getTestCaseStatus()) {
		case FAIL:
			logFail(test, testcaseName, requestPath, requestFileName, responsePath, responseFileName, node);
			break;
		case PASS:
			logPass(test, testcaseName, requestPath, requestFileName, responsePath, responseFileName, node);
			break;
		case SKIP:
			test.log(Status.SKIP, testcaseName);
			break;
		default:
			logPass(test, testcaseName, requestPath, requestFileName, responsePath, responseFileName, node);	
		}
	}

	private String[] generateNames(InterfaceTCDetails node) {
		String testcaseName;
		String requestFileName;
		String responseFileName;

		if (node.getIteration() > 1) {
			testcaseName = node.getModule() + "==>" + node.getInterfaceName() + "==> Iteration: " + node.getIteration();
			requestFileName = "Request_" + node.getModule() + "_" + node.getInterfaceName() + "_" + node.getIteration();
			responseFileName = "Response_" + node.getModule() + "_" + node.getInterfaceName() + "_" + node.getIteration();
		} else {
			testcaseName = node.getModule() + "==>" + node.getInterfaceName();
			requestFileName = "Request_" + node.getModule() + "_" + node.getInterfaceName();
			responseFileName = "Response_" + node.getModule() + "_" + node.getInterfaceName();
		}

		return new String[]{testcaseName, requestFileName, responseFileName};
	}

	private void logFail(ExtentTest test, String testcaseName, String requestPath, 
			String requestFileName, String responsePath, 
			String responseFileName, InterfaceTCDetails node) {
		test.log(Status.FAIL, testcaseName);
		test.log(Status.INFO, writeDataToTextFile(requestPath, requestFileName, node.getRequestData(), node.getFileFormat()));
		test.log(Status.INFO, writeDataToTextFile(responsePath, responseFileName, node.getResponseData(), node.getFileFormat()));
		test.log(Status.ERROR, node.getErrorMessage());
	}

	private void logPass(ExtentTest test, String testcaseName, String requestPath, 
			String requestFileName, String responsePath, 
			String responseFileName, InterfaceTCDetails node ) {
		test.log(Status.PASS, testcaseName);

		test.log(Status.INFO, writeDataToTextFile(requestPath, requestFileName, node.getRequestData(), node.getFileFormat()));
		test.log(Status.INFO, writeDataToTextFile(responsePath, responseFileName, node.getResponseData(), node.getFileFormat()));
		test.log(Status.PASS, node.getErrorMessage());
	}


	public void createExtentReportCategory(String reportPath, List<Map<UUID, TestCaseDetails>> testcaseRepository,
			String testName, String environment, String executedBy) throws InterruptedException {
		reportPath = reportPath + "/TestResults_Category.html";
		ExtentHtmlReporter htmlReporterLCL = new ExtentHtmlReporter(reportPath);
		ExtentReports extentLCL = new ExtentReports();
		try {
			extentLCL.attachReporter(htmlReporterLCL);
		} catch (Exception e) {
			logAndThrowException(e);
		}
		extentLCL.setSystemInfo(hostNameFieldName, testName);
		extentLCL.setSystemInfo(environmentFieldName, environment);
		extentLCL.setSystemInfo(userNameFieldName, executedBy);

		if (testcaseRepository != null) {
			for (Map<UUID, TestCaseDetails> dictTC : testcaseRepository) {
				TestCaseDetails testCaseDetails = dictTC.values().stream().findFirst().orElse(null);
				if (testCaseDetails != null) {
					processTestCaseDetails(testCaseDetails, extentLCL);
				}
			}
		}
		extentLCL.flush();
	}

	private void processTestCaseDetails(TestCaseDetails testCaseDetails, ExtentReports extent) {
		String testcaseName = testCaseDetails.testCaseName;
		String moduleName = testCaseDetails.module;
		String browser = testCaseDetails.browser;

		ExtentTest test = extent.createTest(testcaseName);
		test.assignCategory(moduleName);
		test.assignCategory(browser);

		List<TestStepDetails> testStepDetails = testCaseDetails.getStepDetails();
		ExtentTest parentNode = null;

		String businessArea = "";
		String screenName = "";
		boolean testFailed = false;
		for (TestStepDetails ts : testStepDetails) {
			parentNode = determineParentNode(test,  ts, businessArea, screenName);
			logTestStepDetails(parentNode, ts);
			if (ts.extentStatus == Status.FAIL) {   
				testFailed = true;
				}
		}
		if (testFailed) {
			test.fail("Test case failed due to one or more failed steps.");
		} 
		else {
			test.pass("Test case passed.");
		}
	}

	private ExtentTest determineParentNode(ExtentTest test, TestStepDetails ts, 
			String businessArea, String screenName) {

		switch (ts.testStepType) {
		case MULTI_TC:
			multiTCNode = test.createNode(ts.stepName);
			newParentNode = multiTCNode;
			break;
		case MODULE:
			moduleNode = createOrUpdateNode(test, multiTCNode, ts.moduleName, businessArea);
			newParentNode = moduleNode;
			break;
		case SCREEN:
			screenNode = createOrUpdateNode(test, moduleNode, ts.screenName, screenName);
			newParentNode = screenNode;
			break;
		case MODULE_SCREEN:
			moduleNode = createOrUpdateNode(test, multiTCNode, ts.moduleName, businessArea);
			screenNode = createOrUpdateNode(test, moduleNode, ts.screenName, screenName);
			newParentNode = screenNode;
			break;
		default:

			break;
		}
		return newParentNode;
	}

	private ExtentTest createOrUpdateNode(ExtentTest parent, ExtentTest existingNode, String nodeName, String existingName) {


		if (existingNode == null) {
			if (!existingName.equals(nodeName)) {
				existingNode = parent.createNode(nodeName);
			}
		} else {
			if (!existingName.equals(nodeName)) {
				existingNode = existingNode.createNode(nodeName);
			}
		}


		return existingNode;
	}


	private void logTestStepDetails(ExtentTest parentNode, TestStepDetails ts) {
		switch (ts.testStepType) {
		case TEST_STEP:
			logTestStep(parentNode, ts);
			break;
		case VERIFICATION:
			logVerificationStep(parentNode, ts);
			break;
		case EXCEPTION:
			logExceptionStep(parentNode, ts);
			break;
		default:
			logTestStep(parentNode, ts);

		}
	}

	private void logTestStep(ExtentTest parentNode, TestStepDetails ts) {
		if (ts.testStepFormat == TestStepDetails.TestStepFormat.PLAIN) {
			if (ts.screenShotData.isEmpty()) {
				parentNode.log(ts.extentStatus, ts.stepName);
			} else {
				if(ts.extentStatus == Status.FAIL) {

					if(!ts.getErrorMessage().isEmpty() && !ts.getErrorDetails().isEmpty()) {

						parentNode.log(ts.extentStatus, writeImageToReport(ts.screenShotData,"STEP ACTION ==> " + ts.stepName + getAccurateDescription(ts.getStepDescription())+ "<br>" + 
								"ERROR ==>" + ts.getErrorMessage() +"<br>" + 
								"DETAIL LOGS ==>" + trimStackTrace(ts.getErrorDetails())));
					}		
				}else{
					parentNode.log(ts.extentStatus, writeImageToReport(ts.screenShotData, ts.stepName + " " + ts.getStepDescription()));

				}
			}
		} 
	}


	private void logVerificationStep(ExtentTest parentNode, TestStepDetails ts) {
		parentNode.log(ts.extentStatus, ts.stepName);
		ExtentTest childNode = parentNode.createNode(ts.stepName);
		switch (ts.testStepFormat) {
		case XML:
			childNode.info(MarkupHelper.createCodeBlock(ts.expectedResponse, CodeLanguage.XML));
			childNode.log(ts.extentStatus, MarkupHelper.createCodeBlock(ts.actualResponse, CodeLanguage.XML));
			break;
		case JSON:
			childNode.info(MarkupHelper.createCodeBlock(ts.expectedResponse, CodeLanguage.JSON));
			childNode.log(ts.extentStatus, MarkupHelper.createCodeBlock(ts.actualResponse, CodeLanguage.JSON));
			break;
		default:
			if (ts.extentStatus == Status.PASS) {
				parentNode.pass(MarkupHelper.createLabel("Expected: " + ts.expectedResponse + "\nActual: " + ts.actualResponse, ExtentColor.GREEN));
			} else {
				parentNode.fail(MarkupHelper.createLabel("Expected: " + ts.expectedResponse + "\nActual: " + ts.actualResponse, ExtentColor.RED));
			}
			break;
		}
	}

	private void logExceptionStep(ExtentTest parentNode, TestStepDetails ts) {
		ExceptionInfo ei = new ExceptionInfo();
		ei.setExceptionName(ts.errorMessage);
		ei.setStackTrace(ts.errorDetails);
		parentNode.log(ts.extentStatus, ei);
		parentNode.log(ts.extentStatus, writeImageToReport(ts.screenShotData, ts.stepName));
	}

	private void logAndThrowException(Exception e) throws InterruptedException {
		if (logger.isInfoEnabled()) {
			logger.info(e.getMessage());
		}
		if (logger.isDebugEnabled()) {
			logger.debug(Arrays.toString(e.getStackTrace()));
		}
		throw new InterruptedException();
	}

	public void createExtentReportModel(String reportPath, List<Map<UUID, TestCaseDetails>> testCaseRepository,
			String testName, String environment, String executedBy) throws InterruptedException {
		try {
			reportPath = reportPath + "/TestResults.html";
			ExtentHtmlReporter htmlReporterLCL = new ExtentHtmlReporter(reportPath);
			ExtentReports extentLCL = new ExtentReports();
			extentLCL.attachReporter(htmlReporterLCL);
			extentLCL.setSystemInfo("Host Name", testName);
			extentLCL.setSystemInfo("Environment", environment);
			extentLCL.setSystemInfo("Username", executedBy);

			if (testCaseRepository != null) {
				for (Map<UUID, TestCaseDetails> dictTC : testCaseRepository) {
					TestCaseDetails testCaseDetails = dictTC.values().stream().findFirst().orElse(null);
					if (testCaseDetails != null) {
						processTestCaseDetails(extentLCL, testCaseDetails);
					}
				}
			}
			extentLCL.flush();
		} catch (Exception e) {
			logAndThrowException(e);
		}
	}

	private void processTestCaseDetails(ExtentReports extent, TestCaseDetails testCaseDetails) {
		String testcaseName = testCaseDetails.testCaseName;
		String moduleName = testCaseDetails.module;
		String browser = testCaseDetails.browser;

		ExtentTest test = extent.createTest(testcaseName);
		test.assignCategory(moduleName);
		test.assignCategory(browser);

		List<TestStepDetails> testStepDetails = testCaseDetails.getStepDetails();
		String businessArea = "";
		String screenName = "";

		for (TestStepDetails ts : testStepDetails) {
			if((ts.testStepType==TestStepDetails.TestStepType.TEST_STEP)||
					(ts.testStepType==TestStepDetails.TestStepType.VERIFICATION)||
					(ts.testStepType==TestStepDetails.TestStepType.EXCEPTION))
			{	
				logTestStep(test,ts);
			}

			if (ts.testStepType == TestStepDetails.TestStepType.MODULE && !businessArea.equals(ts.moduleName)) {
				businessArea = ts.moduleName;
				test.log(Status.INFO, moduleLoggerHeaderLine + businessArea + loggerHeaderLine);
			}

			if (ts.testStepType == TestStepDetails.TestStepType.SCREEN && !screenName.equals(ts.screenName)) {
				screenName = ts.screenName;
				test.log(Status.INFO, screenLoggerHeaderLine + screenName + loggerHeaderLine);
			}

			if (ts.testStepType == TestStepDetails.TestStepType.MODULE_SCREEN) {
				businessArea = ts.moduleName;
				screenName = ts.screenName;
				test.log(Status.INFO, moduleLoggerHeaderLine + businessArea + loggerHeaderLine +"<br>"+screenLoggerHeaderLine + screenName + loggerHeaderLine);
			}

		}
	}


	public void addScreenShotDetails(TestStepDetails ts,ExtentTest node )
	{
		if (ts.screenShotData.equals(""))
		{
			node.log(ts.extentStatus, ts.stepName);
		}
		else
		{
			node.log(ts.extentStatus,writeImageToReport(ts.screenShotData, ts.stepName));
		}		
	}

	public String writeDataToTextFile(String filePath, String fileName,String fileContent,String fileFormat)
	{
		filePath = filePath + File.separator + fileName + fileFormat;
		try(BufferedWriter writer = new BufferedWriter(new FileWriter(filePath)))
		{
			writer.write(fileContent);


			filePath= filePath.replace("\\","/");

			filePath = "<a href = 'file:///"+ filePath + "'>"+ fileName + "</a>";
			return filePath;
		}

		catch (Exception e)
		{
			return filePath;
		}
	}

	public String writeImageToReport(String filePath, String fileName)
	{

		try
		{
			filePath = filePath.replace("\\", "/");
			filePath = "<a href = 'file:///" + filePath + "'>" + fileName + "</a>";
			return filePath;
		}

		catch (Exception e)
		{
			return filePath;
		}
	}

	public static Map<Integer, String> getDictExtentTestCase() {
		return dictExtentTestCase;
	}

	public static void setDictExtentTestCase(Map<Integer, String> dictExtentTestCase) {
		ExtentReport.dictExtentTestCase = dictExtentTestCase;
	}
	public static String trimStackTrace(String logs) {
		String[] lines = logs.split(";");
		StringBuilder trimedTrace = new StringBuilder();
		for(String line : lines) {
			if(line.contains("java.base/jdk.internal.reflect")) {
				break;
			}
			trimedTrace.append(line);
		}
		return trimedTrace.toString().trim();
	}
	
	public static String getAccurateDescription(String desc) {

		if(desc.contains("By.xpath:")) {
			String [] res = desc.split("By.");
			return res[1].trim();
		}else {
			return "";
		}
	}

}