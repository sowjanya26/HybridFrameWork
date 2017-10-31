package utility;

import java.io.IOException;
import java.util.Hashtable;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.SkipException;

public class commonUtils {
	
	public static void checkExecution(String testSuiteName, String testCaseName,
										String dataRunMode, ExcelReader excel) throws InvalidFormatException, IOException{
		
		
		if(!isSuiteExecutable(testSuiteName)){
			throw new SkipException("Skipping the Test "+testCaseName+" as RunMode of suite "+testSuiteName+" is No");
		}
		if(!isTestExecutable(testCaseName,excel)){
			throw new SkipException("Skipping the Test "+testCaseName+" as RunMode of testCase  "+testCaseName+" is No");
		}
		if(dataRunMode.equalsIgnoreCase(constants.RUNMODE_NO)){
			throw new SkipException("Skipping the Test "+testCaseName+" as RunMode of dataRunMode is No");
		}
	}
	
	public static boolean isTestExecutable(String testCaseName, ExcelReader excel){
		String  sheetName = constants.TEST_CASE_SHEET_NAME;
		
		int rows = excel.getRowCount(sheetName);
		for(int rowNum = 2;rowNum <rows;rowNum++){
			String testCase = excel.getCellData(sheetName, constants.TEST_CASE_COL, rowNum);
			if(testCase.equalsIgnoreCase(testCaseName)){
				String runMode = excel.getCellData(sheetName, constants.TEST_CASE_RUN_MODE, rowNum);
				if(runMode.equalsIgnoreCase(constants.RUNMODE_YES)){
					return true;
				}else
					return false;
			}
		}
		return false;
		
	}
	public static boolean isSuiteExecutable(String suiteName) throws InvalidFormatException, IOException{
		
		ExcelReader excel = new ExcelReader(constants.TEST_SUITE_XLS_PATH);
		String sheetName = constants.TEST_SUITE_SHEET;
		int rows = excel.getRowCount(sheetName);
		for(int rowNum = 2;rowNum <rows;rowNum++){
			String testSuiteName = excel.getCellData(sheetName, constants.TEST_SUITE_COL, rowNum);
			if(testSuiteName.equalsIgnoreCase(suiteName)){
				String runMode = excel.getCellData(sheetName, constants.SUITE_RUN_MODE, rowNum);
				if(runMode.equalsIgnoreCase(constants.RUNMODE_YES)){
					return true;
				}else
					return false;
			}
		}
		return false;
		
	}

	public static Object[][] getData(String testcase, ExcelReader excel) throws InvalidFormatException, IOException{
		
		String sheetName = "testData";
		String testCase = testcase;
		
		
		//Test case start from 
		int testcaseRowNum =1;
		while(!excel.getCellData(sheetName, 0, testcaseRowNum).equalsIgnoreCase(testCase)){
			testcaseRowNum++;
		}
		System.out.println("TestCase Starts from : "+testcaseRowNum)	; 
		// total cols & rows  - clos starts and test data starts from
		int colsStartRowNum = testcaseRowNum+1;
		int dataStartRowNum = colsStartRowNum+1;
		
		//total columns in test are
		int cols =0;
		while(!excel.getCellData(sheetName, cols, colsStartRowNum).trim().equals("")){
			cols++;	
		}
		
		System.out.println("Total columns in a test case : "+cols);
		int rows =0;
		while(!excel.getCellData(sheetName, 0, dataStartRowNum+rows).trim().equals("")){
			rows++;
		}
		System.out.println("Total rows are : "+rows);
		//reading data from the sheet
		Object[][] data = new Object[rows][1];
		
		//String TestData;
		//String ColName;
		int i=0;
		for(int row = dataStartRowNum; row<dataStartRowNum+rows; row++){
			Hashtable<String,String> table = new Hashtable<String,String>();
			for(int col=0; col< cols; col++){
				//data[row-dataStartRowNum][col] = excel.getCellData(sheetName, col, row);
				String TestData = excel.getCellData(sheetName, col, row);
				String ColName = excel.getCellData(sheetName, col, colsStartRowNum);
				table.put(ColName, TestData);
			}
			data[i][0] = table;
			i++;			
		}
		return data;

	}
}
