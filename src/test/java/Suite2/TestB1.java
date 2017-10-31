package Suite2;

import java.io.IOException;
import java.util.Hashtable;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

import utility.ExcelReader;
import utility.commonUtils;
import utility.constants;
import utility.dataProviders;

public class TestB1 {

	@Test(dataProviderClass = dataProviders.class,dataProvider = "getDataSuite2")
	public void testB1(Hashtable<String, String> data) throws InvalidFormatException, IOException{
		ExcelReader excel = new ExcelReader(constants.SUITE2_XL_PATH);
		System.out.println("Runmode from data provider :" +data.get(constants.TEST_CASE_RUN_MODE));
		
		System.out.println("Test case is executable?:"+commonUtils.isTestExecutable("TestB1", excel));
		
		commonUtils.checkExecution("Suite2", "TestB1", data.get(constants.TEST_CASE_RUN_MODE), excel);
		System.out.println("Executed testcase "+this.getClass().getSimpleName()+" with data ---- Iteration :"+data.get("Iteration")
							+ " TestData :"+data.get("TestData")+" Browser :"+data.get("Browser")
							+ " RunMode :"+data.get("RunMode"));
	}
}
