package utility;

import java.io.IOException;
import java.lang.reflect.Method;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.DataProvider;

public class dataProviders {
	
	@DataProvider(name = "getDataSuite1")
	public static Object[][] getDataSuite1(Method m) throws InvalidFormatException, IOException{
		
		ExcelReader excel = new ExcelReader(constants.SUITE1_XL_PATH);
		String testCase = m.getName();
		return commonUtils.getData(testCase, excel);
		
	}
	
	@DataProvider(name = "getDataSuite2")
	public static Object[][] getDataSuite2(Method m) throws InvalidFormatException, IOException{
		
		ExcelReader excel = new ExcelReader(constants.SUITE2_XL_PATH);
		String testCase = m.getName();
		return commonUtils.getData(testCase, excel);
		
	}

}
