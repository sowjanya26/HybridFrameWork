package rough;

import java.io.IOException;
import java.util.Hashtable;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class testParameterization {
	
	@Test(dataProvider = "getData")
	public void testData(Hashtable<String,String> data1){
		//System.out.println(i+"--------"+t+"--------"+b+"-----------"+r);
		System.out.println(data1.get("Iteration")+"-----"+data1.get("TestData")+"-----"+data1.get("Browser")+"--------"+data1.get("RunMode"));
	}
	
	@DataProvider
	public static Object[][] getData() throws InvalidFormatException, IOException{
		ExcelReader excel = new ExcelReader(System.getProperty("user.dir")+"\\src\\test\\java\\rough\\testData.xls");
		System.out.println(System.getProperty("user.dir"));
		String sheetName = "testData";
		String testCase = "UserRegTest";
		
		
		//Test case start from 
		int testcaseRowNum =1;
		while(!excel.getCellData(sheetName, 0, testcaseRowNum).equalsIgnoreCase(testCase)){
			testcaseRowNum++;
		}
		System.out.println("TestCse Starts from : "+testcaseRowNum); 
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
