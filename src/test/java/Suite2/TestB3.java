package Suite2;

import java.util.Hashtable;

import org.testng.annotations.Test;

import utility.dataProviders;

public class TestB3 {
	@Test(dataProviderClass = dataProviders.class,dataProvider = "getDataSuite2")
	public void testB3(Hashtable<String, String> data){
		
	}
}
