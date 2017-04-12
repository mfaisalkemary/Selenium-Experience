package com.stta.utility;

import java.io.IOException;

public class SuiteUtility {

	public static Boolean checkToRunUtility(Read_XLS xls,String sheetName, String ToRun,String Testsuite){
		Boolean Flag=false;
		if (xls.retrieveToRunFlag(sheetName, ToRun, Testsuite).equalsIgnoreCase("y")){
			Flag=true;
		}
		else{
			Flag=false;
			}
		return Flag;
	}
		public static String[] checkToRunUtilityOfData(Read_XLS xls, String sheetName, String ColName){		
			return xls.retrieveToRunFlagTestData(sheetName,ColName);		 	
		}
	 
		public static Object[][] GetTestDataUtility(Read_XLS xls, String sheetName){
		 	return xls.retrieveTestData(sheetName);	
		}
	 
		public static  void WriteResultUtility(Read_XLS xls, String sheetName, String ColName, int rowNum, String Result) throws IOException{			
			 xls.writeResult(sheetName, ColName, rowNum, Result);		 	
		}
	 
		
	}


