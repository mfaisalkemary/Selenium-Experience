package com.stta.utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import javax.management.RuntimeErrorException;
import org.apache.bcel.generic.RETURN;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_XLS {

	String FilePath;
	FileInputStream Input;
	FileOutputStream Output;
	XSSFWorkbook WBook;
	XSSFSheet Sheet;
	
	
	public Read_XLS(String FilePath){
		this.FilePath=FilePath;
		try{
			Input = new FileInputStream(FilePath);
			WBook = new XSSFWorkbook(Input);
			Sheet = WBook.getSheetAt(0);
			Input.close();
		}
		catch (Exception e){
			e.printStackTrace();
		}
	}
	
	public int retrieveNoOfRows(String wsName){
		int SheetIndex=WBook.getSheetIndex(wsName);
		if (SheetIndex == -1){
			return 0;
		}
		else
		{
			Sheet=WBook.getSheetAt(SheetIndex);
			int RowCount=Sheet.getLastRowNum();
			return RowCount+1;
		}
		
		
	}
	
	public int retrieveNoOfCols(String wsName){
		int SheetIndex = WBook.getSheetIndex(wsName);
		if (SheetIndex == -1){
			return 0;
		}
		else
		{
			Sheet=WBook.getSheetAt(SheetIndex);
			int Colnum = Sheet.getRow(0).getLastCellNum();
			return Colnum;
		}
	}
	
	
	public String retrieveToRunFlag(String wsName, String colName, String rowName){
		
		int Sheetindex = WBook.getSheetIndex(wsName);
		if (Sheetindex == -1){
			return "Wrong Sheet Name";
		}
		
		int ColNum = retrieveNoOfCols(wsName);
		int RowNum = retrieveNoOfRows(wsName);
		int ReqColNum=-1;
		int ReqRowNum=-1;
		XSSFRow Row = Sheet.getRow(0);

		for (int i=0 ; i<ColNum ; i++){
		if (Row.getCell(i).getStringCellValue().equals(colName.trim()))
		{
			//System.out.println(Row.getCell(i).getStringCellValue());
			ReqColNum=i;
		}
		}
		
		 if (ReqColNum==-1){
		 	return "Wrong Column Name";
		}
			
		
		for(int j =0;j<RowNum;j++){
			if (Sheet.getRow(j).getCell(0).toString().equals(rowName)){
             ReqRowNum =j;				
			}
		}
		
		
			if (ReqRowNum ==-1){
				return "Wrong Testcase or Test Suite Name ";
			}
		
				
		XSSFRow ReqRow = Sheet.getRow(ReqRowNum);
		XSSFCell ReqCell = ReqRow.getCell(ReqColNum);
		String ReqData= ReqCell.getStringCellValue();
		
		if (ReqData == null){
			return "  ";
		}
		return ReqData;
		
	}
	
	
	
	public String[] retrieveToRunFlagTestData(String wsName, String colName){
		int index = WBook.getSheetIndex(wsName);
		if (index == -1){
			System.out.println("Wrong Sheet Name");
		}
		
		
		int Col= retrieveNoOfCols(wsName);
		int Row =retrieveNoOfRows(wsName);
		String [] Data = new String [Row-1];
		int ReqCol=-1;
		
	    XSSFRow Row2 = Sheet.getRow(0);
		for (int i=0;i<Col;i++){
			if (Row2.getCell(i).getStringCellValue().equals(colName)){
		   ReqCol=i;		
			}
			}
		
		if (ReqCol==-1){
		System.out.println("Wrong Column Name");	
		}
		 
		DataFormatter formatter= new DataFormatter();
		for (int j =0;j<Row-1;j++){
			XSSFRow Row3=Sheet.getRow(j+1);
			XSSFCell Cell3=Row3.getCell(ReqCol,Row3.RETURN_BLANK_AS_NULL);
			
			if(Cell3 == null){
				Data[j]="";
			}
			else 
			Data[j]=formatter.formatCellValue(Cell3);
		}
		return Data;
	}
	
	public Object[][] retrieveTestData(String wsName){
	int Rows = retrieveNoOfRows(wsName);
	int Cols=retrieveNoOfCols(wsName);
	
	Object [][] Data = new Object [Rows][Cols];
	for (int i =0;i<Rows;i++){
	XSSFRow Row=Sheet.getRow(i);
		for (int j =0 ; j<Cols ; j++){
			if (Row==null){
				Data[i][j]="";
			}
			
			else{
			XSSFCell Cell = Row.getCell(j);
				if (Cell==null){
				Data[i][j]="";
			     }	
			  else {
      //  String Data1 = Sheet.getRow(i).getCell(j).toString();
        Data[i][j]=Sheet.getRow(i).getCell(j).toString();        
		}
		
	
	}
	}
	}
	return Data;
	
		
	}

	
	public  void  writeResult(String wsName, String colName, int rowNumber, String Result) throws IOException{
		int ReqColNumber =-1;
		int colnum = retrieveNoOfCols(wsName);
		XSSFRow Row = Sheet.getRow(0);
		
		for (int i = 0; i<colnum ; i++){
			if (Row.getCell(i).getStringCellValue().equals(colName)){
				ReqColNumber=i;
				System.out.println(Row.getCell(i).getStringCellValue());
			}
		}
			if(ReqColNumber == -1){
				System.out.println("Wrong col name");
			}
			XSSFCell Cell = Sheet.getRow(rowNumber).getCell(ReqColNumber);
			Cell.setCellValue(Result);
			Output = new FileOutputStream(FilePath);
			WBook.write(Output);
			Output.flush();
			Output.close();
		
		
		
	}
	
	public static void main (String [] Args) throws IOException{
		
		Read_XLS Read = new Read_XLS("C:\\Users\\Abo Mazen\\workspace\\WDDF\\src\\com\\stta\\ExcelFiles\\SuiteTwo - Copy.xlsx");
		int Row=Read.retrieveNoOfRows("TestCasesList");
		int Colnum = Read.retrieveNoOfCols("TestCasesList");
		//String [] Data1= Read.retrieveToRunFlagTestData("TestCasesList","CaseToRun");
		//Object [][] Data2 = Read.retrieveTestData("TestCasesList");
		Read.writeResult("TestCasesList", "CaseToRun", 6, "hello");
		
		
		//String Data =Read.retrieveToRunFlag("TestCasesList","CaseToRun","Test1");
	//	System.out.println("Rows: "+Row+"     "+"Columns:  " +Colnum+"     "+Arrays.deepToString(Data2));
		
	}
	
}
