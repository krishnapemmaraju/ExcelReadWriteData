package com.test.qa.test;

import java.io.IOException;

import com.test.qa.utils.GetExcelData;

public class MainTestClass extends GetExcelData {
	
	public MainTestClass() {
		
		super("D:\\KrishnaData\\MyData.xlsx");
	}
	
	public static void main(String ar[]) throws IOException {
		MainTestClass mainTest = new MainTestClass();
	/*	System.out.println("The Row Count is " + GetExcelData.getRowCount("Sheet1"));	
	    System.out.println("Is the Sheet Created -- "  +GetExcelData.createSheet("Krishna"));
	//    System.out.println("Is the Sheet Removed -- "  +GetExcelData.removeSheet("Krishna"));
		System.out.println("Is Sheet Exists -- " + GetExcelData.isSheetExists("Krishna"));
		System.out.println("The No Of Columns are  -- " + GetExcelData.getColumnCount("Sheet1"));
	//	System.out.println("The Cell Number is  --     " + GetExcelData.getCellRowNum("Sheet1", "Username","Data"));
		System.out.println("The Cell Number removed  is  --     " + GetExcelData.removeColNum("Sheet1", 2));
		System.out.println("The Cell Values are  is  --     " + GetExcelData.getCellData("Sheet1", "password", 0));*/
//		System.out.println("The New Column Created like  --     " + GetExcelData.createNewColumn("Sheet1", "Test3"));
		System.out.println("The Cell Value to Set   --     " + GetExcelData.setCellData("Sheet1", "GYTHRT", "password",10));
	}

	
}
