package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	static String excelPath = "./data/";
	static XSSFWorkbook workbook;
	static String file1 = "BEFORE PBRER 5.0.xlsx";
	static String file2 = "AFTER PBRER 5.0.xlsx";
	
	public static void main(String[] args) throws Exception{
		
		ExcelUtils excel1 = new ExcelUtils(file1);
		ExcelUtils excel2 = new ExcelUtils(file2);
		
		XSSFWorkbook workbook1 = new XSSFWorkbook(excelPath + file1);
		XSSFWorkbook workbook2 = new XSSFWorkbook(excelPath + file2);
		
		//Check if both the workbook have same number of sheets
		matchNumberOfSheets(workbook1, workbook2);	
		
		//Check for mismatch rowCount in both the sheet
		matchRowsOfSheets(workbook1, workbook2);
		
		//Getting a cell data
		//getCellData(workbook1);
		
	}
	
	public static void matchNumberOfSheets(XSSFWorkbook workbook1, XSSFWorkbook workbook2) {
		if(workbook1.getNumberOfSheets() == workbook2.getNumberOfSheets()) {
			System.out.println("Both the workbook have same number of sheets !");
		}else {
			System.out.println("No of sheets are different in both the workbook !"
					+ "\nPlease delete the extra sheet and then compare again !");
			return;
		}	
	}
	
	public static void matchRowsOfSheets(XSSFWorkbook workbook1, XSSFWorkbook workbook2) {
		//Assuming the number of sheets are same in both the workbook now.
		int noOfSheets = workbook1.getNumberOfSheets();
		
		for(int i = 0; i < noOfSheets; i++) {
			String sheetName1 = workbook1.getSheetName(i);
			String sheetName2 = workbook2.getSheetName(i);
			
			//Matching sheet Name
			if(sheetName1.equals(sheetName2)) {
				System.out.println("Both the sheet have same name : " + sheetName1);
			}else {
				System.out.println("Sheet number " + i + " has different names! : "
						+ "\nFirstSheetName : " + sheetName1 
						+ "\nSecondSheetName : " + sheetName2 );
			}
			
			XSSFSheet sheet1 = workbook1.getSheetAt(i);
			XSSFSheet sheet2 = workbook2.getSheetAt(i);
			
			int rowCountOfSheet1 = sheet1.getPhysicalNumberOfRows();
			int rowCountOfSheet2 = sheet2.getPhysicalNumberOfRows();
			
			if(rowCountOfSheet1 == rowCountOfSheet2) {
				System.out.println("Both the sheet have same number of rows for : " + sheetName1 
						+ " which is " + rowCountOfSheet1 + "/" + rowCountOfSheet2);
			}else {
				System.out.println("Sheets has different row count! : "
						+ "\n\t" + file1 + "." + sheetName1 + " : " + rowCountOfSheet1
						+ "\n\t" + file2 + "." + sheetName2 + " : " + rowCountOfSheet2);
			}	
		}
	}

}
