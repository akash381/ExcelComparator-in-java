package utils;

import java.util.Iterator;
import java.io.FileWriter;   
//import java.io.IOException;  

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCompare {
	static String file1 = "BEFORE PBRER 5.0.xlsx";
	static String file2 = "AFTER PBRER 5.0.xlsx";
	static String logFile = "Log.txt";
	
	public static void main(String[] args) throws Exception{
		new FileWriter("./data/" + logFile, false).close(); //clearing the file
		FileWriter myWriter = new FileWriter("./data/" + logFile);
		
		ExcelUtils excel1 = new ExcelUtils(file1);
		ExcelUtils excel2 = new ExcelUtils(file2);
		
		//Check if both the workbook have same number of sheets
		matchNumberOfSheets(excel1.workbook, excel2.workbook, myWriter);	
		
		//Check for mismatch rowCount in both the sheet
		//matchRowsOfSheets(excel1.workbook, excel2.workbook);
		
		//Getting a cell data
		//XSSFSheet sheet = excel1.workbook.getSheet("Criteria Sheet");
		//System.out.println(excel1.getCellData(sheet,1,0));
		
		//Comparing two sheets
		//compareSheet(excel1.workbook.getSheetAt(1), excel2.workbook.getSheetAt(1));
		
		//Comparing the sheets
		compareExcel(excel1.workbook, excel2.workbook, myWriter);
		myWriter.close();
		
	}
	
	
	public static void compareExcel(XSSFWorkbook workbook1, XSSFWorkbook workbook2, FileWriter myWriter) throws Exception{
		int noOfSheets = workbook1.getNumberOfSheets();
		for(int i = 1; i < noOfSheets; i++) {
//			String sheetName1 = workbook1.getSheetName(i);
//			String sheetName2 = workbook2.getSheetName(i);
			
			XSSFSheet sheet1 = workbook1.getSheetAt(i);
			XSSFSheet sheet2 = workbook2.getSheetAt(i);
			
			compareSheet(sheet1, sheet2, myWriter);
			
		}
	}
	
	public static void compareSheet (XSSFSheet sheet1, XSSFSheet sheet2, FileWriter myWriter) throws Exception {
		
		myWriter.write("----------Comparing " + sheet1.getSheetName() + "----------\n");
		
		//System.out.println("----------Comparing " + sheet1.getSheetName() + "----------");
		
		Iterator<Row> itr1 = sheet1.iterator();  //iterating over rows of sheet1
		Iterator<Row> itr2 = sheet2.iterator();  //iterating over rows of sheet2
		int mismatchCount = 0;
		int r = 0;
		int c = 0;
		while (itr1.hasNext() && itr2.hasNext())                 
		{  	
			r++;
			Row row1 = itr1.next();  
			Row row2 = itr2.next();  
			Iterator<Cell> cellIterator1 = row1.cellIterator();   //iterating over columns of sheet1
			Iterator<Cell> cellIterator2 = row2.cellIterator();   //iterating over columns of sheet2
			while (cellIterator1.hasNext() && cellIterator2.hasNext())   
			{   c++;
				Cell celldata1 = cellIterator1.next();
				Cell celldata2 = cellIterator2.next();
				DataFormatter df = new DataFormatter();
				Object value1 = df.formatCellValue(celldata1);
				Object value2 = df.formatCellValue(celldata2);
				
				if(value1.toString().contains("Run Date and Time")) {
					continue;
				}
				if (value1.equals(value2)) {
					//System.out.println("Match found in cell " + value1 + " : " + value2);  
				}
				else {
					mismatchCount++;
					myWriter.write("Mismatch found in cell R" + r + "/C" + c + " " + value1 + " : " + value2 +"\n\n");
					//System.out.println("Mismatch found in cell " + r + "/" + c + " " + value1 + " : " + value2);  
				}
			}  
			c = 0;
		} 
		if(mismatchCount == 0) {
			myWriter.write("Everything matched in sheet : " + sheet1.getSheetName() + "\n");
			//System.out.println("Everything matched in sheet : " + sheet1.getSheetName());
		}
		myWriter.write("--------------------------------------------------------------\n\n\n");
		//System.out.println("------------------------------------------------------------\n");
	}	
	
	public static void matchNumberOfSheets(XSSFWorkbook workbook1, XSSFWorkbook workbook2, FileWriter myWriter) throws Exception{
		if(workbook1.getNumberOfSheets() == workbook2.getNumberOfSheets()) {
			myWriter.write("Both the workbook have same number of sheets !\n\n");
			//System.out.println("Both the workbook have same number of sheets !");
		}else {
			myWriter.write("No of sheets are different in both the workbook !"
					+ "\nPlease delete the extra sheet and then compare again !\n\n");
			//System.out.println("No of sheets are different in both the workbook !"
			//		+ "\nPlease delete the extra sheet and then compare again !");
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
