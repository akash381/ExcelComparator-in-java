package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	static String path = "./data/";
	XSSFWorkbook workbook;
	public ExcelUtils(String fileName) {
		try {
			workbook = new XSSFWorkbook(path + fileName);
		}catch(Exception exp) {
			exp.printStackTrace();
		}
	}
	
	public Object getCellData(XSSFSheet sheet, int row, int column) {
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(row).getCell(column));
		return value;		
	}
	
}
