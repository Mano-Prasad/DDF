package DDF;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil_2 {

	File f;
	FileInputStream fis;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	Object ob = null;

	public ExcelUtil_2(String filePath) throws IOException, InvalidFormatException {
		f = new File(filePath);
		wb = new XSSFWorkbook(f);
	}

	public int rowCount(String sheetName) {
		sheet = wb.getSheet(sheetName);
		int lastRowNum = sheet.getLastRowNum();
		lastRowNum = lastRowNum + 1;
		return lastRowNum;
	}

	public short columCount(String sheetName) {
		sheet = wb.getSheet(sheetName);
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		return lastCellNum;
	}

	public String getExcelData(String sheetName, int r, int c) {
		sheet = wb.getSheet(sheetName);
		String stringCellValue = sheet.getRow(r).getCell(c).getStringCellValue();
		return stringCellValue;
	}

}
