package DDF;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WIthExcel3 {

	public static void main(String[] args) throws InvalidFormatException, IOException {
	
		String filePath = "D:\\DDF sheet\\Demo.xlsx"; 
		File f = new File(filePath);
		XSSFWorkbook wb = new XSSFWorkbook(f);
		XSSFSheet sheet = wb.getSheet("Sheet1");

		int totalRow = sheet.getLastRowNum();
		System.out.println(totalRow);
		
		short totalColumn = sheet.getRow(0).getLastCellNum();
		System.out.println(totalColumn);
		
		for(int i = 0; i<=totalRow ;i++) {
			String firstColumn = sheet.getRow(i).getCell(0).getStringCellValue();
			System.out.println(firstColumn);
		}
		
		System.out.println("******************");
		
		for (int j = 0; j <totalColumn; j++) {
			String firstRow = sheet.getRow(0).getCell(j).getStringCellValue();
			System.out.println(firstRow);
			
		}
		
		for(int i=0;i<=totalRow;i++) {
			System.out.println("-------------------");
			for (int j = 0; j <totalColumn; j++) {
				String cellValue = sheet.getRow(i).getCell(j).getStringCellValue();
				System.out.println(cellValue);
			}	
		}
		
		System.out.println("-----------------");
		
		for (int j = 0; j <=totalRow; j++) {
			int rowNum = sheet.getRow(j).getLastCellNum();
			System.out.println(rowNum);
		}
		

	}

}
