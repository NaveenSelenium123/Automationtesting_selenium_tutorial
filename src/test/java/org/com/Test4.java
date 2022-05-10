package org.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test4 {
	public static void main(String[] args) throws IOException {
		File f =new File("D:\\Song\\TESTNG WORKSPACE\\DataDrivenFrameworkPractice\\Excel\\TestData.xlsx");
		FileInputStream fin =new FileInputStream(f);
		Workbook w =new XSSFWorkbook(fin);
		Sheet s = w.getSheet("Sheet1");
		Row r = s.getRow(2);
		Cell c = r.getCell(4);
		int Type = c.getCellType();
		System.out.println(Type);
		if(Type == 1) {
			String stringCellValue = c.getStringCellValue();
			System.out.println(stringCellValue);
		}
		else if(Type == 0) {
			if(DateUtil.isCellDateFormatted(c)) {
				Date dateCellValue = c.getDateCellValue();
				SimpleDateFormat dateFormat =new SimpleDateFormat("dd-mm-yyyy");
				String format = dateFormat.format(dateCellValue);
				System.out.println(format);
			}
			
			else {
				double numericCellValue = c.getNumericCellValue();
				long l =(long) numericCellValue;
				String valueOf = String.valueOf(l);
				System.out.println(valueOf);
			}
		}
		
}
}