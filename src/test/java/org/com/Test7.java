package org.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test7 {
	public static void main(String[] args) throws IOException {
		File f =new File("D:\\Song\\TESTNG WORKSPACE\\DataDrivenFrameworkPractice\\Excel\\TestData.xlsx");
		FileInputStream fin =new FileInputStream(f);
		Workbook w =new XSSFWorkbook(fin);
		Sheet s = w.getSheet("New Sheet");
		Row r = s.getRow(3);
		Cell cell = r.getCell(1);
		String Value = cell.getStringCellValue();
		if(Value.equals("Hello User")) {
			cell.setCellValue("Naveen");
		}
		FileOutputStream fOut =new FileOutputStream(f);
		w.write(fOut);
		System.out.println("Done------------------");
}
}