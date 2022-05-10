package org.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test3 {
	public static void main(String[] args) throws IOException {
		File f =new File("D:\\Song\\TESTNG WORKSPACE\\DataDrivenFrameworkPractice\\Excel\\TestData.xlsx");
		FileInputStream fin =new FileInputStream(f);
		Workbook w =new XSSFWorkbook(fin);
		Sheet s = w.getSheet("Sheet1");
		for (int i = 0; i < 2; i++) {
		Row row = s.getRow(i);	
		for (int j = 0; j <2; j++) {
			Cell cell = row.getCell(j);
			System.out.println(cell);
		}
		}
	}
}
