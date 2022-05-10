package org.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test1 {
	public static void main(String[] args) throws IOException {
		File f =new File("D:\\Song\\TESTNG WORKSPACE\\DataDrivenFrameworkPractice\\Excel\\TestData.xlsx");
		FileInputStream fin =new FileInputStream(f);
		Workbook w =new XSSFWorkbook(fin);
		Sheet s = w.getSheet("Sheet1");
		for (Row row : s) {
			for (Cell cell : row) {
				System.out.println(cell);
			}
		}
	}
}
