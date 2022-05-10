package org.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	public static void main(String[] args) throws IOException {
		File f =new File("D:\\Song\\TESTNG WORKSPACE\\DataDrivenFrameworkPractice\\Excel\\TestData.xlsx");
		FileInputStream fin =new FileInputStream(f);
		Workbook w =new XSSFWorkbook(fin);
		Sheet s = w.getSheet("Sheet1");
	
		int physicalNumberOfRows = s.getPhysicalNumberOfRows();
		System.out.println("Row count" +physicalNumberOfRows);
		Row row1 = s.getRow(2);
		int physicalNumberOfCells = row1.getPhysicalNumberOfCells();
		System.out.println("cell count" +physicalNumberOfCells);
		System.out.println("============================");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row row = s.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				System.out.println(cell);	
			}
		}
	}

}
