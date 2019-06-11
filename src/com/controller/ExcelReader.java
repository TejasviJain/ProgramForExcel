package com.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	@SuppressWarnings("incomplete-switch")
	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream(new File("Test.xlsx"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
		
		for(Row row : sheet) {
			for(Cell cell : row)
			{
				switch(formulaEvaluator.evaluateInCell(cell).getCellType())
				{
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
				}
			}
			System.out.println();
		}
		wb.close();
	}

}
