package com.Data_driven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_data {

	public static void Particular_Cell_Data() throws IOException {

		File f = new File("C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\userdata.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet s = wb.getSheetAt(0);

		Row r = s.getRow(2);

		Cell c = r.getCell(2);

		CellType type = c.getCellType();

		if (type.equals(CellType.STRING)) {

			String cellValue = c.getStringCellValue();

			System.out.println(cellValue);

		} else if (type.equals(CellType.NUMERIC)) {

			double cellValue = c.getNumericCellValue();

			int value = (int) cellValue;

			System.out.println(value);

		}
		wb.close();

	}

	public static void All_Data() throws IOException {

		File f = new File("C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\userdata.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet s = wb.getSheetAt(0);

		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {

			Row r = s.getRow(i);

			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {

				Cell c = r.getCell(j);

				CellType type = c.getCellType();

				if (type.equals(CellType.STRING)) {

					String value = c.getStringCellValue();

					System.out.println(value);

				} else if (type.equals(CellType.NUMERIC)) {

					double numericCellValue = c.getNumericCellValue();

					int v = (int) numericCellValue;

					String value = Integer.toString(v);

					System.out.println(value);

				}
				wb.close();

			}

		}

	}

	public static void Read_Particular_Row() throws IOException {

		File f = new File("C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\userdata.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet s = wb.getSheetAt(0);

		Row r = s.getRow(2);

		int numberOfCells = r.getPhysicalNumberOfCells();

		for (int i = 0; i < numberOfCells; i++) {

			Cell c = r.getCell(i);

			CellType type = c.getCellType();

			if (type.equals(CellType.STRING)) {

				String value = c.getStringCellValue();

				System.out.println(value);

			} else if (type.equals(CellType.NUMERIC)) {

				double numericCellValue = c.getNumericCellValue();

				int v = (int) numericCellValue;

				String value = Integer.toString(v);

				System.out.println(value);
			}
			wb.close();
		}

	}

	public static void Read_Particular_column() throws IOException {

		File f = new File("C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\userdata.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet s = wb.getSheetAt(0);

		int numberOfRows = s.getPhysicalNumberOfRows();

		for (int i = 0; i < numberOfRows; i++) {

			Row r = s.getRow(i);

			Cell c = r.getCell(1);

			CellType type = c.getCellType();

			if (type.equals(CellType.STRING)) {

				String value = c.getStringCellValue();

				System.out.println(value);
			}

			else if (type.equals(CellType.NUMERIC)) {

				double numericCellValue = c.getNumericCellValue();

				int v = (int) numericCellValue;

				String value = Integer.toString(v);

				System.out.println(value);
			}
			wb.close();
		}

	}

	public static void main(String[] args) throws IOException {

		System.out.println("****particular cell data****");
		Particular_Cell_Data();

		System.out.println("\n****All data****");
		All_Data();

		System.out.println("\n****particular Row data****");
		Read_Particular_Row();

		System.out.println("\n****particular Column data****");
		Read_Particular_column();

	}

}
