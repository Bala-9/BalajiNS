package com.Data_driven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\ELCOT\\Desktop\\roughsheet.xlsx");
		
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fis);
		
		wb.createSheet("Emp_details").createRow(0).createCell(0).setCellValue("EMP NAME");
		
		wb.getSheet("Emp_details").getRow(0).createCell(1).setCellValue("EMP ID");
		
		wb.getSheet("Emp_details").createRow(1).createCell(0).setCellValue("BALAJI");
		
		wb.getSheet("Emp_details").getRow(1).createCell(1).setCellValue("Bala@123");
		
		wb.getSheet("Emp_details").createRow(2).createCell(0).setCellValue("NAVEEN");
		
		wb.getSheet("Emp_details").getRow(2).createCell(1).setCellValue("Navi@456");
		
		FileOutputStream fos = new FileOutputStream(f);
		
		wb.write(fos);
		
		System.out.println("Write Successfully");
		
		wb.close();
		
		
	}

}
