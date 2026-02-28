package com.revision;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class read_excel_revision {

	public static void main(String[] args) {
		readExcel();

	}

	public static void readExcel() {
		File f = new File(
				"C:\\Users\\Dell\\eclipse-workspace\\revision\\SeleniumAutomation\\src\\test\\resources\\Test\\Revison.xlsx");
		try {
			FileInputStream fs = new FileInputStream(f);
			Workbook wb = new XSSFWorkbook(fs);
			Sheet s = wb.getSheet("Sheet1");
//			Row row = s.getRow(1);
//			Cell cell = row.getCell(1);
			int physicalNumberOfRows = s.getPhysicalNumberOfRows();
			for (int i = 0; i < physicalNumberOfRows; i++) {
				Row row = s.getRow(i);
				int physicalNumberOfCells = row.getPhysicalNumberOfCells();
				for (int j = 0; j < physicalNumberOfCells; j++) {
					Cell cell = row.getCell(j);
//					System.out.println(cell);
					int cellType = cell.getCellType();
					if (cellType == 1) {
						String stringCellValue = cell.getStringCellValue();
						System.out.println(stringCellValue);
					} else if (cellType == 0) {
						if (DateUtil.isCellDateFormatted(cell)) {
							Date dateCellValue = cell.getDateCellValue();
							SimpleDateFormat ss = new SimpleDateFormat("dd/MM/yyyy");
							String format = ss.format(dateCellValue);
							System.out.println(format);
						}

					} else {
						
							double numericCellValue = cell.getNumericCellValue();
							int a = (int) numericCellValue;
							String valueOf = String.valueOf(a);
							System.out.println(valueOf);

						}
					}

			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
