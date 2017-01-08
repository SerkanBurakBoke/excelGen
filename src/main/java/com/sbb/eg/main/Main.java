package com.sbb.eg.main;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

public class Main {
	static String[] days = { "Pazar", "Pazartesi", "Sal�", "�ar�amba", "Per�embe", "Cuma", "Cumartesi" };
	static String[] months = { "Ocak", "�ubat", "Mart", "Nisan", "May�s", "Haziran", "Temmuz", "A�ustos", "Eyl�l",
			"Ekim", "Kas�m", "Aral�k" };
	public static void main(String[] args) {
		try {

			String filename = "C:/NewExcelFile.xlsx";
			XSSFWorkbook workbook = new XSSFWorkbook();

			Calendar calendar = Calendar.getInstance();
			int currentYear = calendar.get(Calendar.YEAR);
			int month = -1;
			short day = 1;
			XSSFSheet sheet = null;
			
			while (calendar.get(Calendar.YEAR) == currentYear) {
				if (month != calendar.get(Calendar.MONTH)) {
					if(month >= 0)
					addTable(sheet, 6, day, "TableStyleMedium9",months[month]);
					day = 1;
					System.out.println("\n\n\n\n");
					System.out.println("-----------------------------------");
					System.out.println("\n\n\n\n");
					month = calendar.get(Calendar.MONTH);
					
					sheet = workbook.createSheet(months[month]);

					XSSFRow rowhead = sheet.createRow((short) 0);
					rowhead.createCell(0).setCellValue("Tarih");
					rowhead.createCell(1).setCellValue("Yap� Kredi");
					rowhead.createCell(2).setCellValue("Kuveyt T�rk");
					rowhead.createCell(3).setCellValue("Nakit");
					rowhead.createCell(4).setCellValue("KK");
					rowhead.createCell(5).setCellValue("Toplam");

				}
				
				if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
					calendar.add(Calendar.DAY_OF_MONTH, 1);
					continue;
				}
				
				XSSFRow row = sheet.createRow(day);
				day++;
				row.createCell(0).setCellValue(formatDate(calendar));

				row.createCell(1).setCellValue("");
				row.createCell(2).setCellValue("");
				row.createCell(3).setCellValue("");
				row.createCell(4);
				row.getCell(4).setCellType(Cell.CELL_TYPE_FORMULA);
				row.getCell(4).setCellFormula("SUM(B"+day+",C"+day+")");
				row.createCell(5);
				row.getCell(5).setCellType(Cell.CELL_TYPE_FORMULA);
				row.getCell(5).setCellFormula("SUM(D"+day+",E"+day+")");



				System.out.println(formatDate(calendar));
				calendar.add(Calendar.DAY_OF_MONTH, 1);
				if (month != calendar.get(Calendar.MONTH)) {
					row.createCell(0).setCellValue("");
					row.createCell(1).setCellValue("");
					row.createCell(2).setCellValue("");
					row.createCell(3).setCellValue("");
					row.createCell(4);
					row.getCell(4).setCellType(Cell.CELL_TYPE_FORMULA);
					row.getCell(4).setCellFormula("SUM(B"+day+",C"+day+")");
					row.createCell(5);
					row.getCell(5).setCellType(Cell.CELL_TYPE_FORMULA);
					row.getCell(5).setCellFormula("SUM(D"+day+",E"+day+")");
					
				}

			}
			
			addTable(sheet, 6, day, "TableStyleMedium9",months[month]);

			

			
			
			FileOutputStream fileOut = new FileOutputStream(filename);
			workbook.write(fileOut);
			workbook.close();
			fileOut.close();
			System.out.println("Your excel file has been generated!");

		} catch (Exception ex) {
			System.out.println(ex);
		}
	}

	private static void addTable(XSSFSheet sheet, int columnRange, int rowRange, String tableStyle, String tableName) {
		if(sheet == null)
			return;
		
		XSSFTable my_table = sheet.createTable();

		CTTable cttable = my_table.getCTTable();

		CTTableStyleInfo table_style = cttable.addNewTableStyleInfo();
		table_style.setName(tableStyle);

		table_style.setShowColumnStripes(false); // showColumnStripes=0
		table_style.setShowRowStripes(true); // showRowStripes=1

		AreaReference my_data_range = new AreaReference(new CellReference(0, 0), new CellReference(rowRange -1 , columnRange - 1));

		cttable.setRef(my_data_range.formatAsString());
		cttable.setDisplayName(
				tableName); /* this is the display name of the table */
		cttable.setName(tableName);
		cttable.setId(1L); // id attribute against table as long value

		CTTableColumns columns = cttable.addNewTableColumns();
		columns.setCount(columnRange); // define number of columns

		/* Define Header Information for the Table */
		for (int i = 0; i < columnRange; i++) {
			CTTableColumn column = columns.addNewTableColumn();
			column.setName(tableName + i);
			column.setId(i + 1);
		}
	}
	
	private static String formatDate(Date today) {
		SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("dd.MM.yyyy");
		String date = DATE_FORMAT.format(today);
		return date;
	}
	
	private static String formatDate(Calendar calendar) {
		String dateStr = formatDate(calendar.getTime());
		String dateSp[] = dateStr.split("\\.");
		dateStr = dateSp[0] + " " + months[calendar.get(Calendar.MONTH)] + " " + dateSp[2];
		return new StringBuilder().append(dateStr)
				.append(" ").append(days[calendar.get(Calendar.DAY_OF_WEEK) - 1]).toString();
	}

}