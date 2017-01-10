package com.sbb.eg.main;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

public class Main {
	Logger log = Logger.getLogger(Main.class);
	static String[] days = { "Pazar", "Pazartesi", "Salý", "Çarþamba", "Perþembe", "Cuma", "Cumartesi" };
	static String[] months = { "Ocak", "Þubat", "Mart", "Nisan", "Mayýs", "Haziran", "Temmuz", "Aðustos", "Eylül",
			"Ekim", "Kasým", "Aralýk" };
	static String[] headers = { "Tarih", "Yapý Kredi", "Kuveyt Türk", "Nakit", "KK", "Toplam" };
	static String sumInfo = "Toplam";

	public static void main(String[] args) {
		BasicConfigurator.configure();
		new Main().generateExcel(2019);
	}

	private void generateExcel(int year) {
		try {

			String filename = "C:/ciro_" + year + ".xlsx";
			XSSFWorkbook workbook = new XSSFWorkbook();

			Calendar calendar = Calendar.getInstance();
			calendar.set(Calendar.YEAR, year);
			calendar.set(Calendar.DAY_OF_YEAR, 1);
			int currentYear = calendar.get(Calendar.YEAR);
			int month = -1;
			short day = 1;
			XSSFSheet sheet = null;

			while (calendar.get(Calendar.YEAR) == currentYear) {
				if (month != calendar.get(Calendar.MONTH)) {
					if (month >= 0)
						addTable(sheet, 6, day, "TableStyleMedium9", months[month]);
					day = 1;
					log.info("\n\n\n\n-----------------------------------\n\n\n\n");

					month = calendar.get(Calendar.MONTH);
					sheet = workbook.createSheet(months[month]);
					XSSFRow rowhead = sheet.createRow((short) 0);
					for (int i = 0; i < headers.length; i++)
						rowhead.createCell(i).setCellValue(headers[i]);
				}

				if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
					calendar.add(Calendar.DAY_OF_MONTH, 1);
					continue;
				}

				XSSFRow row = sheet.createRow(day);
				day++;
				row.createCell(0).setCellValue(formatDate(calendar));

				createCell(row, 1, generateNumericDataStyle(workbook));
				createCell(row, 2, generateNumericDataStyle(workbook));
				createCell(row, 3, generateNumericDataStyle(workbook));

				createRowSumCell(day, row, 4, "B", "C");
				createRowSumCell(day, row, 5, "D", "E");

				log.info(formatDate(calendar));
				calendar.add(Calendar.DAY_OF_MONTH, 1);
				if (month != calendar.get(Calendar.MONTH)) {
					row.createCell(0).setCellValue(sumInfo);

					for (int i = 1; i < headers.length; i++) {
						char cellChar = (char) ('A' + i);
						row.getCell(i).setCellType(CellType.FORMULA);
						row.getCell(i).setCellFormula("SUM(" + cellChar + "2" + ":" + cellChar + (day - 1) + ")");
					}

				}

			}

			addTable(sheet, 6, day, "TableStyleMedium9", months[month]);

			FileOutputStream fileOut = new FileOutputStream(filename);
			workbook.write(fileOut);
			workbook.close();
			fileOut.close();
			log.info("Your excel file has been generated!");

		} catch (Exception ex) {
			log.error(ex);
		}
	}

	private XSSFCellStyle generateNumericDataStyle(XSSFWorkbook workbook) {
		XSSFCellStyle style = workbook.createCellStyle();
		style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
		return style;
	}

	private void createRowSumCell(short day, XSSFRow row, int cellNum, String firstCell, String secondCell) {
		createCell(row, cellNum);
		row.getCell(cellNum).setCellType(CellType.FORMULA);
		row.getCell(cellNum).setCellFormula("SUM(" + firstCell + day + "," + secondCell + day + ")");
		row.getCell(cellNum).setCellStyle(generateNumericDataStyle(row.getSheet().getWorkbook()));
	}

	private void createCell(XSSFRow row, int cellNum, XSSFCellStyle style) {
		row.createCell(cellNum);
		row.getCell(cellNum).setCellType(CellType.NUMERIC);
		row.getCell(cellNum).setCellStyle(style);
	}

	private void createCell(XSSFRow row, int cellNum) {
		createCell(row, cellNum, generateNumericDataStyle(row.getSheet().getWorkbook()));
	}

	private static void addTable(XSSFSheet sheet, int columnRange, int rowRange, String tableStyle, String tableName) {
		if (sheet == null)
			return;

		XSSFTable my_table = sheet.createTable();

		CTTable cttable = my_table.getCTTable();

		CTTableStyleInfo table_style = cttable.addNewTableStyleInfo();
		table_style.setName(tableStyle);

		table_style.setShowColumnStripes(false); // showColumnStripes=0
		table_style.setShowRowStripes(true); // showRowStripes=1

		AreaReference my_data_range = new AreaReference(new CellReference(0, 0),
				new CellReference(rowRange - 1, columnRange - 1));

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
		return new StringBuilder().append(dateStr).append(" ").append(days[calendar.get(Calendar.DAY_OF_WEEK) - 1])
				.toString();
	}

}