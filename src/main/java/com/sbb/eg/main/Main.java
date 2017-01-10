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
		new Main().generateExcel(2017);
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
			XSSFRow row = null;

			while (calendar.get(Calendar.YEAR) == currentYear) {
				if (month != calendar.get(Calendar.MONTH)) {
					if (month >= 0){
						row = generateLastRow(day, sheet);
						addTable(sheet, 6, day, "TableStyleMedium9", "table" + months[month]);
					}
					day = 1;
					log.info("\n\n\n\n-----------------------------------\n\n\n\n");

					month = calendar.get(Calendar.MONTH);
					sheet = workbook.createSheet(months[month]);
					createHeaderRow(sheet);
				}

				if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
					calendar.add(Calendar.DAY_OF_MONTH, 1);
					continue;
				}

				row = sheet.createRow(day);
				day++;
				row.createCell(0).setCellValue(formatDate(calendar));

				createCell(row, 1, generateNumericDataStyle(workbook));
				createCell(row, 2, generateNumericDataStyle(workbook));
				createCell(row, 3, generateNumericDataStyle(workbook));

				createRowSumCell(day, row, 4, "B", "C");
				createRowSumCell(day, row, 5, "D", "E");

				log.info(formatDate(calendar));
				calendar.add(Calendar.DAY_OF_MONTH, 1);

			}

			addTable(sheet, 6, day, "TableStyleMedium9", months[month]);

			FileOutputStream fileOut = new FileOutputStream(filename);
			workbook.write(fileOut);
			workbook.close();
			fileOut.close();
			log.info("Your excel file has been generated!");

		} catch (Exception ex) {
			ex.printStackTrace();
			log.error(ex);
		}
	}

	private XSSFRow generateLastRow(short day, XSSFSheet sheet) {
		XSSFRow row;
		row = sheet.createRow(day);
		row.createCell(0).setCellValue(sumInfo);

		for (int i = 1; i < headers.length; i++) {
			char cellChar = (char) ('A' + i);
			row.createCell(i).setCellType(CellType.FORMULA);
			row.getCell(i).setCellFormula("SUM(" + cellChar + "2" + ":" + cellChar + (day) + ")");
			row.getCell(i).setCellStyle(generateNumericDataStyle(row.getSheet().getWorkbook()));
		}
		return row;
	}

	private void createHeaderRow(XSSFSheet sheet) {
		XSSFRow rowhead = sheet.createRow((short) 0);
		for (int i = 0; i < headers.length; i++)
			rowhead.createCell(i).setCellValue(headers[i]);
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

		XSSFTable table = sheet.createTable();

		CTTable cttable = table.getCTTable();

		CTTableStyleInfo ctTableStyle = cttable.addNewTableStyleInfo();
		ctTableStyle.setName(tableStyle);

		ctTableStyle.setShowColumnStripes(false); 
		ctTableStyle.setShowRowStripes(true); 

		AreaReference dataRange = new AreaReference(new CellReference(0, 0),
				new CellReference(rowRange, columnRange - 1));

		cttable.setRef(dataRange.formatAsString());
		cttable.setDisplayName(
				tableName); 
		cttable.setName(tableName);
		cttable.setId(1L); 

		CTTableColumns columns = cttable.addNewTableColumns();
		columns.setCount(columnRange); 

		for (int i = 0; i < columnRange; i++) {
			CTTableColumn column = columns.addNewTableColumn();
			column.setName(tableName + i);
			column.setId(i + 1);
		}
	}

	private static String formatDate(Date today) {
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
		String date = dateFormat.format(today);
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