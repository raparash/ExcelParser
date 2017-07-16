/**
 * 
 */
package com.oracle.excel.util.helper;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.oracle.excel.util.bean.ExcelBean;
import com.oracle.excel.util.bean.ExcelColumnPool;
import com.oracle.excel.util.bean.ExcelRow;
import com.oracle.excel.util.bean.ExcelSheet;

/**
 * @author raparash
 *
 */
public class ExcelParser {
	private static final Logger LOGGER = Logger.getLogger(ExcelParser.class.getName());

	/**
	 * Given a file name, the method parses the Excel Sheet
	 * 
	 * @param fileName
	 * @param sheetName
	 * @return
	 */
	public static ExcelSheet parseExcelSheet(String fileName, String sheetName) {
		ExcelSheet excelSheet = null;
		try {
			excelSheet = parseExcelSheet(new FileInputStream(fileName), sheetName);
			excelSheet.setFileName(fileName);
			LOGGER.log(Level.INFO, "Inside|parseExcelSheet(String fileName, String sheetName)|ExcelParser");
		} catch (Exception e) {
			LOGGER.log(Level.WARNING, "Exception|parseExcelSheet(String fileName, String sheetName)|ExcelParser ", e);
		}
		return excelSheet;
	}

	/**
	 * Given an inputstream and sheet name, this method parses the excel sheet
	 * @param inputstream
	 * @param sheetName
	 * @return
	 */
	public static ExcelSheet parseExcelSheet(InputStream inputstream, String sheetName) {
		Workbook workbook = null;
		Sheet sheet = null;
		boolean isHeader = true;
		ExcelColumnPool columnPool = new ExcelColumnPool();
		DataFormatter dataFormatter = new DataFormatter();
		ExcelSheet excelSheet = new ExcelSheet();
		List<ExcelRow> excelRows = new ArrayList<ExcelRow>();
		try {
			workbook =WorkbookFactory.create(inputstream);
			sheet = workbook.getSheet(sheetName);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row currentRow = rowIterator.next();
				LOGGER.log(Level.INFO,
						"Inside|parseExcelSheet(InputStream inputstream, String sheetName)|ExcelParser| Current row= "
								+ currentRow);
				if (currentRow != null) {
					if (isHeader) {
						columnPool.initColumnHeaders(currentRow.getPhysicalNumberOfCells());
						for (int i = 0; i < columnPool.getColumnPool().length; i++) {
							columnPool.getColumn(i).setName(dataFormatter.formatCellValue(currentRow.getCell(i)));
						}
						isHeader = false;
						LOGGER.log(Level.INFO,
								"Inside|parseExcelSheet(InputStream inputstream, String sheetName)|ExcelParser| Headers skipped");
						continue;
					}
					ExcelRow excelRow = new ExcelRow();
					for (int i = 0; i < columnPool.getColumnPool().length; i++) {
						excelRow.getMapOfColums().put(columnPool.getColumn(i),
								dataFormatter.formatCellValue(currentRow.getCell(i)));
					}
					excelRows.add(excelRow);
					excelRow.setColumnPool(columnPool);
				}
				excelSheet.setSheetName(sheetName);
				excelSheet.setExcelRows(excelRows);
			}
		} catch (Exception e) {
			LOGGER.log(Level.WARNING,
					"Exception|parseExcelSheet(InputStream inputstream, String sheetName)|ExcelParser ", e);
		} finally {
			CommonUtil.safeClose(workbook, inputstream);
		}
		return excelSheet;
	}

	/**
	 * Parses a given excel sheet stream and returns list of ExcelBean
	 * @param inputstream
	 * @param cls
	 * @return
	 */
	public static <T extends ExcelBean> List<T> parseExcel(InputStream inputstream, Class<T> cls) {
		List<T> listOfExcelBean = null;
		try {
			ExcelSheet excelSheet = parseExcelSheet(inputstream, "Sheet1");
			if (excelSheet != null) {
				listOfExcelBean = new ArrayList<T>();
				LOGGER.log(Level.WARNING,"Inside|parseExcel(InputStream inputstream)|ExcelParser| No. of excel rows = "+excelSheet.getExcelRows().size());
				for (ExcelRow excelRow : excelSheet.getExcelRows()) {
					T excelbean = cls.newInstance();
					excelbean.convertToBean(excelRow);
					listOfExcelBean.add(excelbean);
				}
			}
		} catch (Exception e) {
			LOGGER.log(Level.WARNING,
					"Exception|parseExcel(InputStream inputstream)|ExcelParser ", e);
		}
		
		return listOfExcelBean;
	}

}
