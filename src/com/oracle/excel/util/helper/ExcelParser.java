/**
 * 
 */
package com.oracle.excel.util.helper;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.oracle.excel.util.bean.ExcelColumnPool;
import com.oracle.excel.util.bean.ExcelRow;
import com.oracle.excel.util.bean.ExcelSheet;

/**
 * @author raparash
 *
 */
public class ExcelParser {
	
	
	public static Workbook getExcelWorkbook(String fileName){
		Workbook workbook=null;
		try{
			workbook=new HSSFWorkbook(new FileInputStream(fileName));
			System.out.println("HSSFWorkbook found");
		}catch(Exception e){
			System.out.println(e.getMessage());
			try {
				workbook=new XSSFWorkbook(new FileInputStream(fileName));
			} catch (Exception e1) {
				System.out.println(e1.getMessage());
			}
		}
		return workbook;
	}

	/**
	 * parses an Excel Sheet
	 * @param fileName
	 * @param sheetName
	 * @return
	 */
	public static ExcelSheet parseExcelSheet(String fileName, String sheetName)
	{
		Workbook workbook = null;
		Sheet sheet = null;
		boolean isHeader = true;
		ExcelColumnPool columnPool = new ExcelColumnPool();
		DataFormatter dataFormatter = new DataFormatter();
		ExcelSheet excelSheet = new ExcelSheet();
		List<ExcelRow> excelRows = new ArrayList<ExcelRow>();
		try {
			workbook = getExcelWorkbook(fileName);
			sheet = workbook.getSheet(sheetName);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row currentRow = rowIterator.next();
				if(currentRow!=null){
				if (isHeader) {
					columnPool.initColumnHeaders(currentRow.getPhysicalNumberOfCells());
					for (int i = 0; i < columnPool.getColumnPool().length; i++) {
						columnPool.getColumn(i).setName(dataFormatter.formatCellValue(currentRow.getCell(i)));
					}
					isHeader = false;
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
			excelSheet.setFileName(fileName);
			excelSheet.setSheetName(sheetName);
			excelSheet.setExcelRows(excelRows);
			}
		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		} finally {
			CommonUtil.safeClose(workbook);		
		}
		return excelSheet;
	}

}
