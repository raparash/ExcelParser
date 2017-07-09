/**
 * 
 */
package com.oracle.excel.util.test;

import com.oracle.excel.util.bean.ExcelRow;
import com.oracle.excel.util.bean.ExcelSheet;
import com.oracle.excel.util.helper.ExcelParser;

/**
 * @author raparash
 *
 */
public class ExcelTest {

	public static void main(String[] args) {
		ExcelSheet sheet=ExcelParser.parseExcelSheet("D:\\JiraSubtaskTemplate.xlsx", "Sheet1");
		for(ExcelRow row:sheet.getExcelRows()){
			System.out.println(row.toString());
			System.out.println(row.getColumnValue(0));
		}
	}

}
