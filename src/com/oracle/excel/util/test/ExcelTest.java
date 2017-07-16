/**
 * 
 */
package com.oracle.excel.util.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

import com.oracle.excel.util.bean.JiraSubtask;
import com.oracle.excel.util.helper.ExcelParser;

/**
 * @author raparash
 *
 */
public class ExcelTest {

	public static void main(String[] args) throws FileNotFoundException {
		List<JiraSubtask> list=ExcelParser.parseExcel(new FileInputStream("D:\\JiraSubtaskTemplate.xlsx"),JiraSubtask.class);
		for(JiraSubtask jiratask:list){
			System.out.println(jiratask);
		}
		//System.out.println(sheet.getExcelRows().get(0).getColumnPool().getColumn(1).getName());
		
	/*	for(ExcelRow row:sheet.getExcelRows()){
			System.out.println(row.toString());
			System.out.println(row.getColumnValue(0));
		}
		*/
		
		
	}

}
