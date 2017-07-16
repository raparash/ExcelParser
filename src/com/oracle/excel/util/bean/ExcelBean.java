/**
 * 
 */
package com.oracle.excel.util.bean;

/**
 * 
 * Marker interface to define Excel related java beans
 * @author raparash
 *
 */
public interface ExcelBean {
	
	/**
	 * Unmarshalls ExcelRow to ExcelBean
	 * @param excelRow
	 * @return
	 */
	public ExcelBean convertToBean(ExcelRow excelRow);

}
