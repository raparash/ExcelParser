/**
 * 
 */
package com.oracle.excel.util.bean;

import com.oracle.excel.util.helper.CommonUtil;

/**
 * @author raparash
 *
 */
public class ExcelColumnPool {

	private ExcelColumnHeader[] columnPool;
	
	public void initColumnHeaders(int poolSize){
		columnPool=new ExcelColumnHeader[poolSize];
		for(int i=0;i<poolSize;i++){
			columnPool[i]=new ExcelColumnHeader();
		}
	}
	public ExcelColumnHeader getColumn(int columnId){
		return columnPool[columnId];
	}
	
	/**
	 * Returns ExcelColumnHeader for a given columnName if present
	 * @param columnName
	 * @return
	 */
	public ExcelColumnHeader getColumn(String columnName){
		ExcelColumnHeader columnHeader=null;
		columnName=CommonUtil.safeTrim(columnName).toUpperCase();
		for(ExcelColumnHeader column:columnPool){
			if(column.getColumnId()==columnName.hashCode()){
				columnHeader=column;
				break;
			}
		}
		return columnHeader;
	}
	
	/**
	 * @return the columnPool
	 */
	public ExcelColumnHeader[] getColumnPool() {
		return columnPool;
	}
	/**
	 * @param columnPool the columnPool to set
	 */
	public void setColumnPool(ExcelColumnHeader[] columnPool) {
		this.columnPool = columnPool;
	}
	
	
}
