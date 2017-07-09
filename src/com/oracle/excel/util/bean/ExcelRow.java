/**
 * 
 */
package com.oracle.excel.util.bean;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import com.oracle.excel.util.helper.CommonUtil;

/**
 * @author raparash
 *
 */
public class ExcelRow {
	private Map<ExcelColumnHeader, String> mapOfColums=new LinkedHashMap<ExcelColumnHeader, String>();
	private ExcelColumnPool columnPool;
	/**
	 * @return the columnPool
	 */
	public ExcelColumnPool getColumnPool() {
		return columnPool;
	}

	/**
	 * @param columnPool the columnPool to set
	 */
	public void setColumnPool(ExcelColumnPool columnPool) {
		this.columnPool = columnPool;
	}

	/**
	 * @return the mapOfColums
	 */
	public Map<ExcelColumnHeader, String> getMapOfColums() {
		return mapOfColums;
	}

	/**
	 * @param mapOfColums the mapOfColums to set
	 */
	public void setMapOfColums(Map<ExcelColumnHeader, String> mapOfColums) {
		this.mapOfColums = mapOfColums;
	}
	
	/**
	 * Gets the value by column index
	 * @param columnId
	 * @return
	 */
	public String getColumnValue(int columnId){
		ExcelColumnHeader columnHeader=getColumnPool().getColumn(columnId);
		String value=CommonUtil.safeTrim(mapOfColums.get(columnHeader));
		return value;
	}
	
	/**
	 * Gets the value by column name
	 * @param columnName
	 * @return
	 */
	public String getColumnValue(String columnName){
		ExcelColumnHeader columnHeader=getColumnPool().getColumn(columnName);
		String value=CommonUtil.safeTrim(mapOfColums.get(columnHeader));
		return value;
	}

	/* (non-Javadoc)
	 * @see java.lang.Object#toString()
	 */
	@Override
	public String toString() {
		StringBuilder builder=new StringBuilder();
		for(Entry<ExcelColumnHeader,String> entry:mapOfColums.entrySet()){
			builder.append(entry.getKey().getName()+" : "+entry.getValue()+", ");
		}
		builder.deleteCharAt(builder.lastIndexOf(","));
		return builder.toString();
	}
	
	
	
}
