/**
 * 
 */
package com.deb.excelreader.client;

import com.deb.excelreader.data.ExcelData;
import com.deb.excelreader.parser.ExcelParcerImpl;
import com.deb.excelreader.parser.IExcelParser;

/**
 * @author MaRoy
 *
 */
public class ExcelClient {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		
		IExcelParser parser = new ExcelParcerImpl();
		ExcelData data = parser.getExcelData("test.xls");
		System.out.println(data);
	}

}
