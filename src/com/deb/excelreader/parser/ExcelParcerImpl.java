/**
 * 
 */
package com.deb.excelreader.parser;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.deb.excelreader.data.ExcelData;

/**
 * @author MaRoy
 *
 */
public class ExcelParcerImpl implements IExcelParser{

	@Override
	public ExcelData getExcelData(String relativeExcelPath) {
		
		FileInputStream fis = null;
		ExcelData dataMap = new ExcelData();
		
		try{
			
			fis = new FileInputStream(relativeExcelPath);

			//
			// Create an excel workbook from the file system.
			//
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			//
			// Get the first sheet on the workbook.
			//
			HSSFSheet sheet = workbook.getSheetAt(0);
			
			Iterator<Row> rows = sheet.rowIterator();
			boolean headerFlag = true;
			List<String> hdrList = new ArrayList<String>();
			
			while (rows.hasNext()) {
				HSSFRow row = (HSSFRow) rows.next();
				Iterator<Cell> cells = row.cellIterator();
				if(headerFlag){
					
					while (cells.hasNext()) {
						
						HSSFCell cell = (HSSFCell) cells.next();
						String hdrValue = cell.getStringCellValue();
						dataMap.put(hdrValue, new ArrayList<String>());
						hdrList.add(hdrValue);
						headerFlag = false;
					}
				}
				else{
					Iterator<String> hdrIterator = hdrList.iterator();
					while (cells.hasNext()) {
						HSSFCell cell = (HSSFCell) cells.next();
						String cellvalue = "";
						if(cell.getCellType() == 0)
							cell.setCellType(Cell.CELL_TYPE_STRING);
						cellvalue = cell.getStringCellValue();
						String headerName = hdrIterator.next();
						List<String> data = dataMap.get(headerName);
						data.add(cellvalue);
						dataMap.put(headerName, data);
					}
				}
				
			}
		}
		catch(Exception ex){
			ex.printStackTrace();
		}
		finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		return dataMap;
	}

}
