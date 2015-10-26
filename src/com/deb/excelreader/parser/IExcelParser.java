package com.deb.excelreader.parser;

import com.deb.excelreader.data.ExcelData;

public interface IExcelParser {

	public ExcelData getExcelData(String relativeExcelPath);
}
