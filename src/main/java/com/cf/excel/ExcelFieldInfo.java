package com.cf.excel;

import java.lang.reflect.Field;

import com.cf.excel.annotation.ExcelField;

public class ExcelFieldInfo {

	private ExcelField excelField;

	private Field field;

	public ExcelField getExcelField() {
		return excelField;
	}

	public void setExcelField(ExcelField excelField) {
		this.excelField = excelField;
	}

	public Field getField() {
		return field;
	}

	public void setField(Field field) {
		this.field = field;
	}

}
