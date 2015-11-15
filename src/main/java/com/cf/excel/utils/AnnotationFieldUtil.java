package com.cf.excel.utils;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

import com.cf.excel.ExcelFieldInfo;
import com.cf.excel.annotation.ExcelField;
import com.cf.excel.annotation.ExcelObject;

public class AnnotationFieldUtil {
	public static Map<String, ExcelFieldInfo> fieldMaps(Class<?> clazz) {
		Map<String, ExcelFieldInfo> fieldMaps = new HashMap<String, ExcelFieldInfo>();
		ExcelObject excelObject = clazz.getAnnotation(ExcelObject.class);
		if (null == excelObject || !excelObject.value()) {
			return fieldMaps;
		}
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0; i < fields.length; i++) {
			fields[i].setAccessible(true);
			ExcelField excelField = fields[i].getAnnotation(ExcelField.class);
			if (null != excelField) {
				ExcelFieldInfo excelFieldVO = new ExcelFieldInfo();
				excelFieldVO.setField(fields[i]);
				excelFieldVO.setExcelField(excelField);
				fieldMaps.put(excelField.index(), excelFieldVO);
			}
		}
		return fieldMaps;
	}

}
