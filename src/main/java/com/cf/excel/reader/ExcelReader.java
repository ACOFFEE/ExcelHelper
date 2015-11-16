package com.cf.excel.reader;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.cf.excel.ExcelFieldInfo;
import com.cf.excel.annotation.ExcelObject;
import com.cf.excel.utils.AnnotationFieldUtil;
import com.cf.excel.utils.WorkBookUtil;

public class ExcelReader<T> {
	private Integer startColumn;
	private Integer endColumn;
	private Integer startRow;
	private Integer endRow;

	private Class<T> clazz;

	private Map<String, ExcelFieldInfo> fieldMaps;

	public void setStartColumn(Integer startColumn) {
		this.startColumn = startColumn;
	}

	public void setEndColumn(Integer endColumn) {
		this.endColumn = endColumn;
	}

	public void setStartRow(Integer startRow) {
		this.startRow = startRow;
	}

	public void setEndRow(Integer endRow) {
		this.endRow = endRow;
	}

	public ExcelReader(Integer startColumn, Integer endColumn,
			Integer startRow, Integer endRow, Class<T> clazz) {
		this.startColumn = startColumn;
		this.endColumn = endColumn;
		this.startRow = startRow;
		this.endRow = endRow;
		this.clazz = clazz;
		this.fieldMaps = AnnotationFieldUtil.fieldMaps(this.clazz);
	}

	public ExcelReader(Integer startColumn, Integer endColumn,
			Integer startRow, Class<T> clazz) {
		this(startColumn, endColumn, startRow, null, clazz);
	}

	public ExcelReader(Integer startColumn, Integer endColumn, Class<T> clazz) {
		this(startColumn, endColumn, null, null, clazz);
	}

	public List<T> read(InputStream in) throws Exception {
		Workbook hssfWorkbook = WorkBookUtil.createWorkBook(in);
		List<T> list = new ArrayList<T>();
		if (null == startColumn) {
			throw new NullPointerException("Please set startColumn！");
		}
		if (null == endColumn) {
			throw new NullPointerException("Please set endColumn！");
		}
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			Sheet sheet = hssfWorkbook.getSheetAt(numSheet);
			if (sheet == null) {
				continue;
			}
			int lastRowNum = null != endRow ? endRow : sheet.getLastRowNum();
			int rowNum = null != startRow ? startRow : 0;
			if (clazz == List.class) {
				readToList(sheet, list, lastRowNum, rowNum);
			} else {
				readToBean(sheet, list, lastRowNum, rowNum);
			}
		}
		return list;
	}

	private void readToList(Sheet sheet, List<T> list, int lastRowNum,
			int rowNum) throws Exception {
		for (; rowNum <= lastRowNum; rowNum++) {
			Row hssfRow = sheet.getRow(rowNum);
			if (null == hssfRow) {
				break;
			}
			List<Object> inList = new ArrayList<Object>();
			for (int column = startColumn; column <= endColumn; column++) {
				Cell cell = hssfRow.getCell(column);
				String str = getValue(cell);
				inList.add(str);
			}
			list.add((T) inList);

		}
	}

	private void readToBean(Sheet sheet, List<T> list, int lastRowNum,
			int rowNum) throws Exception {
		ExcelObject excelObject = clazz.getAnnotation(ExcelObject.class);
		if (null == excelObject || !excelObject.value()) {
			return;
		}
		for (; rowNum <= lastRowNum; rowNum++) {
			Row hssfRow = sheet.getRow(rowNum);
			if (null == hssfRow) {
				break;
			}
			T obj = (T) Class.forName(clazz.getName()).newInstance();
			int excelColumn = 1;
			for (int column = startColumn; column <= endColumn; column++) {
				Cell cell = hssfRow.getCell(column);
				ExcelFieldInfo excelField = fieldMaps.get(excelColumn++ + "");
				String strValue = getValue(cell);
				if (null == strValue) {
					continue;
				}
				Field field = excelField.getField();
				Class<?> fieldClazz = field.getType();
				if (fieldClazz == String.class) {
					field.set(obj, strValue);
				} else if (fieldClazz == Integer.class
						|| fieldClazz == int.class) {
					field.set(obj, CellDataReader.readInt(strValue));
				} else if (fieldClazz == short.class
						|| fieldClazz == Short.class) {
					field.set(obj, CellDataReader.readShort(strValue));
				} else if (fieldClazz == long.class || fieldClazz == Long.class) {
					field.set(obj, CellDataReader.readLong(strValue));
				} else if (fieldClazz == float.class
						|| fieldClazz == Float.class) {
					int decimal = Integer.parseInt(excelField.getExcelField()
							.decimal());
					field.set(obj, CellDataReader.readFloat(strValue, decimal));
				} else if (fieldClazz == double.class
						|| fieldClazz == Double.class) {
					int decimal = Integer.parseInt(excelField.getExcelField()
							.decimal());
					field.set(obj, CellDataReader.readDouble(strValue, decimal));
				} else if (fieldClazz == BigInteger.class) {
					field.set(obj, CellDataReader.readBigInteger(strValue));
				} else if (fieldClazz == BigDecimal.class) {
					int decimal = Integer.parseInt(excelField.getExcelField()
							.decimal());
					field.set(obj,
							CellDataReader.readBigDecimal(strValue, decimal));
				} else if (fieldClazz == Date.class) {
					String dateFormat = excelField.getExcelField().dateFormat();
					field.set(obj,
							CellDataReader.readDate(strValue, dateFormat));
				}
			}
			list.add(obj);
		}
	}

	private String getValue(Cell cell) {
		if (null == cell) {
			return null;
		}
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(cell.getBooleanCellValue());
		} else if (cellType == Cell.CELL_TYPE_NUMERIC) {
			double doubleValue = cell.getNumericCellValue();
			BigDecimal tempNumber = new BigDecimal(doubleValue);
			String strValue = tempNumber.toString();
			String[] strArray = strValue.split("\\.");
			DecimalFormat df = null;
			if (strArray.length == 2) {
				StringBuilder strb = new StringBuilder();
				strb.append("0");
				for (int i = 0; i < strArray[1].length(); i++) {
					if (i == 0) {
						strb.append(".");
					}
					strb.append("0");
				}
				df = new DecimalFormat(strb.toString());
			} else {
				df = new DecimalFormat("0");
			}
			return df.format(doubleValue);
		} else {
			return String.valueOf(cell.getStringCellValue());
		}
	}
}
