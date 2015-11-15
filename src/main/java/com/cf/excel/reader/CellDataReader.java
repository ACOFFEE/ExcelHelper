package com.cf.excel.reader;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.DateUtil;

import com.cf.excel.utils.RegexUtil;

public class CellDataReader {
	private static void checkNumber(String value) {
		if (null == value || value.length() == 0) {
			throw new NullPointerException("value is null!");
		}
		if (!RegexUtil.isNumber(value)) {
			throw new RuntimeException(value + "is not a number!");
		}
	}

	private static String createInteger(String value) {
		checkNumber(value);
		int dotIndex = value.lastIndexOf(".");
		if (dotIndex != -1) {
			value = value.substring(0, dotIndex);
		}
		return value;
	}

	public static int readInt(String value) {
		value = createInteger(value);
		return Integer.parseInt(value);
	}

	public static short readShort(String value) {
		value = value = createInteger(value);
		return Short.parseShort(value);
	}

	public static double readLong(String value) {
		value = createInteger(value);
		return Long.parseLong(value);
	}

	public static float readFloat(String value, int decimal) {
		checkNumber(value);
		BigDecimal bigDecimalValue = new BigDecimal(value);
		return bigDecimalValue.setScale(decimal, BigDecimal.ROUND_HALF_UP)
				.floatValue();
	}

	public static double readDouble(String value, int decimal) {
		checkNumber(value);
		BigDecimal bigDecimalValue = new BigDecimal(value);
		return bigDecimalValue.setScale(decimal, BigDecimal.ROUND_HALF_UP)
				.doubleValue();
	}

	public static BigInteger readBigInteger(String value) {
		value = createInteger(value);
		return new BigInteger(value);
	}

	public static BigDecimal readBigDecimal(String value, int decimal) {
		checkNumber(value);
		return new BigDecimal(value)
				.setScale(decimal, BigDecimal.ROUND_HALF_UP);
	}

	public static Date readDate(String value, String dateFormat)
			throws ParseException {
		Date date = null;
		if (RegexUtil.isNumber(value)) {
			date = DateUtil.getJavaDate(Double.parseDouble(value));
		} else {
			if (null != dateFormat
					&& (dateFormat = dateFormat.trim()).length() != 0) {
				SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
				date = sdf.parse(value);
			}
		}
		return date;
	}
}
