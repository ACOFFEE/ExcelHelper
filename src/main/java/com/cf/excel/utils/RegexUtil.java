package com.cf.excel.utils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RegexUtil {
	public static boolean isNumber(String src) {
		String reg = "^[-+]?(([0-9]+)([.]([0-9]+))?|([.]([0-9]+))?)$";
		Pattern pattern = Pattern.compile(reg);
		Matcher matcher = pattern.matcher(src);
		return matcher.matches();
	}

}
