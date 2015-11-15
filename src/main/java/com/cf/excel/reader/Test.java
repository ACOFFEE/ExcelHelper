package com.cf.excel.reader;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import com.cf.excel.ExcelConstants;
import com.cf.excel.PersonVO;
import com.cf.excel.writer.NewExcelWriter;

public class Test {
	public static void main(String[] args) throws Exception {
		// readBean();
		readList();
	}

	/**
	 * 读取结果为实体Bean
	 * 
	 * @Title readBean
	 * @Description TODO
	 * @throws Exception
	 * @author XuZhen
	 * @date 2015年11月13日-上午10:15:48
	 * @update
	 *
	 */
	private static void readBean() throws Exception {
		long startTime = System.currentTimeMillis();
		FileInputStream in = new FileInputStream("E:\\person.xlsx");
		ExcelReader<PersonVO> excelReader = new ExcelReader<PersonVO>(0, 3, 1,
				null, PersonVO.class);
		List<PersonVO> list = excelReader.read(in);
		in.close();
		long endTime = System.currentTimeMillis();
		System.out.println((endTime - startTime) / 1000);
		System.out.println("DataSize:" + list.size());
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		for (int i = 0; i < list.size(); i++) {
			PersonVO person = list.get(i);
			Date birthday = person.getBirthday();
			System.out.println("name:"
					+ person.getName()
					+ ";age:"
					+ person.getAge()
					+ ";income:"
					+ person.getIncome().toString()
					+ ";birthday:"
					+ (null == birthday ? null : sdf.format(person
							.getBirthday())));
		}
		InputStream templateIn = new FileInputStream("E:\\template.xlsx");
		OutputStream out = new FileOutputStream("E:\\result.xlsx");
		NewExcelWriter<PersonVO> excelWriter = new NewExcelWriter<PersonVO>(
				templateIn, out, ExcelConstants.OFFICE_2007, "测试",
				PersonVO.class);
		excelWriter.setStartRow(1);
		excelWriter.push(list);
		excelWriter.flush();
	}

	/**
	 * 去读结果为List
	 * 
	 * @Title readList
	 * @Description TODO
	 * @throws Exception
	 * @author XuZhen
	 * @date 2015年11月13日-上午10:16:16
	 * @update
	 *
	 */
	private static void readList() throws Exception {
		FileInputStream in = new FileInputStream("E:\\person.xlsx");
		List<Map<String, Object>> l = new ArrayList<Map<String, Object>>();
		ExcelReader<List> excelReader = new ExcelReader<List>(0, 3, 1, null,
				List.class);
		List<List> list = excelReader.read(in);
		in.close();
		for (int i = 0; i < list.size(); i++) {
			List<Object> ls = list.get(i);
			for (int j = 0; j < ls.size(); j++) {
				System.out.println(ls.get(j));
			}

		}
		InputStream templateIn = new FileInputStream("E:\\template.xlsx");
		OutputStream out = new FileOutputStream("E:\\result.xlsx");
		NewExcelWriter<List> excelWriter = new NewExcelWriter<List>(templateIn,
				out, ExcelConstants.OFFICE_2007, "测试", List.class) {
			@Override
			public List<Object> changeDate(List object) {
				return object;
			}
		};
		excelWriter.setStartRow(1);
		excelWriter.push(list);
		excelWriter.flush();
	}
}
