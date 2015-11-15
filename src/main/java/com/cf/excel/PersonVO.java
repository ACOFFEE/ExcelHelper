package com.cf.excel;

import java.math.BigDecimal;
import java.util.Date;

import com.cf.excel.annotation.ExcelField;
import com.cf.excel.annotation.ExcelObject;

@ExcelObject(true)
public class PersonVO {
	@ExcelField(index = "1")
	private String name;
	@ExcelField(index = "2")
	private int age;
	@ExcelField(index = "4", decimal = "3")
	private BigDecimal income;

	@ExcelField(index = "3", dateFormat = "yyyy年MM月dd")
	private Date birthday;

	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public BigDecimal getIncome() {
		return income;
	}

	public void setIncome(BigDecimal income) {
		this.income = income;
	}

}
