package com.example.excel;

import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.junit.jupiter.api.Test;

import com.example.comm.User;

class ExcelReaderTest {

	private final static Logger logger = Logger.getLogger(ExcelReaderTest.class.getName());
	private final static String excelName = "C:\\Users\\Disen\\OneDrive\\桌面\\template.xlsx";

	@Test
	void testExample1() {

		try {
			ExcelReader er = ExcelReader.build(excelName).withAutoClose(false);
			List<User> list1 = er.registerConverter("age", (Cell cell, User user) -> {
				user.setAge(1000);
			}).registerConverter("professional", (Cell cell, User user) -> {
				user.setProfessional("计算机应用");
			}).doRead(User.class);
			List<User> list2 = er.doRead(User.class);
			er.complete();
			assertNotNull(list1);
			logger.info("test1====>" + list1.toString());
			logger.info("test2====>" + list2.toString());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	void testExample2() {
		// 指定列表名读取
		List<String> fields = new ArrayList<String>();
		fields.add("id");
		fields.add("name");
		fields.add("age");
		try {
			List<User> list = ExcelReader.build(excelName).withColumnField(fields).doRead(2, User.class);
			assertNotNull(list);
			logger.info("Example2====>" + list.toString());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}
