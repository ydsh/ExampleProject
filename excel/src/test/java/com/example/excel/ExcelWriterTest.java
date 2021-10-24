package com.example.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.junit.jupiter.api.Test;

import com.example.comm.User;

class ExcelWriterTest {

	@Test
	void test() {
		List<User> list = new ArrayList<User>();
		for(int i=0;i<10;i++) {
			User user = new User();
			user.setAge(25);
			user.setDegree("博士");
			user.setEmail("123456@163.com");
			user.setGraduateSchool("华中科技大学");
			user.setGraduateTime(new Date());
			user.setIntro("积极能干");
			user.setId("100"+i);
			user.setName("小"+i);
			user.setLocation("科莱亚");
			user.setProfessional("计算机信息");
			list.add(user);
		}
		try {
			List<String> fields = new ArrayList<String>();
			fields.add("id");
			fields.add("name");
			fields.add("age");
			ExcelWriter.build().withColumnField(fields).doWrite(list).writeOut("C:\\Users\\Disen\\OneDrive\\桌面\\write_data.xlsx");
			assertTrue(true);
		} catch (Exception e) {
			
			e.printStackTrace();
		}
	}

}
