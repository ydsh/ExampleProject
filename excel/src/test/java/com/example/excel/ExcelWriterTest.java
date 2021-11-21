package com.example.excel;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com.example.comm.User;

class ExcelWriterTest {
	private List<User> list = new ArrayList<User>();
    @BeforeEach
	 void init() {
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
    }
	@Test
	void testWriteTemplate1() {
		try {
			//build指定写数据的模板，并创建workbook实例
			ExcelWriter excelWriter = ExcelWriter.build("C:\\Users\\Disen\\OneDrive\\桌面\\tmp.xlsx");
			Workbook workbook = excelWriter.getWorkbook();
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setDataFormat(workbook.createDataFormat().getFormat("#,#00.00000"));
			//注册转换器写
			excelWriter.registerConverter("age", (Cell cell, User user)->{
				cell.setCellStyle(cellStyle);
				cell.setCellValue(user.getAge());
			}).doWriteTemplate(null,1,list).writeOut("C:\\Users\\Disen\\OneDrive\\桌面\\write_template1.xlsx");
			assertTrue(true);
		} catch (Exception e) {
			
			e.printStackTrace();
		}
	}
	@Test
	void testWriteTemplate2() {
		try {
			//build指定写数据的模板，并创建workbook实例
			ExcelWriter.build("C:\\Users\\Disen\\OneDrive\\桌面\\tmp.xlsx")
			//注册写入模板的字段
			.registerColIndexFieldMap(Arrays.asList(new String[]{"id","name","age"}))
			//1指定开始写的行，list指定写入的数据
			.doWriteTemplate(null,1,list)
			.writeOut("C:\\Users\\Disen\\OneDrive\\桌面\\write_template2.xlsx");
			assertTrue(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	@Test
	void testWrite1() {
		List<List<String>> cols = new ArrayList<List<String>>();
		List<String> col1 = new ArrayList<String>();
		col1.add("基本信息");
		col1.add("基本信息");
		col1.add("ID");
		List<String> col2 = new ArrayList<String>();
		col2.add("基本信息");
		col2.add("基本信息");
		col2.add("姓名");
		List<String> col3 = new ArrayList<String>();
		col3.add("基本信息");
		col3.add("基本信息");
		col3.add("年龄");
		cols.add(col1);
		cols.add(col2);
		cols.add(col3);
		try {
			ExcelWriter.build().registerColumnFieldList(Arrays.asList(new String[]{"id","name","age"}), cols)
			.doWrite(list).writeOut("C:\\Users\\Disen\\OneDrive\\桌面\\write1.xlsx");
			assertTrue(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}





