package com.example.excel;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.junit.jupiter.api.Test;

import com.example.comm.User;
import com.example.excel.util.CellUtil;
import com.example.excel.util.Excel;
import com.example.excel.util.ExcelDataCheck;

public class ExceTest {

	@Test
	public void excelReadTest() {
		String fileName = "C:\\Users\\Disen\\OneDrive\\桌面\\template.xlsx";
		try {
			List<User> list = ExcelReader.readExcel(fileName).build().doRead(User.class);
			System.err.println(list);
			assertTrue(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void excelWriteTest() {
		String fileName = "C:\\Users\\Disen\\OneDrive\\桌面\\template1234.xlsx";
		List<User> dataList = new ArrayList<User>(100);
		User user = new User();
		user.setId("1101");
		user.setName("小明");
		user.setAge(25);
		user.setEmail("123456@123.com");
		user.setDegree("本科");
		user.setGraduateSchool("华中科技大学");
		user.setGraduateTime(new Date());
		user.setJob("工程师");
		user.setLocation("华中33号");
		user.setProfessional("计算机科学");
		user.setReference("柏林");
		user.setIntro("积极向上");
		user.setTime(new Date());
		List<User> list = new ArrayList<User>();
		list.add(user);
		int n = list.size();
		for (int i = 0; i < 100; i++) {
			if (i < n) {
				dataList.add(list.get(i));
			} else {
				dataList.add(list.get(i % n));
			}
		}
		try {
//			 ExcelWriter.writeExcel(fileName).build("C:\\Users\\Disen\\OneDrive\\桌面\\userTemplate.xlsx").doWriteTemplate("Sheet1",
//			 dataList);
//			ExcelWriter.writeExcel(fileName).build("C:\\Users\\Disen\\OneDrive\\桌面\\userTemplate.xlsx").withAutoClose(false)
//			.doWriteTemplate("新1", dataList)
//			.doWriteTemplate("新2", dataList)
//			.doWriteTemplate("新3", dataList)
//			.writeOut()
//			.complete();
			ExcelWriter.writeExcel(fileName).build().withAutoClose(false).doWrite(dataList).doWrite("A22", dataList)
					.writeOut().complete();
			assertTrue(true, "数据写入完毕");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void templateWrite() {
		String sourceFilePath = "C:\\Users\\Disen\\OneDrive\\桌面\\template.xlsx";

		try {

			List<String> msgs = new ArrayList<String>();
			ExcelWriter excelWriter = ExcelWriter.writeExcel("C:\\Users\\Disen\\OneDrive\\桌面\\templateWrite.xlsx")
					.build(sourceFilePath);
			Workbook workbook = excelWriter.getWorkbook();
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			System.err.println("Number of " + workbook.getNumberOfSheets());
			Sheet sheet = workbook.getSheetAt(0);
			ClientAnchor clientAnchor = new XSSFClientAnchor(0, 0, 0, 0, 3, 3, 5, 6);
			ExcelDataCheck<User> userCheck = (int rowIndex, User user) -> {
				boolean result = true;
				try {
					if (user.getAge() > 100 || user.getAge() <= 0) {
						Field field = User.class.getDeclaredField("age");
						if (field.getAnnotation(Excel.class) != null) {
							Excel ex = field.getAnnotation(Excel.class);
							int colIndex = ex.order();
							if (sheet.getRow(rowIndex) != null) {
								Row row = sheet.getRow(rowIndex);
								Cell cell = row.getCell(colIndex);
								cell.setCellStyle(cellStyle);
								CellUtil.setCellComment(cell, clientAnchor, "提示", "用户年龄非法，请输入正确年龄。");
							}
						}
						result &= false;
					}
					if (user.getIntro() == null || user.getIntro().isBlank() || user.getIntro().length() < 200) {
						Field field = User.class.getDeclaredField("intro");
						if (field.getAnnotation(Excel.class) != null) {
							Excel ex = field.getAnnotation(Excel.class);
							int colIndex = ex.order();
							if (sheet.getRow(rowIndex) != null) {
								Row row = sheet.getRow(rowIndex);
								Cell cell = row.getCell(colIndex);
								cell.setCellStyle(cellStyle);
								CellUtil.setCellComment(cell, clientAnchor, "提示", "简介不能为空，或者长度小于200个字符");
							}
						}
						result &= false;
					}
				} catch (NoSuchFieldException e) {
					e.printStackTrace();
				} catch (SecurityException e) {
					e.printStackTrace();
				}
				return result;
			};
			List<User> list = ExcelReader.readExcel(sourceFilePath).build().doReadCheck(User.class, userCheck);
			excelWriter.writeOut();
			excelWriter.complete();
			System.err.println(list.size());
			System.err.println(msgs);
			assertTrue(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void writeMapDataTest() {
		try {
			List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();

			for (int i = 0; i < 10; i++) {
				Map<String, Object> map = new HashedMap<String, Object>();
				map.put("name", "哈哈哈凝视对方");
				map.put("email", "1123748@123.com");
				list.add(map);
			}
			//int len ="abc123哈哈哈凝视对方".length();
			//System.err.println("长度："+len);
			ExcelWriter.writeExcel("C:\\Users\\Disen\\OneDrive\\桌面\\mapData.xlsx").build().mapDataToSheet("mapData", list);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
