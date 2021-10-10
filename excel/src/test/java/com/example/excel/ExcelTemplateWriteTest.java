package com.example.excel;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

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

import com.example.comm.Excel;
import com.example.comm.User;
import com.example.excel.util.CellUtil;
import com.example.excel.util.ExcelDataCheck;

class ExcelTemplateWriteTest {

	@Test
	void test() {
		String sourceFilePath = "C:\\Users\\Disen\\OneDrive\\桌面\\template.xlsx";

		try {

			List<String> msgs = new ArrayList<String>();
			ExcelWriter excelWriter = ExcelWriter.writeExcel().withAutoClose(false)
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
			//File file = new File("C:\\Users\\Disen\\OneDrive\\桌面\\123456.xlsx");
			//excelWriter.writeOut(new FileOutputStream(file));
			excelWriter.writeOut();
			excelWriter.complete();
			System.err.println(list.size());
			System.err.println(msgs);
			assertTrue(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
