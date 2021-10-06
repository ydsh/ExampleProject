package com.example.comm;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.excel.ExcelReader;

public class ExcelUtils {
	private static Logger logger = Logger.getLogger(ExcelReader.class.getName());
	public static final String XLS = "xls";
	public static final String XLSX = "xlsx";

	/**
	 * 创建默认工作簿
	 * 
	 * @return
	 */
	public static Workbook createWorkbook() {
		return createWorkbook(XLSX);
	}

	/**
	 * 创建工作簿
	 * 
	 * @param fileType
	 * @return
	 */
	public static Workbook createWorkbook(String fileType) {
		Workbook workbook = null;
		if (XLS.equalsIgnoreCase(fileType)) {
			workbook = new HSSFWorkbook();
			logger.info("已创建`" + XLS + "`类型工作簿。");
		} else {
			// 默认创建xlsx工作簿
			// workbook = new XSSFWorkbook();
			logger.info("已创建`" + XLSX + "`类型工作簿。");
			workbook = new SXSSFWorkbook();
		}
		return workbook;
	}

	/**
	 * 使用输入流创建工作簿对象，并返回工作簿对象。 文件类型不同创建工作簿的方法不同。
	 * 
	 * @param inputStream
	 * @param fileType
	 * @return
	 * @throws IOException
	 */
	public static Workbook createWorkbook(InputStream inputStream, String fileType) throws IOException {
		Workbook workbook = null;
		if (XLS.equalsIgnoreCase(fileType)) {
			workbook = new HSSFWorkbook(inputStream);
			logger.info("已创建`" + XLS + "`类型工作簿。");
		} else if (XLSX.equalsIgnoreCase(fileType)) {
			workbook = new XSSFWorkbook(inputStream);
			logger.info("已创建`" + XLSX + "`类型工作簿。");
		}

		return workbook;
	}

	/**
	 * 创建sheet表格
	 * 
	 * @param workbook
	 * @return
	 */
	public static Sheet createSheet(Workbook workbook) {
		if (workbook != null) {

			return workbook.createSheet();

		}
		return null;
	}

	/**
	 * 创建自定义名字的sheet表格
	 * 
	 * @param workbook
	 * @param sheetName
	 * @return
	 */
	public static Sheet createSheet(Workbook workbook, String sheetName) {
		if (workbook != null) {

			return workbook.createSheet(sheetName);
		}
		return null;
	}

	/**
	 * 设置表头数据
	 * 
	 * @param sheet
	 * @param headColumnMap
	 */
	public static void setHeadRow(Sheet sheet, Map<Integer, String> headColumnMap) {
		if (headColumnMap != null) {
			Set<Integer> columns = headColumnMap.keySet();
			// 设置每列宽度
			for (Integer c : columns) {
				sheet.setColumnWidth(c, 3000);
			}
			// 设置行高
			sheet.setDefaultRowHeight(Short.valueOf("300"));
			// 表头单元格样式
			CellStyle headCellStyle = setHeadCellStyle(sheet.getWorkbook());
			// 创建表投行
			Row row = sheet.createRow(0);
			for (Integer c : columns) {
				Cell cell = row.createCell(c);
				;
				// 设置列表名
				cell.setCellValue(headColumnMap.get(c));
				// 添加样式
				cell.setCellStyle(headCellStyle);
			}
		}

	}

	/**
	 * 根据对象属性设置表格对应数据列的格式
	 * setDefaultColumnStyle这个只对HSSFWorkbook有效
	 * @param workbook
	 * @param columnMap
	 */
	public static void setColumnStyle(Sheet sheet, Map<Integer, String> columnMap, Class<?> clazz) throws Exception {
		for (Integer c : columnMap.keySet()) {
			CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
			;

			String fieldName = columnMap.get(c);
			Field field = clazz.getDeclaredField(fieldName);
			if (field.isAnnotationPresent(Excel.class)) {
				Excel ex = field.getAnnotation(Excel.class);
				if (!"".equals(ex.fmt())) {
					DataFormat format = sheet.getWorkbook().createDataFormat();
					cellStyle.setDataFormat(format.getFormat(ex.fmt()));

				} else {
					// 根据对象属性类型设置数据列的格式
					columnStyle(cellStyle, field.getType());
				}
			}
			sheet.setDefaultColumnStyle(c, cellStyle);
			// logger.info(sheet.getColumnStyle(c).getDataFormatString());
		}
	}

	/**
	 * 根据对象属性类型设置数据列的格式 注意：此方法作为默认设置
	 * 
	 * @param style
	 * @param clazz
	 */
	private static void columnStyle(CellStyle style, Class<?> clazz) {
		switch (clazz.getSimpleName()) {
		case "Integer":
			;
		case "int":
			;
		case "Long":
			;
		case "long":
			;
		case "Short": 
			;
		case "short":
			;
		case "Byte":
			;
		case "byte": {
			// 整形数没有小数点
			style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
		}
		case "Double":
			;
		case "double":
			;
		case "Float":
			;
		case "float":
			;
		case "BigDecimal": {
			// 浮点数设置6位小数
			style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.000000"));
			break;
		}
		case "Date": {
			style.setDataFormat(HSSFDataFormat.getBuiltinFormat("yyyy-mm-dd"));
			break;
		}
		case "String":
			;
		default:
			style.setDataFormat(HSSFDataFormat.getBuiltinFormat("@"));
			break;
		}
	}

	/**
	 * excel表格表头列和名字映射
	 * 
	 * @param headRow
	 * @return
	 */
	public static Map<Integer, String> headColumnMap(Class<?> clazz) {
		Map<Integer, String> columnMap = new HashMap<Integer, String>();
		Field[] fileds = clazz.getDeclaredFields();
		for (Field field : fileds) {
			// 判断对象属性是否存在Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				// 将列和字段名字放入map中
				if (ex.order() != -1) {
					// 列和数据对象属性映射
					columnMap.put(ex.order(), ex.name()[0]);
				}
			}
		}
		return columnMap;
	}

	/**
	 * 设置表头单元格样式
	 * 
	 * @param workbook
	 * @return
	 */
	private static CellStyle setHeadCellStyle(Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		;
		// 文字水平居中
		style.setAlignment(HorizontalAlignment.CENTER);
		// 文字垂直居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		// 设置边框样式
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		// 设置边框颜色
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		// 设置背景色
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		// 字体样式
		Font font = workbook.createFont();
		// 粗体
		font.setBold(true);
		// 字体族
		font.setFontName("宋体");
		// 是否粗体
		font.setBold(true);
		// 字体高度
		font.setFontHeight(Short.valueOf("12"));
		// 字号
		font.setFontHeightInPoints(Short.valueOf("12"));

		style.setFont(font);
		return style;
	}

	/**
	 * excel表格列和数据对象属性映射
	 * 
	 * @param headRow
	 * @return
	 */
	public static Map<Integer, String> columnMap(Class<?> clazz) {
		Map<Integer, String> columnMap = new HashMap<Integer, String>();
		Field[] fileds = clazz.getDeclaredFields();
		for (Field field : fileds) {
			// 判断对象属性是否存在Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				// 将列和字段名字放入map中
				if (ex.order() != -1) {
					// 列和数据对象属性映射
					columnMap.put(ex.order(), field.getName());
				}
			}
		}
		return columnMap;
	}

	/**
	 * excel表格列和数据对象属性映射
	 * 
	 * @param headRow
	 * @return
	 */
	public static Map<Integer, String> columnMap(Row headRow, Class<?> clazz) {
		Map<Integer, String> columnMap = new HashMap<Integer, String>();
		Field[] fileds = clazz.getDeclaredFields();
		int columns = headRow.getPhysicalNumberOfCells();
		for (int i = 0; i < columns; i++) {
			Object cellValue = cellValue(headRow.getCell(i));
			for (Field field : fileds) {

				// 判断对象属性是否存在Excel注解
				if (field.isAnnotationPresent(Excel.class)) {
					// 获取Excel注解
					Excel ex = field.getAnnotation(Excel.class);
					// 将列和字段名字放入map中
					if (ex.order() != -1) {
						columnMap.put(ex.order(), field.getName());
					} else if (cellValue != null && cellValue.toString().equals(ex.name())) {// 列名字与字段注解名字一样
						columnMap.put(i, field.getName());
					}
				}
			}
		}
		return columnMap;
	}

	/**
	 * 获取单元格的数据
	 * 
	 * @param cell
	 * @return
	 */
	public static Object cellValue(Cell cell) {
		if (cell == null) {
			return null;
		}
		Object obj = null;
		switch (cell.getCellType()) {
		case NUMERIC: {
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				// 日期
				obj = cell.getDateCellValue();
			} else {
				// 数字转为BigDecimal类型，并保留6位小数
				obj = new BigDecimal(cell.getNumericCellValue()).setScale(6, RoundingMode.HALF_UP);
			}
			break;
		}
		case STRING: {
			// 字符串
			obj = cell.getStringCellValue();
			break;
		}
		case BOOLEAN: {
			// bool
			obj = cell.getBooleanCellValue();
			break;
		}
		case FORMULA: {
			try {
				// obj = cell.getCellFormula(); 公式转为获取值
				obj = new BigDecimal(cell.getNumericCellValue()).setScale(6, RoundingMode.HALF_UP);
			} catch (IllegalStateException e) {
				obj = cell.getRichStringCellValue();
			}
			break;
		}
		case BLANK:
			;
		case ERROR:
			;
		default:
			;
		}
		return obj;
	}

	/**
	 * BigDecimal类型转其他数字类型
	 * 
	 * @param bigDecimal
	 * @param toType
	 * @return
	 */
	public static Object bigDecimalToNum(BigDecimal bigDecimal, String toType) {
		Object obj = null;
		switch (toType) {
		case "Integer":
			;
		case "int": {
			obj = bigDecimal.intValue();
			break;
		}
		case "Double":
			;
		case "double": {
			obj = bigDecimal.doubleValue();
			break;
		}
		case "Float":
			;
		case "float": {
			obj = bigDecimal.floatValue();
			break;
		}
		case "Long":
			;
		case "long": {
			obj = bigDecimal.longValue();
			break;
		}
		case "Short":
			;
		case "short": {
			obj = bigDecimal.shortValue();
			break;
		}
		case "Byte":
			;
		case "byte": {
			obj = bigDecimal.byteValue();
			break;
		}
		default:
			obj = bigDecimal;
			break;
		}
		return obj;
	}
}
