package com.example.excel.util;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Calendar;
import java.util.Date;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import com.example.comm.Excel;

public class CellUtil {
	private static final Logger logger = Logger.getLogger(CellUtil.class.getName());

	private CellUtil() {
	}

	/**
	 * 数值转换并写入到Excel单元格中
	 * 
	 * @param cell
	 * @param obj
	 */
	public static void setCellValue(Cell cell, Object obj) {
		if (obj != null) {
			logger.config("数据类型转换并写入单元格");
			if (obj instanceof Date) {
				cell.setCellValue((Date) obj);
			} else if (obj instanceof Calendar) {
				cell.setCellValue((Calendar) obj);
			} else if (obj instanceof Boolean) {
				cell.setCellValue((Boolean) obj);
			} else if (obj instanceof Integer || obj instanceof Byte || obj instanceof Character || obj instanceof Short
					|| obj instanceof Double || obj instanceof Float) {
				cell.setCellValue(Double.valueOf(String.valueOf(obj)));
			} else if (obj instanceof BigDecimal) {
				BigDecimal bg = (BigDecimal) obj;
				cell.setCellValue(bg.doubleValue());
			} else {
				cell.setCellValue(String.valueOf(obj));
			}
		}
	}

	/**
	 * 单元格的值获取
	 * 
	 * @param cell
	 * @return
	 */
	public static Object getCellValue(Cell cell) {
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
				obj = new BigDecimal(cell.getNumericCellValue()).setScale(12, RoundingMode.HALF_UP);
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
				obj = new BigDecimal(cell.getNumericCellValue()).setScale(12, RoundingMode.HALF_UP);
			} catch (IllegalStateException e) {
				obj = cell.getRichStringCellValue();
				logger.info("公式转换出错");
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
		if (bigDecimal == null) {
			bigDecimal = new BigDecimal("0");
		}
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
	/**
	 * 获取字段所在列序号
	 * @param <T>
	 * @param fieldName
	 * @param clazz
	 * @return
	 */
    public static <T> int columnIndex(String fieldName,Class<T> clazz) {
    	int result = -1;
    	try {
			Field field = clazz.getDeclaredField(fieldName);
			if(field.isAnnotationPresent(Excel.class)) {
				Excel ex = field.getAnnotation(Excel.class);
				result = ex.order();
			}
		} catch (NoSuchFieldException | SecurityException e) {
			logger.warning("无法获取该字段所在单元格列序号");
			e.printStackTrace();
		}
    	return result;
    }
    /**
     * 获取字段所在列的名字，如果是复杂表头则返回最后一行表头列的名字。
     * @param <T>
     * @param fieldName
     * @param clazz
     * @return
     */
    public static <T> String columnName(String fieldName,Class<T> clazz) {
    	String result = "";
    	try {
    		Field field = clazz.getDeclaredField(fieldName);
			if(field.isAnnotationPresent(Excel.class)) {
				Excel ex = field.getAnnotation(Excel.class);
				String[] names = ex.name();
				if(names!=null&&names.length > 0) {
					result = names[names.length-1];
				}
				
			}
		} catch (Exception e) {
			logger.warning("无法获取该字段所在列的名字");
			e.printStackTrace();
		}
    	return result;
    }
	/**
	 * 设置单元格批注
	 * 
	 * @param cell
	 * @param comment
	 */
	public static void setCellComment(Cell cell, ClientAnchor clientAnchor, String author, String value) {
		Sheet sheet = cell.getSheet();
		Drawing<?> drawing = sheet.getDrawingPatriarch();
		if (drawing == null) {
			drawing = sheet.createDrawingPatriarch();
		}
		Comment comment = drawing.createCellComment(clientAnchor);
		comment.setAuthor(author);
		comment.setString(new XSSFRichTextString(value));
		cell.setCellComment(comment);
	}
}
