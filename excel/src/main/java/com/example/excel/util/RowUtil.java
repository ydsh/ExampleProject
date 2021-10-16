package com.example.excel.util;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class RowUtil {
	//private static final Logger logger = Logger.getLogger(RowUtil.class.getName());

	private RowUtil() {
	}

	/**
	 * 使用注解给定的格式格式化行单元格
	 * @param <T>
	 * @param row
	 * @param fmtMap 存储注解格式化对应的CellStyle
	 * @param columnMap
	 * @param clazz
	 */
	public static <T> void setRowCellFormat(Row row, Map<String, CellStyle> fmtMap, Map<Integer, String> columnMap,
			Class<?> clazz) {
		Map<String, String> map = getFieldFmtMap(clazz);
		for (String fieldName : map.keySet()) {
			if (map.get(fieldName) == null || "".equals(map.get(fieldName))) {
				continue;
			}
			CellStyle cellStyle = fmtMap.get(map.get(fieldName));
			int colIndex = -1;
			for (Integer key : columnMap.keySet()) {
				if (fieldName.equals(columnMap.get(key))) {
					colIndex = key;
					break;
				}
			}
			if (colIndex != -1) {
				getCell(row, colIndex).setCellStyle(cellStyle);
			}
		}
	}

	/**
	 * 获取注解对应列的字段和格式化字符串的映射
	 * 
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	public static <T> Map<String, String> getFieldFmtMap(Class<T> clazz) {
		Map<String, String> result = new HashMap<String, String>();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			field.setAccessible(true);
			// 判断字段是否标注Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				if (ex.order() > -1 && !"".equals(ex.fmt())) {
					result.put(field.getName(), ex.fmt());
				}
			}
		}
		return result;
	}

	/**
	 * 单元格村直接获取否则新创建一个
	 * 
	 * @param row
	 * @param columnIndex
	 * @return
	 */
	public static Cell getCell(Row row, int columnIndex) {
		Cell cell = row.getCell(columnIndex);
		if (cell == null) {
			cell = row.createCell(columnIndex);
		}
		return cell;
	}
}
