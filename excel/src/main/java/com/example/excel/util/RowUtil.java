package com.example.excel.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import com.example.comm.Excel;

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
				row.getCell(colIndex).setCellStyle(cellStyle);
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
	 * 将数据写入到excel表行中
	 * 
	 * @param <T>
	 * @param row
	 * @param columnMap
	 * @param data
	 * @throws Exception
	 */
	public static <T> void dataToRow(Row row, Map<Integer, String> columnMap, T data) throws Exception {
		@SuppressWarnings("unchecked")
		Class<T> clazz = (Class<T>) data.getClass();
		// excel表格数据列与对象属性映射
		// Map<Integer, String> columnMap = CellUtil.columnFieldMap(clazz);
		// 将数据据写入单元格
		for (Integer c : columnMap.keySet()) {
			Cell cell = getCell(row, c);
			String fieldName = columnMap.get(c);
			String getterName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
			Method[] methods = clazz.getDeclaredMethods();
			for (int i = 0, len = methods.length; i < len; i++) {
				if (getterName.equals(methods[i].getName())) {
					// 数据写入单元格
					CellUtil.setCellValue(cell, methods[i].invoke(data));
				}
			}
		}
	}
	/**
	 * map数据写入到excel表行中
	 * @param row
	 * @param columnMap
	 * @param mapData
	 * @throws Exception
	 */
	public static void mapDataToRow(Row row, Map<Integer, String> columnMap, Map<String,Object> mapData) throws Exception{
		for(String key : mapData.keySet()) {
			int colIndex = -1;
			if(columnMap.containsValue(key)) {
				for(Map.Entry<Integer, String> entry:columnMap.entrySet()) {
					if(key.equals(entry.getValue())) {
						colIndex = entry.getKey();
					}
				}
			}else {
				colIndex = columnMap.size();
				columnMap.put(colIndex, key);
			}
			
			Cell cell = getCell(row, colIndex);
			CellUtil.setCellValue(cell, mapData.get(key));
		}
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
