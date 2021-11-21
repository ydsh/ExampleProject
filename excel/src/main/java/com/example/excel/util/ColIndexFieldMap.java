package com.example.excel.util;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;

/**
 * 列索引和字段关系映射
 *
 */
public class ColIndexFieldMap {
	private Map<Integer, String> columnFieldMap;

	private ColIndexFieldMap() {
	}

	public static ColIndexFieldMap build() {
		return FuncUtil.create(ColIndexFieldMap::new);
	}

	/**
	 * 字段数组转ColumnFieldMap
	 * 
	 * @param list
	 */
	public Map<Integer, String> withColnumFieldMap(String[] array) {
		if (columnFieldMap == null) {
			columnFieldMap = new HashMap<Integer, String>();
		}
		columnFieldMap.clear();
		for (int i = 0, len = array.length; i < len; i++) {
			columnFieldMap.put(i, array[i]);
		}
		return columnFieldMap;
	}

	/**
	 * 列索引和字段映射
	 * 
	 * @param map
	 * @return
	 */
	public Map<Integer, String> withColnumFieldMap(Map<Integer, String> map) {
		columnFieldMap = map;
		return columnFieldMap;
	}

	/**
	 * 字段的信息列表转ColumnFieldMap
	 * 
	 * @param list
	 */
	public <T> Map<Integer, String> withColnumFieldMap(List<T> list) {
		if (columnFieldMap == null) {
			columnFieldMap = new HashMap<Integer, String>();
		}
		columnFieldMap.clear();
		for (int i = 0, len = list.size(); i < len; i++) {

			if (list.get(i) instanceof String) {
				String data = (String) list.get(i);
				columnFieldMap.put(i, data);
			}
			if (list.get(i) instanceof ColumnField) {
				ColumnField data = (ColumnField) list.get(i);
				if (data.getColIndex() > -1 && data.getFieldName() != null) {
					columnFieldMap.put(data.getColIndex(), data.getFieldName());
				}
			}
		}

		return columnFieldMap;
	}

	/**
	 * 字段所在excel表的列索引
	 * 
	 * @param fieldName
	 * @return
	 */
	public int columnIndexOfField(String fieldName) {
		if (fieldName == null || "".equals(fieldName)) {
			return -1;
		}
		Optional<Entry<Integer, String>> optional = columnFieldMap.entrySet().stream().filter(data -> {
			return fieldName.equals(data.getValue());
		}).findFirst();

		if (optional.isPresent()) {
			return optional.get().getKey();
		}
		return -1;
	}

	/**
	 * 字段所在excel表的列索引
	 * 
	 * @param <T>
	 * @param fieldName
	 * @param clazz
	 * @return
	 */
	public <T> int columnIndexOfField(String fieldName, Class<T> clazz) {
		if (fieldName == null || "".equals(fieldName) || clazz == null) {
			return -1;
		}
		try {
			Field field = clazz.getDeclaredField(fieldName);
			field.setAccessible(true);
			if (field.isAnnotationPresent(Excel.class)) {
				Excel excel = field.getAnnotation(Excel.class);
				return excel.order();
			}
		} catch (NoSuchFieldException | SecurityException e) {
			e.printStackTrace();
		}
		return -1;
	}

	public Map<Integer, String> getColumnFieldMap() {
		return columnFieldMap;
	}

	public boolean isEmpty() {
		return this.columnFieldMap.isEmpty();
	}

	public void clear() {
		this.columnFieldMap.clear();
		;
	}

}
