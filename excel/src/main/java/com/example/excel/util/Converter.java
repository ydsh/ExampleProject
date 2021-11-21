package com.example.excel.util;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 转换器，将取单元格的值转换成java数据
 * 
 * @param <T>
 */
@FunctionalInterface
public interface Converter<T extends Cell, K> {
	/**
	 * 默认读cell方法
	 * 
	 * @param cell
	 * @return
	 */
	default K defaultConvert(Cell cell) {
		return (K) CellUtil.getCellValue(cell);
	}

	/**
	 * 默认写cell方法
	 * 
	 * @param cell
	 * @param k
	 */
	default void defaultConvert(Cell cell, K k) {
		CellUtil.setCellValue(cell, k);
	};

	void convert(T t, K k);
}
