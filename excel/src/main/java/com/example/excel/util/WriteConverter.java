package com.example.excel.util;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 转换器，将java数据转换成excel单元格的值
 * @param <T>
 */
@FunctionalInterface
public interface WriteConverter<T extends Cell,K>{
	/**
	 * 默认方法
	 * @param cell
	 * @param s
	 */
  default void defaultConvert(Cell cell,K k) {
	  CellUtil.setCellValue(cell, k);
  };
  void convert(T t,K k);
}
