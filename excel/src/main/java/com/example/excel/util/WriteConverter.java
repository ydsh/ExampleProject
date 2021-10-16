package com.example.excel.util;

import org.apache.poi.ss.usermodel.Cell;
/**
 * 转换器，将java数据转换成excel单元格的值
 * @param <T>
 */
@FunctionalInterface
public interface WriteConverter {
  void convert(Cell cell,Object object);
}
