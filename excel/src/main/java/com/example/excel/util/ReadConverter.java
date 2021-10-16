package com.example.excel.util;

import org.apache.poi.ss.usermodel.Cell;

/**
 *  转换器，将取单元格的值转换成java数据
 * @param <T>
 */
@FunctionalInterface
public interface ReadConverter<T> {
  T convert(Cell cell);
}
