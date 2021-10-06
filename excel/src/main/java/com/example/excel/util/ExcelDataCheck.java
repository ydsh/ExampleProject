package com.example.excel.util;

@FunctionalInterface
public interface ExcelDataCheck<T> {
	boolean check(int rowIndex,T t);
}
