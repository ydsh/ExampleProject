package com.example.excel.util;
/**
 * 数据校验接口
 * @param <T>
 */
@FunctionalInterface
public interface Inspect<T> {
	boolean check(int rowIndex,T t);
}
