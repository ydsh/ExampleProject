package com.example.excel.util;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
/**
 *全局读写转换器 
 *
 */
public class Converters{
	private Map<String,ReadConverter<Cell,?>> globalReadConverters = new HashMap<String,ReadConverter<Cell, ?>>(0);
	private Map<String,WriteConverter<Cell,?>> globalWriteConverters = new HashMap<String,WriteConverter<Cell,?>>(0);
    /**
     * 获取读转换器
     * @param <T>
     * @param clazz
     * @return
     */
	public <T> Map<String, ReadConverter<Cell, ?>> getReadConveters(){
		return globalReadConverters;
	}
	/**
	 *  获取写转换器
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	public <T> Map<String, WriteConverter<Cell, ?>> getWriteConveters(){
		return globalWriteConverters;
	}
	/**
	 *  注册读转换器
	 * @param <T>
	 * @param clazz
	 * @param fieldName
	 * @param converter
	 */
	public <T> void registerReadConverter(String fieldName,ReadConverter<Cell, T> converter) {
		globalReadConverters.put(fieldName, converter);
	}
	/**
	 *  注册写转换器
	 * @param <T>
	 * @param clazz
	 * @param fieldName
	 * @param converter
	 */
	public <T> void registerWriteConverter(String fieldName,WriteConverter<Cell, T> converter) {
		globalWriteConverters.put(fieldName, converter);
	}
    /**
     * 清空读转换器
     */
	public void clearReadConveter() {
		globalReadConverters.clear();
	}
	/**
	 * 清空写转换器
	 */
	public void clearWriteConverter() {
		globalWriteConverters.clear();
	}
}
