package com.example.excel.util;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
/**
 *读写转换器 
 *
 */
public class Converters{
	private Map<String,Converter<Cell,?>> globalConverters = new HashMap<String,Converter<Cell,?>>(0);
	private Converters() {}
	
	public static Converters build() {
		return FuncUtil.create(Converters::new);
	}
    
	/**
	 *  获取转换器
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	public <T> Map<String, Converter<Cell, ?>> getConverters(){
		return globalConverters;
	}
	
	/**
	 *  注册转换器
	 * @param <T>
	 * @param clazz
	 * @param fieldName
	 * @param converter
	 */
	public <T> void registerConverter(String fieldName,Converter<Cell, T> converter) {
		globalConverters.put(fieldName, converter);
	}
    
	/**
	 * 清空转换器
	 */
	public void clearConverter() {
		globalConverters.clear();
	}
}
