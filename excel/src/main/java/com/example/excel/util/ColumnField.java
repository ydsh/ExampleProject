package com.example.excel.util;

import java.util.ArrayList;
import java.util.List;

/**
 * excel表格列和字段的信息
 */
public class ColumnField {
	// 列索引
	private int colIndex = -1;
	// 列的名字
	private String[] colNames = {};
	// 字段名字
	private String fieldName;
	private ColumnField() {}
    public static ColumnField build() {
    	return FuncUtil.create(ColumnField::new);
    }
	public int getColIndex() {
		return colIndex;
	}

	public void setColIndex(int colIndex) {
		this.colIndex = colIndex;
	}

	public String[] getColNames() {
		return colNames;
	}

	public void setColNames(String[] colNames) {
		this.colNames = colNames;
	}

	public String getFieldName() {
		return fieldName;
	}

	public void setFieldName(String fieldName) {
		this.fieldName = fieldName;
	}
    /**
     * 列名字和字段名称信息关系列表
     * @param fieldNames
     * @param colNamesList
     * @return
     * @throws Exception
     */
	public static List<ColumnField> columnFieldList(List<String> fieldNames,List<List<String>> colNamesList) throws Exception{
		List<ColumnField> columnFieldList = new ArrayList<ColumnField>();
		if(fieldNames==null||colNamesList==null) {
			throw new Exception("参数不能为空");
		}
		//两者长度不同取最小
		int len1 = fieldNames.size();
		int len2 = colNamesList.size();
		for(int i=0,len = len1<=len2?len1:len2;i<len;i++) {
			ColumnField columnField = ColumnField.build();
			columnField.setColIndex(i);;
			columnField.setFieldName(fieldNames.get(i));
			columnField.setColNames(colNamesList.get(i).toArray(new String[0]));
			columnFieldList.add(columnField);
		}
		return columnFieldList;
	}
	
}
