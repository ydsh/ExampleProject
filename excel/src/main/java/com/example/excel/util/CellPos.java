package com.example.excel.util;

/**
 * 单元格坐标位置， rowIdex行坐标，colIndex列坐标
 */
public class CellPos {
	// 行序号
	private int rowIndex;
	// 列序号
	private int colIndex;

	private CellPos() {
	}

	public static CellPos build() {
		CellPos cellPos = new CellPos();
		return cellPos;
	}

	public CellPos withRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
		return this;
	}

	public CellPos withColIndex(int colIndex) {
		this.colIndex = colIndex;
		return this;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public int getColIndex() {
		return colIndex;
	}

	@Override
	public String toString() {
		return "(" + rowIndex + "," + colIndex + ")";
	}

}
