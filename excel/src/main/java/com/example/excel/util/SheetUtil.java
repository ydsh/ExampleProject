package com.example.excel.util;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.example.comm.Excel;

public class SheetUtil {
	private static final Logger logger = Logger.getLogger(SheetUtil.class.getName());

	/**
	 * 不给外部提创建供实例
	 */
	private SheetUtil() {
	}

	/**
	 * 创建默认sheet表格
	 * 
	 * @param workbook
	 * @return
	 */
	public static Sheet createSheet(Workbook workbook) {
		return createSheet(workbook, null);
	}

	/**
	 * 使用自定义名字创建sheet表格
	 * 
	 * @param workbook
	 * @param sheetName
	 * @return
	 */
	public static Sheet createSheet(Workbook workbook, String sheetName) {
		Sheet sheet = null;
		if (sheetName == null) {
			sheetName = "Sheet" + (workbook.getNumberOfSheets() + 1);
		}
		logger.info("创建" + sheetName + "表格");
		sheet = workbook.createSheet(sheetName);

		return sheet;
	}

	/**
	 * 写表头行
	 * 
	 * @param <T>
	 * @param sheet
	 * @param clazz
	 */
	public static <T> void headRowWrite(Sheet sheet, Class<T> clazz) {
		int headRowCount = headRowCount(clazz);
		int headColCount = headColCount(clazz);
		Workbook workbook = sheet.getWorkbook();
		CellStyle cellStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		// 创建表头单元格和设置样式
		for (int i = 0; i < headRowCount; i++) {
			Row row = sheet.createRow(i);
			for (int k = 0; k < headColCount; k++) {
				Cell cell = row.createCell(k);
				cell.setCellStyle(headCellStyle(cellStyle,font));
			}
		}
		// 表头单元格填充数据和合并
		Map<String, List<CellPos>> map = headColCellsMap(clazz);
		for (String name : map.keySet()) {
			List<CellPos> list = map.get(name);
			if (list != null && !list.isEmpty()) {
				CellPos cellPos = list.get(0);
				// 填充数据
				sheet.getRow(cellPos.getRowIndex()).getCell(cellPos.getColIndex()).setCellValue(name);
				// 合并单元格
				if (list.size() == 2) {
					int firstRowIndex = cellPos.getRowIndex();
					int firstColIndex = cellPos.getColIndex();
					int lastRowIndex = list.get(1).getRowIndex();
					int lastColIndex = list.get(1).getColIndex();
					CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex,
							lastColIndex);
					sheet.addMergedRegion(cellRangeAddress);
				}
			}
		}
		logger.info("表头行创建完成。");
	}

	/**
	 * 读取excel表时，获取模板sheet表格上最后一行表头行序号
	 * 
	 * @param <T>
	 * @param sheet
	 * @param clazz
	 * @return
	 */
	public static <T> int headLastRowNum(Sheet sheet, Class<T> clazz) {
		int result = -1;
		Map<Integer, String> headLastRowNames = headLastRowNames(clazz);
		int rowCount = sheet.getLastRowNum();
		for (int i = 0; i <= rowCount; i++) {
			Row row = sheet.getRow(i);
			int cellCount = row.getLastCellNum();
			result = i;
			int count = 0;
			for (int k = 0; k < cellCount; k++) {
				String name = row.getCell(k).getStringCellValue();
				if (isMergeRegion(sheet, i, k)) {
					name = getHeadMergeRegionValue(sheet, i, k);

				}
				if (name != null && !headLastRowNames.containsValue(name)) {
					result = -1;
					break;
				}
				if (name != null && headLastRowNames.containsValue(name)) {
					count++;
				}
			}
			if (count == headLastRowNames.size()) {
				break;
			}
		}
		return result;
	}

	/**
	 * 模板列索引和字段名称映射
	 * 
	 * @param <T>
	 * @param sheet
	 * @param clazz
	 * @return
	 */
	public static <T> Map<Integer, String> templateColumnFieldMap(Sheet sheet, Class<T> clazz) {
		int headLastRowNum = headLastRowNum(sheet, clazz);
		Row row = sheet.getRow(headLastRowNum);
		int colCount = row.getLastCellNum();
		Map<String, String> map = headLastRowField(clazz);
		Map<Integer, String> result = new HashedMap<Integer, String>();
		for (int i = 0; i < colCount; i++) {
			String name = null;
			if (isMergeRegion(sheet, headLastRowNum, i)) {
				name = getHeadMergeRegionValue(sheet, headLastRowNum, i);

			} else {
				if (row.getCell(i) != null) {
					name = row.getCell(i).getStringCellValue();
				}
			}
			if (name != null && map.containsKey(name)) {
				result.put(i, map.get(name));
			}
		}
		return result;
	}

	/**
	 * 获取注解格列索引和字段名称映射
	 * 
	 * @param clazz
	 * @return
	 */
	public static <T> Map<Integer, String> columnFieldMap(Class<T> clazz) {
		Map<Integer, String> columnMap = new HashMap<Integer, String>();
		Field[] fileds = clazz.getDeclaredFields();
		for (Field field : fileds) {
			// 判断对象属性是否存在Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				// 将列和字段名字放入map中
				if (ex.order() != -1) {
					// 列和数据对象属性映射
					columnMap.put(ex.order(), field.getName());
				}
			}
		}
		return columnMap;
	}

	/**
	 * 注解最后一行表头列名称和字段名称映射
	 * 
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	private static <T> Map<String, String> headLastRowField(Class<T> clazz) {
		Map<String, String> result = new HashMap<String, String>();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			// 判断字段是否标注Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				if (ex.order() > -1 && ex.name().length >= 1) {
					String[] names = ex.name();
					int r = names.length;
					// 获取当前字段最后一行表头的列名称
					result.put(names[r - 1], field.getName());
				}
			}
		}
		return result;
	}

	/**
	 * 注解最后一行表头列序号和列名称映射
	 * 
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	private static <T> Map<Integer, String> headLastRowNames(Class<T> clazz) {
		Map<Integer, String> result = new HashMap<Integer, String>();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			// 判断字段是否标注Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				if (ex.order() > -1 && ex.name().length >= 1) {
					String[] names = ex.name();
					int r = names.length;
					// 获取当前字段最后一行表头的列名称
					result.put(ex.order(), names[r - 1]);
				}
			}
		}
		return result;
	}

	/**
	 * 判断是否是合并单元格
	 * 
	 * @param sheet
	 * @param rowIndex
	 * @param colIndex
	 * @return
	 */
	public static boolean isMergeRegion(Sheet sheet, int rowIndex, int colIndex) {
		int mergeRegionCount = sheet.getNumMergedRegions();
		for (int i = 0; i < mergeRegionCount; i++) {
			CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
			int firstRowIndex = cellRangeAddress.getFirstRow();
			int firstColIndex = cellRangeAddress.getFirstColumn();
			int lastRowIndex = cellRangeAddress.getLastRow();
			int lastColIndex = cellRangeAddress.getLastColumn();
			if (firstRowIndex <= rowIndex && rowIndex <= lastRowIndex && firstColIndex <= colIndex
					&& colIndex <= lastColIndex) {
				return true;
			}
		}
		return false;
	}

	/**
	 * 获取表头合并单元格的值
	 * 
	 * @param sheet
	 * @param rowIndex
	 * @param colIndex
	 * @return
	 */
	public static String getHeadMergeRegionValue(Sheet sheet, int rowIndex, int colIndex) {
		int mergeRegionCount = sheet.getNumMergedRegions();
		for (int i = 0; i < mergeRegionCount; i++) {
			CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
			int firstRowIndex = cellRangeAddress.getFirstRow();
			int firstColIndex = cellRangeAddress.getFirstColumn();
			int lastRowIndex = cellRangeAddress.getLastRow();
			int lastColIndex = cellRangeAddress.getLastColumn();
			if (firstRowIndex <= rowIndex && rowIndex <= lastRowIndex && firstColIndex <= colIndex
					&& colIndex <= lastColIndex) {
				Row row = sheet.getRow(firstRowIndex);
				Cell cell = row.getCell(firstColIndex);

				return cell.getStringCellValue();
			}
		}
		return null;
	}

	/**
	 * 表头列名和单元个区域映射，也就是说一个列名占据多少个单元格
	 * 
	 * @param <T>
	 * @param sheet
	 * @param clazz
	 * @return
	 */
	private static <T> Map<String, List<CellPos>> headColCellsMap(Class<T> clazz) {
		Map<String, List<CellPos>> result = new HashMap<String, List<CellPos>>();
		int headRowCount = headRowCount(clazz);
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			// 判断字段是否标注Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				if (ex.order() > -1 && ex.name().length >= 1) {
					int colIndex = ex.order();
					String[] names = ex.name();
					int rows = names.length;
					for (int r = 0; r < headRowCount; r++) {
						String name = r < rows ? names[r] : names[rows - 1];
						result.computeIfAbsent(name, k -> new ArrayList<CellPos>())
								.add(CellPos.build().withRowIndex(r).withColIndex(colIndex));
					}
				}
			}
		}
		return headColCellsMapHandle(result);
	}

	/**
	 * 表头列名和单元格区域映射处理，返回表头列名和单元格区域开始位置、结束位置的映射。
	 * 
	 * @param headColCellsMap
	 * @return
	 */
	private static Map<String, List<CellPos>> headColCellsMapHandle(Map<String, List<CellPos>> headColCellsMap) {
		Map<String, List<CellPos>> result = new HashMap<String, List<CellPos>>();
		for (Map.Entry<String, List<CellPos>> entry : headColCellsMap.entrySet()) {
			if (entry.getValue() != null && !entry.getValue().isEmpty()) {
				if (entry.getValue().size() == 1) {
					result.put(entry.getKey(), entry.getValue());
				} else {
					CellPos cellPosMin = cellPosMin(entry.getValue());
					CellPos cellPosMax = cellPosMax(entry.getValue());
					result.computeIfAbsent(entry.getKey(), k -> new ArrayList<CellPos>())
							.addAll(Arrays.asList(cellPosMin, cellPosMax));
				}
			}
		}
		return result;
	}

	/**
	 * 获取表头行数，标注Excel注解name属性的最大数组长度即为表头的行数
	 * 
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	public static <T> int headRowCount(Class<T> clazz) {
		int result = 0;
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			// 判断字段是否标注Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				if (ex.order() > -1 && ex.name().length > result) {
					result = ex.name().length;
				}
			}
		}
		return result;
	}

	/**
	 * 行存在直接获取否则就新建行
	 * 
	 * @param sheet
	 * @param rowIdex
	 * @return
	 */
	public static Row getRow(Sheet sheet, int rowIdex) {
		Row row = sheet.getRow(rowIdex);
		if (row == null) {
			row = sheet.createRow(rowIdex);
		}
		return row;
	}

	/**
	 * 获取表头列数，标注Excel注解order大于-1的数量
	 * 
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	private static <T> int headColCount(Class<T> clazz) {
		int result = 0;
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			// 判断字段是否标注Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				if (ex.order() > -1) {
					result += 1;
				}
			}
		}
		return result;
	}

	/**
	 * 获取单元格集合中的最小坐标
	 * 
	 * @param cellPosList
	 * @return
	 */
	private static CellPos cellPosMin(List<CellPos> cellPosList) {
		if (cellPosList != null && !cellPosList.isEmpty()) {
			CellPos cellPosMin = cellPosList.get(0);
			for (int i = 1, len = cellPosList.size(); i < len; i++) {
				CellPos cellPos = cellPosList.get(i);
				if (cellPos.getRowIndex() <= cellPosMin.getRowIndex()
						&& cellPos.getColIndex() <= cellPosMin.getColIndex()) {
					cellPosMin = cellPos;
				}
			}
			return cellPosMin;
		}
		return null;
	}

	/**
	 * 获取单元格集合中的最大坐标
	 * 
	 * @param cellPosList
	 * @return
	 */
	private static CellPos cellPosMax(List<CellPos> cellPosList) {
		if (cellPosList != null && !cellPosList.isEmpty()) {
			CellPos cellPosMax = cellPosList.get(0);
			for (int i = 1, len = cellPosList.size(); i < len; i++) {
				CellPos cellPos = cellPosList.get(i);
				if (cellPos.getRowIndex() >= cellPosMax.getRowIndex()
						&& cellPos.getColIndex() >= cellPosMax.getColIndex()) {
					cellPosMax = cellPos;
				}
			}
			return cellPosMax;
		}
		return null;
	}

	/**
	 * 表头默认单元格样式和字体
	 * 
	 * @param workbook
	 * @return
	 */
	public static CellStyle headCellStyle(CellStyle style,Font font) {
		//Workbook workbook = sheet.getWorkbook();
		//CellStyle style = workbook.createCellStyle();
		// 文字水平居中
		style.setAlignment(HorizontalAlignment.CENTER);
		// 文字垂直居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		// 设置边框样式
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		// 设置边框颜色
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		// 设置背景色
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		// 字体样式
		//Font font = workbook.createFont();
		// 粗体
		font.setBold(true);
		// 字体族
		font.setFontName("宋体");
		// 是否粗体
		font.setBold(true);
		// 字体高度
		font.setFontHeight(Short.valueOf("12"));
		// 字号
		font.setFontHeightInPoints(Short.valueOf("12"));

		style.setFont(font);
		return style;
	}

}
