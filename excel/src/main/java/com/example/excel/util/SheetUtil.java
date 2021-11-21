package com.example.excel.util;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.logging.Logger;

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
import org.apache.poi.ss.util.RegionUtil;

public final class SheetUtil {
	private static final Logger logger = Logger.getLogger(SheetUtil.class.getName());

	/**
	 * 不给外部提创建供实例
	 */
	private SheetUtil() {
	}

	/**
	 * 写表头行 注：表头行索引从0开始
	 * 
	 * @param <T>
	 * @param sheet
	 * @param clazz
	 */
	public static <T> void writeHeadRow(Sheet sheet, Class<T> clazz) {
		Map<String, List<CellPos>> map = SheetUtil.headColCellsMap(clazz);
		SheetUtil.writeHead(sheet, map);
	}

	/**
	 * 写表头行 注：表头行索引从0开始
	 * 
	 * @param sheet
	 * @param list
	 */
	public static void writeHeadRow(Sheet sheet, List<ColumnField> list) {
		Map<String, List<CellPos>> map = SheetUtil.headColCellsMap(list);
		SheetUtil.writeHead(sheet, map);
	}

	/**
	 * 写表头
	 * 
	 * @param sheet
	 * @param map
	 */
	private static void writeHead(Sheet sheet, Map<String, List<CellPos>> map) {
		Workbook workbook = sheet.getWorkbook();
		CellStyle cellStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		// 写表头单元格数据和合并单元格操作
		for (String name : map.keySet()) {
			List<CellPos> list = map.get(name);
			if (list != null && !list.isEmpty()) {
				// 只有一个单元格
				if (list.size() == 1) {
					CellPos cellPos = list.get(0);
					Row row = getRow(sheet, cellPos.getRowIndex());
					Cell cell = RowUtil.getCell(row, cellPos.getColIndex());
					cell.setCellStyle(headCellStyle(cellStyle, font));
					cell.setCellValue(name);
				}
				// 不少于1个单元格
				if (list.size() > 1) {
					for (int i = 0, len = list.size(); i < len; i += 2) {
						// 合并单元格的首个单元格写入数据
						Row row = getRow(sheet, list.get(i).getRowIndex());
						Cell cell = RowUtil.getCell(row, list.get(i).getColIndex());
						cell.setCellStyle(headCellStyle(cellStyle, font));
						cell.setCellValue(name);
						// 开始单元格行列和结束单元格行列
						int firstRowIndex = list.get(i).getRowIndex();
						int firstColIndex = list.get(i).getColIndex();
						int lastRowIndex = list.get(i + 1).getRowIndex();
						int lastColIndex = list.get(i + 1).getColIndex();
						// 合并单元格
						CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRowIndex, lastRowIndex,
								firstColIndex, lastColIndex);
						// 合并单元格的边框样式
						RegionUtil.setBorderTop(BorderStyle.THIN, cellRangeAddress, sheet);
						RegionUtil.setBorderRight(BorderStyle.THIN, cellRangeAddress, sheet);
						RegionUtil.setBorderBottom(BorderStyle.THIN, cellRangeAddress, sheet);
						RegionUtil.setBorderLeft(BorderStyle.THIN, cellRangeAddress, sheet);
						// 合并单元添加到sheet
						sheet.addMergedRegion(cellRangeAddress);
					}
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
	public static <T> int templateHeadLastRowNum(Sheet sheet, Class<T> clazz) {
		int result = -1;
		Map<Integer, String> colIndexNameMap = headLastRowNameMap(clazz);
		int rowCount = sheet.getLastRowNum();
		for (int i = 0; i <= rowCount; i++) {
			Row row = getRow(sheet, i);
			int cellCount = row.getLastCellNum();
			result = i;
			int count = 0;
			for (int k = 0; k < cellCount; k++) {
				String name = RowUtil.getCell(row, k).getStringCellValue();
				if (isMergeRegion(sheet, i, k)) {
					name = SheetUtil.getHeadMergeRegionValue(sheet, i, k);
				}
				if (name != null && colIndexNameMap.containsValue(name)) {
					count++;
				}
			}
			if (count == colIndexNameMap.size()) {
				break;
			}
		}
		return result;
	}

	/**
	 * 读取excel表时，获取模板sheet表格上最后一行表头行序号
	 * 
	 * @param sheet
	 * @param list
	 * @return
	 */
	public static int templateHeadLastRowNum(Sheet sheet, List<ColumnField> list) {
		int result = -1;
		Map<Integer, String> colIndexNameMap = headLastRowNameMap(list);
		int rowCount = sheet.getLastRowNum();
		for (int i = 0; i <= rowCount; i++) {
			Row row = getRow(sheet, i);
			int cellCount = row.getLastCellNum();
			result = i;
			int count = 0;
			for (int k = 0; k < cellCount; k++) {
				String name = RowUtil.getCell(row, k).getStringCellValue();
				if (isMergeRegion(sheet, i, k)) {
					name = SheetUtil.getHeadMergeRegionValue(sheet, i, k);
				}
				if (name != null && colIndexNameMap.containsValue(name)) {
					count++;
				}
			}
			if (count == colIndexNameMap.size()) {
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
		int headLastRowNum = SheetUtil.templateHeadLastRowNum(sheet, clazz);
		Row row = SheetUtil.getRow(sheet, headLastRowNum);
		int colCount = row.getLastCellNum();
		Map<String, String> map = SheetUtil.headLastRowFieldMap(clazz);
		Map<Integer, String> result = new HashMap<Integer, String>();
		for (int i = 0; i < colCount; i++) {
			String name = null;
			if (isMergeRegion(sheet, headLastRowNum, i)) {
				name = SheetUtil.getHeadMergeRegionValue(sheet, headLastRowNum, i);

			} else {
				name = RowUtil.getCell(row, i).getStringCellValue();
			}
			if (name != null && map.containsKey(name)) {
				result.put(i, map.get(name));
			}
		}
		return result;
	}

	/**
	 * 模板列索引和字段名称映射
	 * 
	 * @param sheet
	 * @param list
	 * @return
	 */
	public static Map<Integer, String> templateColumnFieldMap(Sheet sheet, List<ColumnField> list) {
		Map<Integer, String> result = new HashMap<Integer, String>();
		int headLastRowNum = SheetUtil.templateHeadLastRowNum(sheet, list);
		Row row = SheetUtil.getRow(sheet, headLastRowNum);
		int colCount = row.getLastCellNum();
		Map<String, String> map = SheetUtil.headLastRowFieldMap(list);
		for (int i = 0; i < colCount; i++) {
			String name = null;
			if (isMergeRegion(sheet, headLastRowNum, i)) {
				name = SheetUtil.getHeadMergeRegionValue(sheet, headLastRowNum, i);

			} else {
				name = RowUtil.getCell(row, i).getStringCellValue();
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
			field.setAccessible(true);
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
	private static <T> Map<String, String> headLastRowFieldMap(Class<T> clazz) {
		Map<String, String> result = new HashMap<String, String>();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			field.setAccessible(true);
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
	 * 最后一行表头列名称和字段名称映射
	 * 
	 * @param list
	 * @return
	 */
	private static Map<String, String> headLastRowFieldMap(List<ColumnField> list) {
		Map<String, String> result = new HashMap<String, String>();
		list.forEach(data -> {
			if (data.getFieldName() != null && !"".equals(data.getFieldName()) && data.getColIndex() > -1
					&& data.getColNames().length > 0) {
				String[] names = data.getColNames();
				int len = names.length;

				result.put(names[len - 1], data.getFieldName());
			}
		});
		return result;
	}

	/**
	 * 注解最后一行表头列序号和列名称映射
	 * 
	 * @param <T>
	 * @param clazz
	 * @return
	 */
	private static <T> Map<Integer, String> headLastRowNameMap(Class<T> clazz) {
		Map<Integer, String> result = new HashMap<Integer, String>();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			field.setAccessible(true);
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
	 * 最后一行表头列序号和列名称映射
	 * 
	 * @param list
	 * @return
	 */
	private static Map<Integer, String> headLastRowNameMap(List<ColumnField> list) {
		Map<Integer, String> result = new HashMap<Integer, String>();

		list.forEach(data -> {
			String[] colNames = data.getColNames();
			int len = colNames.length;
			result.put(data.getColIndex(), colNames[len - 1]);
		});

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
				Row row = SheetUtil.getRow(sheet, firstRowIndex);
				Cell cell = RowUtil.getCell(row, firstColIndex);

				return cell.getStringCellValue();
			}
		}
		return null;
	}

	/**
	 * 表头列名和单元格区域映射 注：这里是表头列名称和它对应的单元格列表
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
			field.setAccessible(true);
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
	 * 表头列名和单元格区域映射 注：这里是表头列名称和它对应的单元格列表
	 * 
	 * @param <T>
	 * @param list
	 * @return
	 */
	private static <T> Map<String, List<CellPos>> headColCellsMap(List<ColumnField> list) {
		Map<String, List<CellPos>> result = new HashMap<String, List<CellPos>>();
		int headRowCount = headRowCount(list);
		list.forEach(data -> {
			String[] names = data.getColNames();
			int rows = names.length;
			for (int r = 0; r < headRowCount; r++) {
				String name = r < rows ? names[r] : names[rows - 1];
				result.computeIfAbsent(name, k -> new ArrayList<CellPos>())
						.add(CellPos.build().withRowIndex(r).withColIndex(data.getColIndex()));
			}
		});
		return headColCellsMapHandle(result);
	}

	/**
	 * 表头列名和单元格区域映射处理函数，如果当前表头有合并单元格，则给出合并单元格的开始位置和结束位置的单元格（这里可能会有多个开始和结束单元格，不过他们都是成对存在）;
	 * 如果没有合并单元格，则一个表头列名称对应一个单元格。
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
					List<CellPos> csList = SheetUtil.cellRegionHandle(entry.getValue());
					result.put(entry.getKey(), csList);

				}
			}
		}
		return result;
	}

	/**
	 * 单元格区域处理,列表长度大于1为合并单元格，处理后的列表只有开始位置和结束位置的单元格
	 * 
	 * @return
	 */
	private static List<CellPos> cellRegionHandle(List<CellPos> list) {
		List<CellPos> result = new ArrayList<CellPos>();
		if (list.size() == 1) {
			return list;
		}
		if (list.size() > 1) {
			// 先对行进行排序
			list.sort((r1, r2) -> {
				Integer rn1 = r1.getRowIndex();
				Integer rn2 = r2.getRowIndex();
				return rn1.compareTo(rn2);
			});
			List<List<CellPos>> cellsList = new ArrayList<List<CellPos>>();
			List<CellPos> cellList = new ArrayList<CellPos>();
			// 对连续行的单元格分组
			for (int i = 0, len = list.size(); i < len; i++) {
				if (cellList.isEmpty()) {
					cellList.add(list.get(i));
				} else {
					int rowIndex = cellList.get(cellList.size() - 1).getRowIndex();
					if (rowIndex == list.get(i).getRowIndex() || rowIndex + 1 == list.get(i).getRowIndex()) {
						cellList.add(list.get(i));
					} else {
						cellsList.add(cellList);
						cellList = new ArrayList<CellPos>();
						cellList.add(list.get(i));
					}
				}
				if (len - 1 == i) {
					cellsList.add(cellList);
				}
			}
			List<List<CellPos>> csList = new ArrayList<List<CellPos>>();
			// 每一组单元格按列排序
			cellsList.forEach(data -> {
				data.sort((c1, c2) -> {
					Integer cn1 = c1.getColIndex();
					Integer cn2 = c2.getColIndex();
					return cn1.compareTo(cn2);
				});
				List<CellPos> cList = new ArrayList<CellPos>();
				// 对连续列，再次分组
				for (int k = 0, len = data.size(); k < len; k++) {
					if (cList.isEmpty()) {
						cList.add(data.get(k));
					} else {
						//int rowIndex = cList.get(cList.size() - 1).getRowIndex();
						int colIndex = cList.get(cList.size() - 1).getColIndex();
						// 同一样列是连续的，同一列行是连续的
						if (colIndex + 1 == data.get(k).getColIndex()
								||  colIndex == data.get(k).getColIndex()) {
							cList.add(data.get(k));
						} else {
							csList.add(cList);
							cList = new ArrayList<CellPos>();
						}
					}
					
					if (len - 1 == k) {
						csList.add(cList);
					}
				}
				
			});
			
			// 找各个单元格区域的开始单元格和结束单元格(都是成对存在)
			csList.forEach(data -> {
				CellPos cellMin = SheetUtil.cellPosMin(data);
				CellPos cellMax = SheetUtil.cellPosMax(data);
				result.add(cellMin);
				result.add(cellMax);
			});
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
			field.setAccessible(true);
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
	 * 获取表头行数 注：colNames数组的长度即为行数
	 * 
	 * @param list
	 * @return
	 */
	public static int headRowCount(List<ColumnField> list) {
		Optional<ColumnField> data = list.stream().max((a, b) -> {
			Integer lena = a.getColNames().length;
			Integer lenb = b.getColNames().length;
			return lena.compareTo(lenb);
		});
		return data.get().getColNames().length;
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
	public static <T> int headColCount(Class<T> clazz) {
		int result = 0;
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			field.setAccessible(true);
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
	private static CellStyle headCellStyle(CellStyle style, Font font) {
		// Workbook workbook = sheet.getWorkbook();
		// CellStyle style = workbook.createCellStyle();
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
		// Font font = workbook.createFont();
		// 粗体
		font.setBold(true);
		// 字体族
		font.setFontName("宋体");
		// 字体高度
		font.setFontHeight(Short.valueOf("12"));
		// 字号
		font.setFontHeightInPoints(Short.valueOf("12"));

		style.setFont(font);
		return style;
	}

}
