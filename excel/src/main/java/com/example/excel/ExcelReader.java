package com.example.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.example.excel.util.CellUtil;
import com.example.excel.util.Converters;
import com.example.excel.util.Inspect;
import com.example.excel.util.ReadConverter;
import com.example.excel.util.RowUtil;
import com.example.excel.util.SheetUtil;
import com.example.excel.util.WorkbookUtil;

public class ExcelReader {
	private static final Logger logger = Logger.getLogger(ExcelReader.class.getName());
	// 单个sheet表的最大行数
	private static final int MAXROW = 100000;
	// 是否自动关闭资源
	private boolean autoClose = true;
	// 需要读取的文件
	private File file;
	private InputStream inputStream;
	// 工作簿
	private Workbook workbook;
	private Converters converters;
	// 默认转换器
	private ReadConverter<Cell, Object> defaultConverter = (Cell cell, Object obj) -> {
	};
	// 列索引和字段映射
	private Map<Integer, String> columnFieldMap;

	/**
	 * 不对外部提供创建实例
	 * 
	 * @param <T>
	 */
	private <T> ExcelReader() {
		converters = new Converters();
	}

	/**
	 * 注册读数据类的字段转换器
	 * 
	 * @param <T>
	 * @param clazz
	 * @param fieldName
	 * @param readConverter
	 * @return
	 */
	public <T> ExcelReader registerConverter(String fieldName, ReadConverter<Cell, T> readConverter) {
		converters.registerReadConverter(fieldName, readConverter);
		return this;
	}

	/**
	 * 注册读数据类的字段转换器
	 * 
	 * @param <T>
	 * @param clazz
	 * @param readConverters
	 * @return
	 */
	public <T> ExcelReader registerConverters(Map<String, ReadConverter<Cell, T>> readConverters) {
		if (readConverters != null && !readConverters.isEmpty()) {
			for (Entry<String, ReadConverter<Cell, T>> entry : readConverters.entrySet()) {
				converters.registerReadConverter(entry.getKey(), entry.getValue());
			}
		}
		return this;
	}

	/**
	 * 指定列索引和字段映射，将字段列表转换成Map，列索引和字段映射
	 * 
	 * @param fieldNames
	 * @return
	 */
	public ExcelReader withColumnField(List<String> fieldNames) {
		if (columnFieldMap == null) {
			columnFieldMap = new HashMap<Integer, String>();
		}
		columnFieldMap.clear();
		for (int i = 0, len = fieldNames.size(); i < len; i++) {
			columnFieldMap.put(i, fieldNames.get(i));
		}
		return this;
	}

	/**
	 * 指定列索引和字段映射
	 * 
	 * @param columnFieldMap
	 * @return
	 */
	public ExcelReader withColumnField(Map<Integer, String> columnFieldMap) {
		this.columnFieldMap = columnFieldMap;
		return this;
	}

	/**
	 * 从给定的文件读取
	 * 
	 * @param fileName
	 * @return
	 */
	public static ExcelReader build(String fileName) {
		File file = new File(fileName);
		return build(file);
	}

	/**
	 * 从给定的文件对象读取
	 * 
	 * @param file
	 * @return
	 */
	public static ExcelReader build(File file) {

		try {
			InputStream input = new FileInputStream(file);
			ExcelReader excelReader = build(input);
			excelReader.file = file;
			excelReader.build();
			return excelReader;
		} catch (Exception e) {
			logger.info(e.getMessage());
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 从给定的输入流读取
	 * 
	 * @param inputStream
	 * @return
	 */
	public static ExcelReader build(InputStream inputStream) {
		ExcelReader excelReader = new ExcelReader();
		excelReader.inputStream = inputStream;
		return excelReader;
	}

	/**
	 * 根据文件或输入流创建工作簿，优先文件创建工作簿
	 * 
	 * @return
	 * @throws Exception
	 */
	private ExcelReader build() throws Exception {
		if (file != null && file.isFile()) {
			workbook = WorkbookUtil.createWorkbook(file);
		} else if (inputStream != null) {
			workbook = WorkbookUtil.createWorkbook(inputStream);
		} else {
			throw new Exception("没有文件或输入流可以读取");
		}
		return this;
	}

	/**
	 * 默认读取第一个sheet表 指定读取数据的类
	 * 
	 * @param <T>
	 * @param clazz
	 * @return
	 * @throws Exception
	 */
	public <T> List<T> doRead(Class<T> clazz) throws Exception {
		if (workbook.getNumberOfSheets() == 0) {
			throw new Exception("没有足够的sheet表可以读取");
		}
		if(columnFieldMap!=null&&!columnFieldMap.isEmpty()) {
			throw new Exception("指定列索引和字段映射时，必须给定开始读取数据的行索引。");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook);
		int startRow = SheetUtil.headLastRowNum(sheet, clazz) + 1;
		if (startRow < 1) {
			logger.warning("没有读取到正确的表头");
		}
		return doRead(sheet, startRow, clazz);
	}

	/**
	 * 默认读取第一个sheet表,从指定行开始读，
	 * 
	 * @param <T>
	 * @param startRow
	 * @param clazz    读取数据的类
	 * @return
	 * @throws Exception
	 */
	public <T> List<T> doRead(int startRow, Class<T> clazz) throws Exception {
		if (workbook.getNumberOfSheets() == 0) {
			throw new Exception("没有足够的sheet表可以读取");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook);
		return doRead(sheet, startRow, clazz);
	}

	/**
	 * 指定sheet表名, 指定读取数据的类
	 * 
	 * @param <T>
	 * @param sheetName
	 * @param clazz
	 * @return
	 * @throws Exception
	 */
	public <T> List<T> doRead(String sheetName, Class<T> clazz) throws Exception {
		if (sheetName == null || "".equals(sheetName) || workbook.getSheet(sheetName) == null) {
			throw new Exception(sheetName + "表不存在");
		}
		if (clazz == null) {
			throw new Exception("必须指定读取数据的类");
		}
		if(columnFieldMap!=null&&!columnFieldMap.isEmpty()) {
			throw new Exception("指定列索引和字段映射时，必须给定开始读取数据的行索引。");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook, sheetName);
		int startRow = SheetUtil.headLastRowNum(sheet, clazz) + 1;
		return doRead(sheet, startRow, clazz);
	}

	/**
	 * 指定sheet表名, 指定开始读取的行号， 指定读取数据的类
	 * 
	 * @param <T>
	 * @param sheetName
	 * @param startRow
	 * @param clazz
	 * @return
	 * @throws Exception
	 */
	public <T> List<T> doRead(String sheetName, int startRow, Class<T> clazz) throws Exception {
		if (sheetName == null || "".equals(sheetName) || workbook.getSheet(sheetName) == null) {
			throw new Exception(sheetName + "表不存在");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook, sheetName);

		return doRead(sheet, startRow, clazz);
	}

	/**
	 * 指定sheet表, 指定开始读取的行号， 指定读取数据的类
	 * 
	 * @param <T>
	 * @param sheet
	 * @param startRow
	 * @param clazz
	 * @return
	 * @throws Exception
	 */
	private <T> List<T> doRead(Sheet sheet, int startRow, Class<T> clazz) throws Exception {
		if (sheet == null) {
			throw new Exception("sheet表不存在");
		}
		if (startRow < 0) {
			throw new Exception("开始读取的行不能小于0");
		}
		if (clazz == null) {
			throw new Exception("必须指定读取数据的类");
		}
		return analysisSheet(sheet, startRow, clazz);
	}

	/**
	 * 读取Map数据 默认读取第一个sheet表 指定开始读取的行号
	 * 
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, Object>> doRead(int startRow) throws Exception {
		if (workbook.getNumberOfSheets() == 0) {
			throw new Exception("没有足够的sheet表可以读取");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook);
		return doRead(sheet, startRow);
	}

	/**
	 * 读取Map数据， 指定sheet表名称, 指定开始读取的行号
	 * 
	 * @param sheetName
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, Object>> doRead(String sheetName, int startRow) throws Exception {
		if (sheetName == null || "".equals(sheetName) || workbook.getSheet(sheetName) == null) {
			throw new Exception(sheetName + "表不存在");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook, sheetName);
		return doRead(sheet, startRow);
	}

	/**
	 * 读取Map数据， 指定sheet, 指定开始读取的行号
	 * 
	 * @param sheet
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	private List<Map<String, Object>> doRead(Sheet sheet, int startRow) throws Exception {
		if (sheet == null) {
			throw new Exception("sheet表不存在");
		}
		if (startRow < 0) {
			throw new Exception("开始读取的行不能小于0");
		}

		return analysisSheetToMapList(sheet, startRow);
	}

	/**
	 * 解析sheet表,获取T类型的列表数据
	 * 
	 * @param <T>
	 * @param sheet
	 * @param startRow
	 * @param clazz
	 * @return
	 * @throws Exception
	 */
	private <T> List<T> analysisSheet(Sheet sheet, int startRow, Class<T> clazz) throws Exception {
		List<T> result = new ArrayList<T>();
		try {
			if (sheet == null) {
				logger.info("不能解析空的sheet表");
				throw new Exception("不能解析空的sheet表");
			}
			int rowCount = sheet.getPhysicalNumberOfRows();
			if (rowCount <= startRow || startRow < 0 || rowCount < 0) {
				logger.info("没有足够的行可以读取。");
				throw new Exception("没有足够的行可以读取。");
			}
			if (rowCount - startRow > MAXROW) {
				throw new Exception("有效数据超过最大行数，当前有效数据行数为" + (rowCount - startRow) + "，单个sheet表效数据行数最大行数为" + MAXROW
						+ "，请拆分sheet表。");
			}
			if (clazz == null) {
				throw new Exception("必须指定读取数据的类");
			}
			
			int rowLastNum = sheet.getLastRowNum();
			Map<Integer, String> columnFields = this.columnFieldMap; 
			 if(columnFieldMap==null||columnFieldMap.isEmpty()) {
				 columnFields = SheetUtil.columnFieldMap(clazz);
	           }

			for (int i = startRow; i <= rowLastNum; i++) {
				Row row = SheetUtil.getRow(sheet, i);
				result.add(analysisRow(row, columnFields, clazz));
			}

		} catch (Exception e) {
			logger.info("读取" + sheet.getSheetName() + "表出现异常");
			e.printStackTrace();
			throw new Exception("读取" + sheet.getSheetName() + "表出现异常");
		} finally {
			// 清空当前读转换器
			converters.clearReadConveter();
			// 每次读完sheet表就清空列索引和字段的映射，所以每次读sheet表之前给定列索引和字段的映射
			if(columnFieldMap!=null) {
				columnFieldMap.clear();
			}
			if (autoClose) {
				complete();
			}
		}
		return result;

	}

	/**
	 * 解析完表格，对每一条数据执行校验方法
	 * 
	 * @param <T>
	 * @param sheetName
	 * @param startRow
	 * @param clazz
	 * @param dataCheck
	 * @return
	 * @throws Exception
	 */
	private <T> List<T> doReadCheck(Sheet sheet, int startRow, Class<T> clazz, Inspect<T> dataCheck) throws Exception {
		List<T> result = new ArrayList<T>();
		try {
			Map<Integer, T> rowDatas = sheetDataToMap(sheet, startRow, clazz);
			boolean mark = true;
			for (Integer k : rowDatas.keySet()) {
				mark &= dataCheck.check(k, rowDatas.get(k));
				result.add(rowDatas.get(k));
			}
			// 有校验不通过的数据，则清空结果
			if (!mark) {
				result.clear();
			}

		} catch (Exception e) {
			logger.info("发生未知异常");
			e.printStackTrace();
			throw new Exception(e.getMessage());
		} finally {
			if (autoClose) {
				complete();
			}
		}
		return result;
	}

	/**
	 * 解析完表格，对每一条数据执行校验方法 默认读取第一个sheet
	 * 
	 * @param <T>
	 * @param clazz
	 * @param dataCheck
	 * @return
	 * @throws Exception
	 */
	public <T> List<T> doReadCheck(Class<T> clazz, Inspect<T> dataCheck) throws Exception {

		if (workbook.getNumberOfSheets() == 0) {
			throw new Exception("sheet表格不存在");
		}
		if(clazz==null) {
			throw new Exception("必须指定读取数据的类。");
		}
		if(columnFieldMap!=null&&!columnFieldMap.isEmpty()) {
			throw new Exception("指定列索引和字段映射时，必须给定开始读取数据的行索引。");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook);
		int startRow = SheetUtil.headLastRowNum(sheet, clazz) + 1;

		return doReadCheck(sheet, startRow, clazz, dataCheck);
	}

	/**
	 * 解析完表格，对每一条数据执行校验方法 默认读取第一个sheet并且从指定的行开始读取
	 * 
	 * @param <T>
	 * @param startRow
	 * @param clazz
	 * @param dataCheck
	 * @return
	 * @throws Exception
	 */
	public <T> List<T> doReadCheck(int startRow, Class<T> clazz, Inspect<T> dataCheck) throws Exception {
		if (workbook.getNumberOfSheets() == 0) {
			throw new Exception("sheet表格不存在");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook);

		return doReadCheck(sheet, startRow, clazz, dataCheck);

	}

	/**
	 * 行索引和数据关系映射
	 * 
	 * @param <T>
	 * @param sheet
	 * @param startRow
	 * @param clazz
	 * @return
	 * @throws Exception
	 */
	private <T> Map<Integer, T> sheetDataToMap(Sheet sheet, int startRow, Class<T> clazz) throws Exception {
		Map<Integer, T> result = new HashMap<Integer, T>();
		try {
			if (sheet == null) {
				logger.info("不能解析空的sheet表");
				throw new Exception("不能解析空的sheet表");
			}
			int rowCount = sheet.getPhysicalNumberOfRows();
			if (rowCount <= startRow || startRow < 0 || rowCount < 0) {
				logger.info("没有足够的行可以读取。");
				throw new Exception("没有足够的行可以读取。");
			}
			if (rowCount - startRow > MAXROW) {
				throw new Exception("有效数据超过最大行数，当前有效数据行数为" + (rowCount - startRow) + "，单个sheet表效数据行数最大行数为" + MAXROW
						+ "，请拆分sheet表。");
			}

			int rowLastNum = sheet.getLastRowNum();

			Map<Integer, String> columnFieldMap = SheetUtil.columnFieldMap(clazz);

			for (int i = startRow; i <= rowLastNum; i++) {
				Row row = SheetUtil.getRow(sheet, i);
				result.put(i, analysisRow(row, columnFieldMap, clazz));
			}

		} catch (Exception e) {
			logger.info("读取" + sheet.getSheetName() + "表出现异常");
			e.printStackTrace();
			throw new Exception("读取" + sheet.getSheetName() + "表出现异常");
		} finally {
			// 清空当前读转换器
			converters.clearReadConveter();
			if (autoClose) {
				complete();
			}
		}
		return result;

	}

	/**
	 * 解析sheet表,获取Map类型的列表数据
	 * 
	 * @param sheet
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	private List<Map<String, Object>> analysisSheetToMapList(Sheet sheet, int startRow) throws Exception {
		List<Map<String, Object>> result = new ArrayList<Map<String, Object>>();
		try {
			if (sheet == null) {
				logger.info("不能解析空的sheet表");
				throw new Exception("不能解析空的sheet表");
			}
			int rowCount = sheet.getPhysicalNumberOfRows();
			if (rowCount <= startRow || startRow < 0 || rowCount < 0) {
				logger.info("没有足够的行可以读取。");
				throw new Exception("没有足够的行可以读取。");
			}
			if (rowCount - startRow > MAXROW) {
				throw new Exception("有效数据超过最大行数，当前有效数据行数为" + (rowCount - startRow) + "，单个sheet表效数据行数最大行数为" + MAXROW
						+ "，请拆分sheet表。");
			}
			int rowLastNum = sheet.getLastRowNum();
			Map<Integer, String> columnNameMap = new HashMap<Integer, String>();
			int headRowIndex = startRow - 1;
			Row headRow = SheetUtil.getRow(sheet, headRowIndex < 0 ? 0 : headRowIndex);
			int colLastNum = headRow.getLastCellNum();
			for (int i = 0; i < colLastNum; i++) {
				if (startRow == 0) {
					columnNameMap.put(i, String.valueOf(i));
				} else {
					Cell cell = RowUtil.getCell(headRow, i);
					cell.getColumnIndex();
					if (cell != null && cell.getStringCellValue() != null && !"".equals(cell.getStringCellValue())) {
						columnNameMap.put(i, cell.getStringCellValue());
					}
				}
			}
			for (int i = startRow; i <= rowLastNum; i++) {
				Row row = SheetUtil.getRow(sheet, i);
				result.add(analysisRowToMap(row, columnNameMap));
			}

		} catch (Exception e) {
			logger.info("读取" + sheet.getSheetName() + "表出现异常");
			e.printStackTrace();
			throw new Exception("读取" + sheet.getSheetName() + "表出现异常");
		} finally {
			//
			if (autoClose) {
				complete();
			}
		}
		return result;
	}

	/**
	 * 将行数据解析成T类型数据
	 * 
	 * @param <T>
	 * @param row
	 * @param columnFieldMap
	 * @param clazz
	 * @param converters
	 * @return
	 * @throws Exception
	 */
	private <T> T analysisRow(Row row, Map<Integer, String> columnFieldMap, Class<T> clazz) throws Exception {
		if(clazz==null) {
			throw new Exception("必须指定读取数据的类。");
		}
		
		T t = clazz.getDeclaredConstructor().newInstance();
		// 遍历给定的列和字段的映射，不用对一行的每一列进行遍历
		for (Integer colIndex : columnFieldMap.keySet()) {

			String fieldName = columnFieldMap.get(colIndex);
			// setter方法名必须是严格的setXxx格式
			String setterMethodName = "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
			Field field = clazz.getDeclaredField(fieldName);
			field.setAccessible(true);
			// setter方法必须只带一个参数，并且参数类型与子段类型一致
			Method method = clazz.getDeclaredMethod(setterMethodName, field.getType());
			method.setAccessible(true);
			Cell cell = RowUtil.getCell(row, colIndex);
			// 获取转换器
			Map<String, ReadConverter<Cell, ?>> cv = converters.getReadConveters();
			if (cv != null && !cv.isEmpty() && cv.containsKey(fieldName)) {
				((ReadConverter<Cell, T>) cv.get(fieldName)).convert(cell, t);
			} else {
				Object value = defaultConverter.defaultConvert(cell);
				if (value == null || "".equals(value)) {
					continue;
				}
				// BigDecimal类型转其他数值类型
				if (value != null && "BigDecimal".equals(value.getClass().getSimpleName())
						&& !"BigDecimal".equals(field.getType().getSimpleName())) {
					value = CellUtil.bigDecimalToNum(new BigDecimal(value.toString()), field.getType().getSimpleName());
				}
				if (value == null || field.getType().isInstance(value) || (value instanceof Number && CellUtil
						.isSimilarNumType(field.getType().getSimpleName(), value.getClass().getSimpleName()))) {
					method.invoke(t, value);
				} else {
					String message = String.format("类型不匹配，字段%s期望的类型是%s,实际得到的是%s%n",
							CellUtil.columnName(fieldName, clazz), field.getType().getName(),
							value.getClass().getName());
					throw new Exception(message);
				}
			}
		}

		return t;
	}

	/**
	 * 行数据解析成map数据
	 * 
	 * @param row
	 * @param columnFieldMap
	 * @return
	 */
	private Map<String, Object> analysisRowToMap(Row row, Map<Integer, String> columnFieldMap) {
		Map<String, Object> result = new HashMap<String, Object>();
		for (Integer colIndex : columnFieldMap.keySet()) {
			Cell cell = RowUtil.getCell(row, colIndex);
			// 获取转换器
			Map<String, ReadConverter<Cell, ?>> cv = converters.getReadConveters();
			if (cv != null && !cv.isEmpty() && cv.containsKey(columnFieldMap.get(colIndex))) {
				((ReadConverter<Cell, Object>) cv.get(columnFieldMap.get(colIndex))).convert(cell, result);
			} else {
				Object value = defaultConverter.defaultConvert(cell);
				result.put(columnFieldMap.get(colIndex), value);
			}
		}

		return result;
	}

	/**
	 * 是否自动关闭资源，默认自动关闭
	 * 
	 * @param autoClose
	 */
	public ExcelReader withAutoClose(boolean autoClose) {
		this.autoClose = autoClose;
		return this;
	}

	/**
	 * 释放资源
	 * 
	 * @return
	 */
	public boolean complete() {
		try {
			if (workbook != null) {
				if (workbook instanceof SXSSFWorkbook) {
					// 清理临时文件
					((SXSSFWorkbook) workbook).dispose();
				}
				workbook.close();
			}
			if (inputStream != null) {
				inputStream.close();
			}
			logger.info("资源释放完成");
		} catch (Exception ex) {
			logger.warning("关闭IO资源发生异常");
			ex.printStackTrace();
			return false;
		}
		return true;
	}
}
