package com.example.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.example.comm.ExcelUtils;
import com.example.excel.util.CellUtil;
import com.example.excel.util.ExcelDataCheck;
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
	private Workbook workbook;

	/**
	 * 不对外部提供创建实例
	 * 
	 * @param <T>
	 */
	private <T> ExcelReader() {
	}

	/**
	 * 从给定的文件读取
	 * 
	 * @param fileName
	 * @return
	 */
	public static ExcelReader readExcel(String fileName) {
		File file = new File(fileName);
		return readExcel(file);
	}

	/**
	 * 从给定的文件对象读取
	 * 
	 * @param file
	 * @return
	 */
	public static ExcelReader readExcel(File file) {
		ExcelReader excelReader = new ExcelReader();
		try {
			excelReader.file = file;
			excelReader.inputStream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			logger.info("创建输入流发生异常");
			e.printStackTrace();
		}
		return excelReader;
	}

	/**
	 * 从给定的输入流读取
	 * 
	 * @param inputStream
	 * @return
	 */
	public static ExcelReader readExcel(InputStream inputStream) {
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
	public ExcelReader build() throws Exception {
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
		int sheetNum = workbook.getNumberOfSheets();
		if (sheetNum == 0) {
			throw new Exception("没有足够的sheet表可以读取");
		}
		Sheet sheet = workbook.getSheetAt(0);
		int startRow = SheetUtil.headLastRowNum(sheet, clazz) + 1;
		return doRead(sheet, startRow, clazz);
	}
	/**
	 * 默认读取第一个sheet表,从指定行开始读，
	 * @param <T>
	 * @param startRow
	 * @param clazz 读取数据的类
	 * @return
	 * @throws Exception
	 */
	public <T> List<T> doRead(int startRow,Class<T> clazz) throws Exception {
		int sheetNum = workbook.getNumberOfSheets();
		if (sheetNum == 0) {
			throw new Exception("没有足够的sheet表可以读取");
		}
		Sheet sheet = workbook.getSheetAt(0);
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
		if (workbook.getSheet(sheetName) == null) {
			throw new Exception(sheetName + "表不存在");
		}
		Sheet sheet = workbook.getSheet(sheetName);
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
		if (workbook.getSheet(sheetName) == null) {
			throw new Exception(sheetName + "表不存在");
		}
		Sheet sheet = workbook.getSheet(sheetName);

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
	public <T> List<T> doRead(Sheet sheet, int startRow, Class<T> clazz) throws Exception {
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
	 * 用于没有都数据的类，把数据转成map列表 默认读取第一个sheet表 指定开始读取的行号
	 * 
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, Object>> doRead(int startRow) throws Exception {
		int sheetNum = workbook.getNumberOfSheets();
		if (sheetNum == 0) {
			throw new Exception("没有足够的sheet表可以读取");
		}
		Sheet sheet = workbook.getSheetAt(0);
		return doRead(sheet, startRow);
	}

	/**
	 * 用于没有都数据的类，把数据转成map列表 指定sheet表名称, 指定开始读取的行号
	 * 
	 * @param sheetName
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, Object>> doRead(String sheetName, int startRow) throws Exception {
		if (sheetName == null || sheetName.isBlank()) {
			throw new Exception("sheet表名称不能为空");
		}
		Sheet sheet = workbook.getSheet(sheetName);
		return doRead(sheet, startRow);
	}

	/**
	 * 用于没有都数据的类，把数据转成map列表 指定sheet, 指定开始读取的行号
	 * 
	 * @param sheet
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, Object>> doRead(Sheet sheet, int startRow) throws Exception {
		if (sheet == null) {
			throw new Exception("sheet表不存在");
		}
		if (startRow < 0) {
			throw new Exception("开始读取的行不能小于0");
		}

		return analysisSheet(sheet, startRow);
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

			int rowLastNum = sheet.getLastRowNum();

			Map<Integer, String> columnFieldMap = SheetUtil.columnFieldMap(clazz);
			for (int i = startRow; i <=rowLastNum; i++) {
				Row row = sheet.getRow(i);
				result.add(analysisRow(row, columnFieldMap, clazz));
			}

		} catch (Exception e) {
			logger.info("读取" + sheet.getSheetName() + "表出现异常");
			e.printStackTrace();

		} finally {
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
	public <T> List<T> doReadCheck(String sheetName, int startRow, Class<T> clazz, ExcelDataCheck<T> dataCheck)
			throws Exception {
		List<T> result = new ArrayList<T>();
		try {
			Sheet sheet = workbook.getSheet(sheetName);
			Map<Integer, T> rowDatas = rowDataMap(sheet, startRow, clazz);
			boolean mark = true;
			for (Integer k : rowDatas.keySet()) {
				mark &= dataCheck.check(k,rowDatas.get(k));
				result.add(rowDatas.get(k));
			}
			// 有校验不通过的数据，则清空结果
			if (!mark) {
				result.clear();
			}

		} catch (Exception e) {
			logger.info("发生未知异常");
			e.printStackTrace();
		} finally {
			complete();
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
	public <T> List<T> doReadCheck(Class<T> clazz, ExcelDataCheck<T> dataCheck) throws Exception {

		Sheet sheet = workbook.getSheetAt(0);
		if (sheet == null) {
			throw new Exception("sheet表格不存在");
		}
		String sheetName = sheet.getSheetName();
		int startRow = SheetUtil.headLastRowNum(sheet, clazz) + 1;
        System.err.println(startRow);
		return doReadCheck(sheetName, startRow, clazz, dataCheck);
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
	public <T> List<T> doReadCheck(int startRow, Class<T> clazz, ExcelDataCheck<T> dataCheck) throws Exception {
		Sheet sheet = workbook.getSheetAt(0);
		if (sheet == null) {
			throw new Exception("sheet表格不存在");
		}
		String sheetName = sheet.getSheetName();

		return doReadCheck(sheetName, startRow, clazz, dataCheck);

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
	private <T> Map<Integer, T> rowDataMap(Sheet sheet, int startRow, Class<T> clazz) throws Exception {
		Map<Integer, T> result = new HashedMap<Integer, T>();
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
				Row row = sheet.getRow(i);
				result.put(i, analysisRow(row, columnFieldMap, clazz));
			}

		} catch (Exception e) {
			logger.info("读取" + sheet.getSheetName() + "表出现异常");
			e.printStackTrace();

		} finally {
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
	private List<Map<String, Object>> analysisSheet(Sheet sheet, int startRow) throws Exception {
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
			Map<Integer, String> columnNameMap = new HashedMap<Integer, String>();
			int headRowIndex = startRow - 1;
			Row headRow = sheet.getRow(headRowIndex < 0 ? 0 : headRowIndex);
			int colLastNum = headRow.getLastCellNum();
			for (int i = 0; i < colLastNum; i++) {
				if (startRow == 0) {
					columnNameMap.put(i, String.valueOf(i));
				} else {
					Cell cell = headRow.getCell(i);
					if (cell != null && cell.getStringCellValue() != null && !"".equals(cell.getStringCellValue())) {
						columnNameMap.put(i, cell.getStringCellValue());
					}
				}
			}
			for (int i = startRow; i <= rowLastNum; i++) {
				Row row = sheet.getRow(i);
				result.add(analysisRow(row, columnNameMap));
			}

		} catch (Exception e) {
			logger.info("读取" + sheet.getSheetName() + "表出现异常");
			e.printStackTrace();
		} finally {
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
	 * @return
	 * @throws Exception
	 */
	private <T> T analysisRow(Row row, Map<Integer, String> columnFieldMap, Class<T> clazz) throws Exception {
		T t = clazz.getDeclaredConstructor().newInstance();
		// 遍历给定的列和字段的映射，不用对一行的每一列进行遍历
		for (Integer colIndex : columnFieldMap.keySet()) {
			Cell cell = row.getCell(colIndex);
			if (cell != null) {
				Object value = CellUtil.getCellValue(cell);
				String fieldName = columnFieldMap.get(colIndex);
				// setter方法名必须是严格的setXxx格式
				String setterMethodName = "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
				Field field = clazz.getDeclaredField(fieldName);
				// setter方法必须只带一个参数，并且参数类型与子段类型一致
				Method method = clazz.getDeclaredMethod(setterMethodName, field.getType());
				// BigDecimal类型转其他数值类型
				if ("BigDecimal".equals(value.getClass().getSimpleName())
						&& !"BigDecimal".equals(field.getType().getSimpleName())) {
					value = ExcelUtils.bigDecimalToNum(new BigDecimal(value.toString()),
							field.getType().getSimpleName());
				}
				method.invoke(t, value);

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
	private Map<String, Object> analysisRow(Row row, Map<Integer, String> columnFieldMap) {
		Map<String, Object> result = new HashedMap<String, Object>();
		for (Integer colIndex : columnFieldMap.keySet()) {
			Cell cell = row.getCell(colIndex);
			if (cell != null) {
				Object value = CellUtil.getCellValue(cell);
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
