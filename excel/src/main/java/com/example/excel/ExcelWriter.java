package com.example.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.example.comm.Excel;
import com.example.excel.util.CellUtil;
import com.example.excel.util.FileUtil;
import com.example.excel.util.RowUtil;
import com.example.excel.util.SheetUtil;
import com.example.excel.util.WorkbookUtil;

public class ExcelWriter {
	private static final Logger logger = Logger.getLogger(ExcelWriter.class.getName());
	// 单个sheet表写入最大行数
	private static final int MAXROW = 100000;
	// 标记是否需要自动释放资源
	private boolean autoClose = true;
	// 输出流
	private OutputStream outputStream;
	private Workbook workbook;
	private boolean isTemplate = false;

	private Map<String, CellStyle> fmtMap;

	/**
	 * 不提供外部创建实例
	 */
	private ExcelWriter() {
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	/**
	 * 写入给定的文件
	 * 
	 * @param fileName
	 * @return
	 */
	public static ExcelWriter writeExcel(String fileName) {
		File file = new File(fileName);
		return writeExcel(file);
	}

	/**
	 * 写入给定的文件对象
	 * 
	 * @param file
	 * @return
	 */
	public static ExcelWriter writeExcel(File file) {
		ExcelWriter excelWriter = new ExcelWriter();
		try {
			excelWriter.outputStream = new FileOutputStream(file);
		} catch (FileNotFoundException e) {
			logger.warning("创建输出流异常");
			e.printStackTrace();
		}
		return excelWriter;
	}

	/**
	 * 写入给定的的输出流
	 * 
	 * @return
	 */
	public static ExcelWriter writeExcel(OutputStream outputStream) {
		ExcelWriter excelWriter = new ExcelWriter();
		excelWriter.outputStream = outputStream;
		return excelWriter;
	}

	/**
	 * 是否自动关闭资源
	 * 
	 * @param autoClose
	 * @return
	 */
	public ExcelWriter withAutoClose(boolean autoClose) {
		this.autoClose = autoClose;
		return this;
	}

	/**
	 * 创建默认工作簿
	 * 
	 * @return
	 * @throws Exception
	 */
	public ExcelWriter build() throws Exception {
		this.isTemplate = false;
		try {
			if (workbook == null) {
				this.workbook = WorkbookUtil.createWorkbook(true);
			}
		} catch (Exception e) {
			logger.info("创建默认workbook模板发生异常。");
			throw new Exception("创建默认workbook模板发生异常。");
		}

		return this;
	}

	/**
	 * 使用指定模板创建工作簿
	 * 
	 * @param templateName
	 * @return
	 * @throws Exception
	 */
	public ExcelWriter build(String templatePath) throws Exception {

		if (templatePath == null || "".equals(templatePath)) {
			throw new Exception(templatePath + "模板不存在");
		}
		File sourceFile = new File(templatePath);
		if (!sourceFile.isFile()) {
			throw new Exception(templatePath + "不是模板文件");
		}
		String tempFile = sourceFile.getParent() + File.separator + System.currentTimeMillis() + sourceFile.getName();
		File targetFile = new File(tempFile);
		try {
			FileUtil.copyFile(targetFile, sourceFile);
			this.workbook = WorkbookUtil.createWorkbook(targetFile);
			this.isTemplate = true;
		} catch (Exception e) {
			logger.info("创建workbook模板发生异常。");
			throw new Exception("创建workbook模板发生异常。");
		} finally {
			targetFile.deleteOnExit();
		}
		return this;
	}

	/**
	 * 将数据写入工作簿,默认写入第一个sheet表中
	 * 
	 * @param <T>
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	public <T> ExcelWriter doWrite(List<T> dataList) throws Exception {
		Sheet sheet = null;
		Map<Integer, String> columnMap = null;
		int startRow = -1;

		// 默认使用第一个sheet表格
		if (workbook.getNumberOfSheets() > 0) {
			sheet = workbook.getSheetAt(0);
		} else {
			sheet = SheetUtil.createSheet(workbook);
		}
		if (dataList != null && !dataList.isEmpty()) {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.headRowWrite(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
			startRow = SheetUtil.headRowCount(t.getClass());
		}
		this.doWrite(sheet, startRow, columnMap, dataList);
		return this;
	}

	/**
	 * 默认sheet表，指定开始写数据的行索引
	 * 
	 * @param <T>
	 * @param startRow
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	public <T> ExcelWriter doWrite(int startRow, List<T> dataList) throws Exception {
		Sheet sheet = null;
		Map<Integer, String> columnMap = null;
		// 默认使用第一个sheet表格
		if (workbook.getNumberOfSheets() > 0) {
			sheet = workbook.getSheetAt(0);
		} else {
			sheet = SheetUtil.createSheet(workbook);
		}
		if (dataList != null && !dataList.isEmpty()) {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.headRowWrite(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
		}

		this.doWrite(sheet, startRow, columnMap, dataList);
		return this;
	}

	/**
	 * 指定sheet表名字写数据
	 * 
	 * @param <T>
	 * @param sheetName
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	public <T> ExcelWriter doWrite(String sheetName, List<T> dataList) throws Exception {
		if (sheetName == null || "".equals(sheetName)) {
			logger.warning("sheet表不存在！");
			throw new Exception("sheet表不存在！");
		}
		Sheet sheet = null;
		if (workbook.getSheet(sheetName) != null) {
			sheet = workbook.getSheet(sheetName);
		} else {
			sheet = SheetUtil.createSheet(workbook, sheetName);
		}
		Map<Integer, String> columnMap = null;
		int startRow = -1;
		if (dataList != null && !dataList.isEmpty()) {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.headRowWrite(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
			startRow = SheetUtil.headRowCount(t.getClass());
		}
		this.doWrite(sheet, startRow, columnMap, dataList);
		return this;
	}

	/**
	 * 指定sheet表名字和开始写数据的行索引
	 * 
	 * @param <T>
	 * @param sheetName
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	public <T> ExcelWriter doWrite(String sheetName, int startRow, List<T> dataList) throws Exception {
		if (sheetName == null || "".equals(sheetName)) {
			logger.warning("sheet表不存在！");
			throw new Exception("sheet表不存在！");
		}
		Sheet sheet = null;
		if (workbook.getSheet(sheetName) != null) {
			sheet = workbook.getSheet(sheetName);
		} else {
			sheet = SheetUtil.createSheet(workbook, sheetName);
		}
		Map<Integer, String> columnMap = null;
		if (dataList != null && !dataList.isEmpty()) {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.headRowWrite(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
		}
		this.doWrite(sheet, startRow, columnMap, dataList);
		return this;
	}

	/**
	 * 指定sheet表格写数据
	 * 
	 * @param <T>
	 * @param sheet
	 * @param startRow
	 * @param columnMap
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	private <T> ExcelWriter doWrite(Sheet sheet, int startRow, Map<Integer, String> columnMap, List<T> dataList)
			throws Exception {
		if (workbook == null) {
			logger.warning("工作簿不存在！");
			throw new Exception("工作簿不存在！");
		}
		if (sheet == null) {
			logger.warning("工作表格不存在！");
			throw new Exception("工作表格不存在！");
		}
		if (columnMap == null || columnMap.isEmpty()) {
			logger.warning("没有表格列和字段对应关系。");
			throw new Exception("没有表格列和字段对应关系，不能写入数据");
		}
		if (startRow < 0) {
			logger.warning("写入数据行索引不能小于0");
			throw new Exception("写入数据行索引不能小于0");
		}
		if (dataList == null) {
			logger.warning("没有数据可以写。");
			throw new Exception("没有数据可以导出");
		}
		if (MAXROW < dataList.size()) {
			logger.warning("sheet表数据量超出限制");
			throw new Exception("单个sheet表最多导出" + MAXROW + "条数据，请分多个sheet表导出。");
		}
		Class<?> clazz = dataList.get(0).getClass();
		withFmtMap(workbook, clazz);
		logger.info("开始向" + sheet.getSheetName() + "表写入数据。");
		for (int i = 0, len = dataList.size(); i < len; i++) {
			Row row = SheetUtil.getRow(sheet, i + startRow);
			// 数据写入到excel表行中
			RowUtil.dataToRow(row, columnMap, dataList.get(i));
			// 设置行单元格注解定义的格式
			RowUtil.setRowCellFormat(row, fmtMap, columnMap, clazz);
			logger.info("第" + (i + 1) + "条数据写入完成。");
		}
		logger.info("所有数据写入" + sheet.getSheetName() + "表完成。");
		if (autoClose) {
			writeOut();
		}
		return this;
	}

	/**
	 * 根据模板写入数据
	 * 
	 * @param <T>
	 * @param sheetName
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	public <T> ExcelWriter doWriteTemplate(String sheetName, List<T> dataList) throws Exception {
		if (!isTemplate) {
			logger.info("没有模板");
			throw new Exception("没有模板");
		}
		Sheet sheet = null;
		// sheet模板名称不存在，默认获取第一个sheet表
		if (sheetName == null || "".equals(sheetName)) {
			sheet = workbook.getSheetAt(0);
		} else {
			sheet = workbook.getSheet(sheetName);
		}
		Map<Integer, String> columnMap = null;
		int startRow = -1;
		if (dataList != null && !dataList.isEmpty()) {
			T t = dataList.get(0);
			columnMap = SheetUtil.templateColumnFieldMap(sheet, t.getClass());
			startRow = SheetUtil.headLastRowNum(sheet, t.getClass()) + 1;
		}
		this.doWrite(sheet, startRow, columnMap, dataList);
		return this;
	}

	/**
	 * 写map数据
	 * 
	 * @param sheetName
	 * @param mapDatas
	 * @return
	 * @throws Exception
	 */
	public ExcelWriter doWriteMap(String sheetName, List<Map<String, Object>> mapDatas) throws Exception {
		try {

			Sheet sheet = workbook.createSheet(sheetName);
			if (sheet == null) {
				throw new Exception("sheet表不存在。");
			}
			if (mapDatas != null && !mapDatas.isEmpty()) {
				// 记录表头列索引
				int headColCount = 0;
				// 记录表头列索引和名字映射
				Map<Integer, String> columnMap = new HashedMap<Integer, String>();
				Map<String, Object> data = mapDatas.get(0);
				// 先写表头行数据
				Row headRow = sheet.createRow(0);
				for (String key : data.keySet()) {
					CellUtil.setCellValue(RowUtil.getCell(headRow, headColCount), key);
					columnMap.put(headColCount, key);
					headColCount += 1;
				}
				int rowCount = mapDatas.size();
				for (int i = 0; i < rowCount; i++) {
					Row row = SheetUtil.getRow(sheet, i + 1);
					RowUtil.mapDataToRow(row, columnMap, data);
				}
			}
		} catch (Exception e) {
			logger.info("写map数据出错");
			e.printStackTrace();
		} finally {
			if (autoClose) {
				writeOut();
			}
		}
		return this;
	}

	/**
	 * workbook写入输出流
	 * 
	 * @param workbook
	 */
	public ExcelWriter writeOut() {
		try {
			workbook.write(outputStream);
			outputStream.flush();
		} catch (Exception e) {
			logger.warning("写excel文件发生异常");
			e.printStackTrace();
		} finally {
			if (autoClose) {
				this.complete();
			}
		}
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
			if (outputStream != null) {
				outputStream.close();
			}
			logger.info("资源释放完成");
		} catch (Exception ex) {
			logger.warning("关闭IO资源发生异常");
			ex.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	 * 存储注解格式化数据的样式
	 * 
	 * @param <T>
	 * @param workbook
	 * @param clazz
	 */
	public <T> void withFmtMap(Workbook workbook, Class<T> clazz) {
		if (fmtMap == null) {
			fmtMap = new HashedMap<String, CellStyle>();
		}
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			// 判断字段是否标注Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				// 获取Excel注解
				Excel ex = field.getAnnotation(Excel.class);
				if (ex.order() > -1 && !"".equals(ex.fmt())) {
					DataFormat dataFormat = workbook.createDataFormat();
					fmtMap.computeIfAbsent(ex.fmt(), k -> workbook.createCellStyle())
							.setDataFormat(dataFormat.getFormat(ex.fmt()));
				}
			}
		}
	}
}
