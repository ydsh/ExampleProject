package com.example.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.example.excel.util.CellUtil;
import com.example.excel.util.Converter;
import com.example.excel.util.Converters;
import com.example.excel.util.Excel;
import com.example.excel.util.ColIndexFieldMap;
import com.example.excel.util.ColumnField;
import com.example.excel.util.FileUtil;
import com.example.excel.util.FuncUtil;
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
	// 中间文件
	private File cacheFile;
	// 工作簿
	private Workbook workbook;
	// 是否是模板
	private boolean isTemplate = false;
	// 全局样式
	private Map<String, CellStyle> fmtMap;
	// 转换器集合
	private Converters converters;
	// 默认转换器
	private Converter<Cell, Object> defaultConverter = (cell,obj) -> new Object();
    //excel表列和字段关系信息
	private List<ColumnField> columnFieldList;
	//excel表列索引和字段映射
	private ColIndexFieldMap colIndexFieldMap;

	/**
	 * 不提供外部创建实例
	 */
	private ExcelWriter() {
		converters =  Converters.build();
	}

	/**
	 * 获取工作簿实例
	 * 
	 * @return
	 */
	public Workbook getWorkbook() {
		return workbook;
	}

	/**
	 * 注册读数据类的字段转换器
	 * 
	 * @param <T>
	 * @param clazz
	 * @param fieldName
	 * @param writeConverter
	 * @return
	 */
	public <T> ExcelWriter registerConverter(String fieldName, Converter<Cell, T> writeConverter) {
		converters.registerConverter(fieldName, writeConverter);
		return this;
	}

	/**
	 * 注册读数据类的字段转换器
	 * 
	 * @param <T>
	 * @param clazz
	 * @param writeConverter
	 * @return
	 */
	public <T> ExcelWriter registerConverters(Map<String, Converter<Cell, T>> writeConverter) {
		if (writeConverter != null && !writeConverter.isEmpty()) {
			for (Entry<String, Converter<Cell, T>> entry : writeConverter.entrySet()) {
				converters.registerConverter(entry.getKey(), entry.getValue());
			}
		}
		return this;
	}
	/**
	 * 注册列名称和字段关系信息列表
	 * @param fieldNames
	 * @param colNamesList
	 * @return
	 * @throws Exception
	 */
    public ExcelWriter registerColumnFieldList(List<String> fieldNames,List<List<String>> colNamesList) throws Exception{
    	if(columnFieldList==null) {
    		columnFieldList = new ArrayList<ColumnField>();
    	}
    	columnFieldList.clear();
    	columnFieldList = ColumnField.columnFieldList(fieldNames, colNamesList);
    	//列索引和字段名称映射关系
    	if(colIndexFieldMap==null) {
    		colIndexFieldMap = ColIndexFieldMap.build();
    	}
    	colIndexFieldMap.withColnumFieldMap(columnFieldList);
    	return this;
    }
    /**
     * 注册列名称和字段关系信息列表
     * @param columnFieldList
     * @return
     */
    public ExcelWriter registerColumnFieldList(List<ColumnField> columnFieldList) {
    	this.columnFieldList = columnFieldList;
    	//列索引和字段名称映射关系
    	if(colIndexFieldMap==null) {
    		colIndexFieldMap = ColIndexFieldMap.build();
    	}
    	colIndexFieldMap.withColnumFieldMap(columnFieldList);
    	return this;
    }
    /**
     * 注册列索引和字段名称映射，用于写模板
     * @param fieldNames
     * @return
     */
    public ExcelWriter registerColIndexFieldMap(List<String> fieldNames){
    	//列索引和字段名称映射关系
    	if(colIndexFieldMap==null) {
    		colIndexFieldMap = ColIndexFieldMap.build();
    	}
    	colIndexFieldMap.withColnumFieldMap(fieldNames);
    	return this;
    }
    /**
     * 注册列索引和字段名称映射，用于写模板
     * @param map
     * @return
     */
    public ExcelWriter registerColIndexFieldMap(Map<Integer,String> map){
    	//列索引和字段名称映射关系
    	if(colIndexFieldMap==null) {
    		colIndexFieldMap = ColIndexFieldMap.build();
    	}
    	colIndexFieldMap.clear();
    	colIndexFieldMap.withColnumFieldMap(map);
    	return this;
    }
	/**
	 * 创建默认工作簿
	 * 
	 * @return
	 * @throws Exception
	 */
	public static ExcelWriter build() throws Exception {
		ExcelWriter excelWriter = FuncUtil.create(ExcelWriter::new);
		excelWriter.isTemplate = false;
		try {
			excelWriter.workbook = WorkbookUtil.createWorkbook(true);
		} catch (Exception e) {
			logger.info("创建默认workbook模板发生异常。");
			throw new Exception("创建默认workbook模板发生异常。");
		}

		return excelWriter;
	}

	/**
	 * 使用指定模板创建工作簿
	 * 
	 * @param templateName
	 * @return
	 * @throws Exception
	 */
	public static ExcelWriter build(String templatePath) throws Exception {
		ExcelWriter excelWriter = new ExcelWriter();
		if (templatePath == null || "".equals(templatePath)) {
			throw new Exception(templatePath + "模板不存在");
		}
		if (!templatePath.toLowerCase().endsWith(WorkbookUtil.XLS.toLowerCase())
				&& !templatePath.toLowerCase().endsWith(WorkbookUtil.XLSX.toLowerCase())) {
			throw new Exception("不支持其他格式文件，只支持excel格式文件");
		}
		File sourceFile = new File(templatePath);
		if (!sourceFile.isFile()) {
			throw new Exception(templatePath + "不是模板文件");
		}
		try {
			String tempFilePath = sourceFile.getParent() + File.separator + System.currentTimeMillis()
					+ sourceFile.getName();
			//复制模板
			excelWriter.cacheFile = new File(tempFilePath);
			FileUtil.copyFile(excelWriter.cacheFile, sourceFile);
			excelWriter.workbook = WorkbookUtil.createWorkbook(excelWriter.cacheFile);
			excelWriter.isTemplate = true;
		} catch (Exception e) {
			logger.info("创建workbook模板发生异常。");
			e.printStackTrace();
			throw new Exception("创建workbook模板发生异常。");
		}
		return excelWriter;
	}

	/**
	 * 定义是否自动关闭资源
	 * 
	 * @param autoClose
	 * @return
	 */
	public ExcelWriter withAutoClose(boolean autoClose) {
		this.autoClose = autoClose;
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
		Map<Integer, String> columnMap = null;
		int startRow = -1;
		if (dataList == null || dataList.isEmpty()) {
			throw new Exception("没有数据可以写。");
		}
		// 默认使用第一个sheet表格
		Sheet sheet = WorkbookUtil.getSheet(workbook);
		if (columnFieldList != null && !columnFieldList.isEmpty()) {
			SheetUtil.writeHeadRow(sheet, columnFieldList);
			startRow = SheetUtil.headRowCount(columnFieldList);
			columnMap = colIndexFieldMap.getColumnFieldMap();
		}else {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.writeHeadRow(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
			startRow = SheetUtil.headRowCount(t.getClass());
		}

		this.dataToSheet(sheet, startRow, columnMap, dataList);
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
		Map<Integer, String> columnMap = null;
		if (dataList == null || dataList.isEmpty()) {
			throw new Exception("没有数据可以写。");
		}

		// 默认使用第一个sheet表格
		Sheet sheet = WorkbookUtil.getSheet(workbook);
		if (columnFieldList != null && !columnFieldList.isEmpty()) {
			SheetUtil.writeHeadRow(sheet, columnFieldList);
			columnMap = colIndexFieldMap.getColumnFieldMap();
		}else {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.writeHeadRow(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
		}

		this.dataToSheet(sheet, startRow, columnMap, dataList);
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
		
		if (sheetName == null || "".equals(sheetName) || workbook.getSheet(sheetName) != null) {
			logger.warning("sheet表不存在！");
			throw new Exception("sheet表不存在！");
		}
		if (dataList == null || dataList.isEmpty()) {
			throw new Exception("没有数据可以写。");
		}
		Map<Integer, String> columnMap = null;
		int startRow = -1;
		Sheet sheet = WorkbookUtil.getSheet(workbook, sheetName);
		if (columnFieldList != null && !columnFieldList.isEmpty()) {
			SheetUtil.writeHeadRow(sheet, columnFieldList);
			startRow = SheetUtil.headRowCount(columnFieldList);
			columnMap = colIndexFieldMap.getColumnFieldMap();
		}else {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.writeHeadRow(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
			startRow = SheetUtil.headRowCount(t.getClass());
		}
		this.dataToSheet(sheet, startRow, columnMap, dataList);
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
		if (sheetName == null || "".equals(sheetName) || workbook.getSheet(sheetName) != null) {
			logger.warning("sheet表不存在！");
			throw new Exception("sheet表不存在！");
		}
		if (dataList == null || dataList.isEmpty()) {
			throw new Exception("没有数据可以写。");
		}
		Sheet sheet = WorkbookUtil.getSheet(workbook, sheetName);
		Map<Integer, String> columnMap = colIndexFieldMap.getColumnFieldMap();
		if (columnFieldList != null&&!columnFieldList.isEmpty()) {
			// 写入注解定义表头行数据
			SheetUtil.writeHeadRow(sheet, columnFieldList);
			columnMap = colIndexFieldMap.getColumnFieldMap();
		}else {
			T t = dataList.get(0);
			// 写入注解定义表头行数据
			SheetUtil.writeHeadRow(sheet, t.getClass());
			columnMap = SheetUtil.columnFieldMap(t.getClass());
		}
		this.dataToSheet(sheet, startRow, columnMap, dataList);
		return this;
	}

	/**
	 * 向excel的sheet表写数据
	 * 
	 * @param <T>
	 * @param sheet
	 * @param startRow
	 * @param columnMap
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	private <T> ExcelWriter dataToSheet(Sheet sheet, int startRow, Map<Integer, String> columnMap, List<T> dataList)
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
			throw new Exception("没有数据可以写。");
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
			// 设置行单元格注解定义的格式
			RowUtil.setRowCellFormat(row, fmtMap, columnMap, clazz);
			// 数据写入到excel表行中
			dataToRow(row, columnMap, dataList.get(i));
			logger.info("excel的" + sheet.getSheetName() + "表第" + (startRow + i + 1) + "行数据写入完成。");
		}
		logger.info("所有数据写入excel的 " + sheet.getSheetName() + "表完毕。");
		// 每次写完sheet表就清空列索引和字段的映射，所以每次写sheet表之前给定列索引和字段的映射
		if (colIndexFieldMap != null) {
			colIndexFieldMap.clear();
		}
		return this;
	}

	/**
	 * 将数据写入到excel表行中
	 * 
	 * @param <T>
	 * @param row
	 * @param columnMap
	 * @param data
	 * @throws Exception
	 */
	private <T> void dataToRow(Row row, Map<Integer, String> columnMap, T data) throws Exception {
		Class<T> clazz = (Class<T>) data.getClass();
		columnMap.entrySet().forEach(v->{
			Cell cell = RowUtil.getCell(row, v.getKey());
			String fieldName = v.getValue();
			Map<String, Converter<Cell, ?>> cv = converters.getConverters();
			if (cv != null && !cv.isEmpty() && cv.containsKey(fieldName)) {
				//如果这里写入了样式就会覆盖注解的样式
				((Converter<Cell, T>)cv.get(fieldName)).convert(cell, data);
			} else {
				String getterName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
				try {
					//必须是一个无参的getter方法
					Method method = clazz.getDeclaredMethod(getterName);
					method.setAccessible(true);
					// 数据写入单元格
					defaultConverter.defaultConvert(cell, method.invoke(data));
				} catch (NoSuchMethodException | SecurityException | 
						IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
					logger.warning(e.getMessage());
					e.printStackTrace();
				}
			}
		});
	
	}

	/**
	 * 指定模板名称写数据
	 * @param <T>
	 * @param sheetName
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	public <T> ExcelWriter doWriteTemplate(String sheetName, List<T> dataList) throws Exception {
		if (!isTemplate) {
			throw new Exception("没有写数据的模板");
		}
		if (dataList == null || dataList.isEmpty()) {
			throw new Exception("没有数据可以写");
		}
		// sheet模板名称不存在，默认获取第一个sheet表
		Sheet sheet = WorkbookUtil.getSheet(workbook, sheetName);
		Map<Integer, String> columnMap = null;
		int startRow = -1;
		if(columnFieldList!=null&&!columnFieldList.isEmpty()) {
			columnMap = colIndexFieldMap.getColumnFieldMap();
			startRow = SheetUtil.templateHeadLastRowNum(sheet, columnFieldList) + 1;
		}else if (columnFieldList==null&&colIndexFieldMap!=null) {
			throw new Exception("注册列索引和字段名映射写模板时要指定开始写的行索引。");
		}else{
			T t = dataList.get(0);
			columnMap = SheetUtil.templateColumnFieldMap(sheet, t.getClass());
			startRow = SheetUtil.templateHeadLastRowNum(sheet, t.getClass()) + 1;
		}
		this.dataToSheet(sheet, startRow, columnMap, dataList);
		return this;
	}
	/**
	 * 指定模板名称和开始行写数据
	 * @param <T>
	 * @param sheetName
	 * @param startRow
	 * @param dataList
	 * @return
	 * @throws Exception
	 */
	public <T> ExcelWriter doWriteTemplate(String sheetName,int startRow, List<T> dataList) throws Exception {
		if (!isTemplate) {
			throw new Exception("没有写数据的模板");
		}
		if (dataList == null || dataList.isEmpty()) {
			throw new Exception("没有数据可以写");
		}
		// sheet模板名称不存在，默认获取第一个sheet表
		Sheet sheet = WorkbookUtil.getSheet(workbook, sheetName);
		Map<Integer, String> columnMap = null;
		if(colIndexFieldMap!=null) {
			columnMap = colIndexFieldMap.getColumnFieldMap();
		}else {
			T t = dataList.get(0);
			columnMap = SheetUtil.templateColumnFieldMap(sheet, t.getClass());
		}
		
		this.dataToSheet(sheet, startRow, columnMap, dataList);
		return this;
	}
	/**
	 * map数据写入到sheet表格中
	 * 
	 * @param sheetName
	 * @param mapDatas
	 * @return
	 * @throws Exception
	 */
	public ExcelWriter mapDataToSheet(String sheetName, int startRow,List<Map<String, Object>> mapDatas) throws Exception {
		try {

			Sheet sheet = workbook.createSheet(sheetName);
			if (sheet == null) {
				throw new Exception("sheet表不存在。");
			}
			if (mapDatas != null && !mapDatas.isEmpty()) {
				// 记录表头列索引和名字映射
				Map<Integer, String> columnMap = new HashMap<Integer, String>();
				Optional<Map<String,Object>> optional = mapDatas.stream().max((a,b)->{
					Integer an = a.size();
					Integer bn = b.size();
					return an.compareTo(bn);
				});
				// 记录表头列索引
				int headColCount = 0;
				// 先写表头行数据，第一行作为表头
				Row headRow = sheet.createRow(0);
				for (String key : optional.get().keySet()) {
					CellUtil.setCellValue(RowUtil.getCell(headRow, headColCount), key);
					columnMap.put(headColCount, key);
					headColCount += 1;
				}
				int rowCount = mapDatas.size();
				for (int i = 0; i < rowCount; i++) {
					Row row = SheetUtil.getRow(sheet, i + startRow);
					mapDataToRow(row, columnMap, mapDatas.get(i));
				}
			}
		} catch (Exception e) {
			logger.info("写map数据出错");
			e.printStackTrace();
			throw new Exception("写map数据出错");
		}
		return this;
	}

	/**
	 * map数据写入到excel表行中
	 * 
	 * @param row
	 * @param columnMap
	 * @param mapData
	 * @throws Exception
	 */
	private void mapDataToRow(Row row, Map<Integer, String> columnMap, Map<String, Object> mapData) throws Exception {
		columnMap.entrySet().forEach(data->{
			Cell cell = RowUtil.getCell(row, data.getKey());
			Map<String, Converter<Cell, ?>> cv = converters.getConverters();
			if (cv != null && !cv.isEmpty() && cv.containsKey(data.getValue())) {
				((Converter<Cell, Map<String, Object>>)cv.get(data.getValue())).convert(cell, mapData);
			} else {
				// CellUtil.setCellValue(cell, mapData.get(key));
				defaultConverter.defaultConvert(cell, mapData.get(data.getValue()));
			}
		});		
	}

	/**
	 * workbook写入输出流
	 * 
	 * @param workbook
	 */
	public void writeOut(String pathName) {
		File file = new File(pathName);
		try {
			outputStream = new FileOutputStream(file);
			writeOut(outputStream);
		} catch (FileNotFoundException e) {
			logger.warning(e.getMessage());
			e.printStackTrace();
		}
	}

	/**
	 * workbook写入指定输出流
	 * 
	 * @param workbook
	 */
	public ExcelWriter writeOut(OutputStream outputStream) {
		try {
			workbook.write(outputStream);
			outputStream.flush();
		} catch (Exception e) {
			logger.warning("写excel文件发生异常，" + e.getMessage());
			e.printStackTrace();
		} finally {
			if (autoClose) {
				complete();
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
		} catch (Exception e) {
			logger.warning("关闭IO资源发生异常");
			e.printStackTrace();
			return false;
		} finally {
			if (cacheFile != null) {
				cacheFile.delete();
			}
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
	private <T> void withFmtMap(Workbook workbook, Class<T> clazz) {
		if (fmtMap == null) {
			fmtMap = new HashMap<String, CellStyle>();
		}
		fmtMap.clear();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0, len = fields.length; i < len; i++) {
			Field field = fields[i];
			field.setAccessible(true);
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
