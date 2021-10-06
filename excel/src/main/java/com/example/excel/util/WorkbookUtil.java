package com.example.excel.util;

import java.io.File;
import java.io.InputStream;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookUtil {
	private final static Logger logger = Logger.getLogger(WorkbookUtil.class.getName());
	public static final String XLS = ".xls";
	public static final String XLSX = ".xlsx";

	/**
	 * 创建HSSFWorkbook / XSSFWorkbook,针对XSSFWorkbook创建SXSSFWorkbook。
	 * 因此，xlsx为true则创建SXSSFWorkbook，为false则创建HSSFWorkbook。
	 * 
	 * @param xlsx
	 * @return
	 * @throws Exception
	 */
	public static Workbook createWorkbook(boolean xlsx) throws Exception {
		Workbook workbook = WorkbookFactory.create(xlsx);
		if (xlsx) {
			workbook = new SXSSFWorkbook((XSSFWorkbook) workbook);
			logger.info("创建SXSSFWorkbook实例");
		} else {
			logger.info("创建HSSFWorkbook实例");
		}
		return workbook;
	}

	/**
	 * 使用文件对象创建HSSFWorkbook / XSSFWorkbook
	 * 
	 * @param file
	 * @return
	 * @throws Exception
	 */
	public static Workbook createWorkbook(File file) throws Exception {
		String fileName = file.getName();
		if(!fileName.toLowerCase().endsWith(XLS.toLowerCase())&&
				!fileName.toLowerCase().endsWith(XLSX.toLowerCase())) {
			throw new Exception("文件类型不合法");
		}
		return WorkbookFactory.create(file);
	}

	/**
	 * 使用文件创建HSSFWorkbook / XSSFWorkbook
	 * 
	 * @param fileName
	 * @return
	 * @throws Exception
	 */
	public static Workbook createWorkbook(String fileName) throws Exception {
		File file = new File(fileName);
		return createWorkbook(file);
	}

	/**
	 * 使用输入流创建HSSFWorkbook / XSSFWorkbook ,
	 * 如果有内存要求，最好使用{@link #createWorkbook(File)}或者{@link #creaWorkbook(String)}
	 * 
	 * @param inputStream
	 * @return
	 * @throws Exception
	 */
	public static Workbook createWorkbook(InputStream inputStream) throws Exception {
		// POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
		// return WorkbookFactory.create(poifsFileSystem);

		return WorkbookFactory.create(inputStream);
	}
}
