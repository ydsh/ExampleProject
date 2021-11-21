package com.example.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

public final class FileUtil {
	private static final Logger logger = Logger.getLogger(FileUtil.class.getName());
	private static final int WRITE_BUFF_SIZE = 1024;

	/**
	 * 不给外部提供创建实例
	 */
	private FileUtil() {

	}
	public static void copyFile(String targetFilePath, String sourceFilePath) throws Exception{
		File targetFile = new File(targetFilePath);
		File sourceFile = new File(sourceFilePath);
		copyFile(targetFile, sourceFile);
	}
    /**
     * copy file
     * @param targetFile
     * @param sourceFile
     * @throws Exception
     */
	public static void copyFile(File targetFile, File sourceFile) throws Exception {
		InputStream inputStream = null;
		OutputStream outputStream = null;

		try {
			inputStream = new FileInputStream(sourceFile);
			outputStream = new FileOutputStream(targetFile);
			int bytesRead;
			byte[] buffer = new byte[WRITE_BUFF_SIZE];
			while ((bytesRead = inputStream.read(buffer, 0, WRITE_BUFF_SIZE)) != -1) {
				outputStream.write(buffer, 0, bytesRead);
			}
			logger.info("文件复制完成");
		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception("不能拷贝文件");
		} finally {
			try {
				if (outputStream != null) {
					outputStream.close();
				}
				if (inputStream != null) {
					inputStream.close();
				}
			} catch (IOException e) {
				throw new Exception("不能关闭输入流或输出流");
			}
		}
	}

}
