package com.example.excel.util;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Locale;
import java.util.logging.Logger;

public class DateTimeUtil {
	private static final Logger logger = Logger.getLogger(DateTimeUtil.class.getName());

	private DateTimeUtil() {

	}

	/**
	 * LocalDate转成Date
	 * 
	 * @param localDate
	 * @return
	 */
	public static Date toDate(LocalDate localDate) {
		ZonedDateTime zonedDateTime = localDate.atStartOfDay(ZoneId.systemDefault());
		Instant instant = zonedDateTime.toInstant();
		return Date.from(instant);
	}

	/**
	 * LocalDateTime转成Date
	 * 
	 * @param localDateTime
	 * @return
	 */
	public static Date toDate(LocalDateTime localDateTime) {
		ZonedDateTime zonedDateTime = localDateTime.atZone(ZoneId.systemDefault());
		Instant instant = zonedDateTime.toInstant();
		return Date.from(instant);
	}

	/**
	 * Date转成LocalDate
	 * 
	 * @param date
	 * @return
	 */
	public static LocalDate toLocalDate(Date date) {
		Instant instant = date.toInstant();
		ZonedDateTime zonedDateTime = instant.atZone(ZoneId.systemDefault());
		return zonedDateTime.toLocalDate();
	}

	/**
	 * Date转成LocalDateTime
	 * 
	 * @param date
	 * @return
	 */
	public static LocalDateTime toLocalDateTime(Date date) {
		Instant instant = date.toInstant();
		ZonedDateTime zondDateTime = instant.atZone(ZoneId.systemDefault());
		return zondDateTime.toLocalDateTime();
	}

	/**
	 * LocaDate自定义格式化
	 * 
	 * @param localDate
	 * @param fmt
	 * @return
	 */
	public static String fmt(LocalDate localDate, String fmt) {
		DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(fmt);
		return localDate.format(dateTimeFormatter);
	}

	/**
	 * LocalDate自定义本地格式化
	 * 
	 * @param localDate
	 * @param fmt
	 * @param locale
	 * @return
	 */
	public static String fmt(LocalDate localDate, String fmt, Locale locale) {
		DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(fmt, locale);
		return localDate.format(dateTimeFormatter);
	}

	/**
	 * LocalDateTime自定义格式化
	 * 
	 * @param localDateTime
	 * @param fmt
	 * @return
	 */
	public static String fmt(LocalDateTime localDateTime, String fmt) {
		DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(fmt);
		return localDateTime.format(dateTimeFormatter);
	}

	/**
	 * LocalDateTime自定义本地格式化
	 * 
	 * @param localDateTime
	 * @param fmt
	 * @param locale
	 * @return
	 */
	public static String fmt(LocalDateTime localDateTime, String fmt, Locale locale) {
		DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(fmt, locale);
		return localDateTime.format(dateTimeFormatter);
	}

	/**
	 * Date自定义格式化
	 * 
	 * @param date
	 * @param fmt
	 * @return
	 */
	public static String fmt(Date date, String fmt) {
		String result = null;
		try {
			LocalDate localDate = toLocalDate(date);
			result = fmt(localDate, fmt);
		} catch (Exception e) {
			LocalDateTime localDateTime = toLocalDateTime(date);
			result = fmt(localDateTime, fmt);
			logger.info("LocalDateTime格式化");
		}
		return result;
	}

	/**
	 * Date自定义本地格式化
	 * 
	 * @param date
	 * @param fmt
	 * @param local
	 * @return
	 */
	public static String fmt(Date date, String fmt, Locale local) {
		String result = null;
		try {
			LocalDate localDate = toLocalDate(date);
			result = fmt(localDate, fmt, local);
		} catch (Exception e) {
			LocalDateTime localDateTime = toLocalDateTime(date);
			result = fmt(localDateTime, fmt, local);
			logger.info("LocalDateTime本地格式化");
		}
		return result;
	}
}
