package com.csxh.util;

import org.apache.log4j.Logger;

public class Log4jUtil {

	private static Logger logger = Logger.getLogger(Log4jUtil.class);

	public static void debug(String message) {
		logger.debug(message);
	}

	public static void info(String message) {
		logger.info(message);
	}

	public static void warn(String message) {
		logger.warn(message);
	}
}
