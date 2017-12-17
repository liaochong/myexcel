package com.github.liaochong.html2excel.exception;

/**
 * @author liaochong
 * @version 1.0
 */
public class ExcelBuildException extends RuntimeException {

    public ExcelBuildException(String message) {
        super(message);
    }

    public ExcelBuildException(String message, Throwable cause) {
        super(message, cause);
    }

    public static ExcelBuildException of(String message, Throwable cause) {
        return new ExcelBuildException(message, cause);
    }
}
