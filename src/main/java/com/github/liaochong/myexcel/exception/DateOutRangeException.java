package com.github.liaochong.myexcel.exception;

/**
 * 日期越界异常类
 *
 * @author ZhuKun
 * @since 2019-10-12
 */
public class DateOutRangeException extends ExcelBuildException {
    public DateOutRangeException(String message) {
        super(message);
    }

    public DateOutRangeException(String message, Throwable cause) {
        super(message, cause);
    }

    public static DateOutRangeException of(String message, Throwable cause) {
        return new DateOutRangeException(message, cause);
    }

    public static DateOutRangeException of(String message) {
        return new DateOutRangeException(message);
    }
}
