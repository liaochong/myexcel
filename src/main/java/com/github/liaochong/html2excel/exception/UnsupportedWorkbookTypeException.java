package com.github.liaochong.html2excel.exception;

/**
 * @author liaochong
 * @version 1.0
 */
public class UnsupportedWorkbookTypeException extends RuntimeException {

    public UnsupportedWorkbookTypeException() {
        super();
    }

    public UnsupportedWorkbookTypeException(String message) {
        super(message);
    }

    public UnsupportedWorkbookTypeException(String message, Throwable cause) {
        super(message, cause);
    }
}
