package com.github.liaochong.html2excel.exception;

/**
 * @author liaochong
 * @version 1.0
 */
public class NoTablesException extends RuntimeException {

    public NoTablesException() {
    }

    public NoTablesException(String message) {
        super(message);
    }

    public NoTablesException(String message, Throwable cause) {
        super(message, cause);
    }

    public static NoTablesException of(String message) {
        return new NoTablesException(message);
    }
}
