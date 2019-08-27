package com.hotstrip.DoExcel.exception;

public class DoExcelException extends RuntimeException {
    public DoExcelException(String message) {
        super(message);
    }


    public DoExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public DoExcelException(Throwable cause) {
        super(cause);
    }
}
