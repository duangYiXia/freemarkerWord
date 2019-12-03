package com.xxkj.demo.exception;

public class UtilException extends BaseException {
    private static final long serialVersionUID = 8582645036430072514L;

    public UtilException() {
    }

    public UtilException(String message) {
        super(message);
    }

    public UtilException(Throwable cause) {
        super(cause);
    }

    public UtilException(String message, Throwable cause) {
        super(message, cause);
    }

    public UtilException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
