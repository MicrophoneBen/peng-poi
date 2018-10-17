package com.github.zzlhy.exception;

/**
 * 自定义导出异常
 * Created by Administrator on 2018-2-12.
 */
public class CustomImportException extends RuntimeException{
    public CustomImportException() {
    }

    public CustomImportException(String message) {
        super(message);
    }
}
