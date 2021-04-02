package com.carbonfive.sstemplates;

@SuppressWarnings("serial")
public class SsTemplateException extends Exception {

    public SsTemplateException() {

        super();
    }

    public SsTemplateException(String message) {

        super(message);
    }

    public SsTemplateException(String message, Throwable cause) {

        super(message, cause);
    }

    public SsTemplateException(Throwable cause) {

        super(cause);
    }
}
