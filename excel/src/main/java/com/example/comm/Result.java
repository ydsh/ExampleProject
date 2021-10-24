package com.example.comm;

public class Result<T> {
    private Boolean success = true;
    private T data;
    private String message = "";
    private int code = 200;

    public Result() {
    }

    public Result<T> withSuccess(Boolean success) {
        this.success = success;
        return this;
    }

    public Result<T> withData(T data) {
        this.data = data;
        return this;
    }

    public Result<T> withMessage(String message) {
        this.message = message;
        return this;
    }

    public Result<T> withCode(int code) {
        this.code = code;
        return this;
    }

    public Boolean getSuccess() {
        return success;
    }

    public void setSuccess(Boolean success) {
        this.success = success;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    @Override
    public String toString() {
        return "{\"success\":" + success + ", \"data\":" + data + ", \"message\":" + message + ", \"code\":" + code
                + "}";
    }
}