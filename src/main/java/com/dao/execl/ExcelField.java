package com.dao.execl;

public class ExcelField {

    private String fieldName;

    private Class fieldType;

    private String cellName;

    private String format;

    private String timezone;

    private Integer cellIndexSort;

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public Class getFieldType() {
        return fieldType;
    }

    public void setFieldType(Class fieldType) {
        this.fieldType = fieldType;
    }

    public String getCellName() {
        return cellName;
    }

    public void setCellName(String cellName) {
        this.cellName = cellName;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public String getTimezone() {
        return timezone;
    }

    public void setTimezone(String timezone) {
        this.timezone = timezone;
    }

    public Integer getCellIndexSort() {
        return cellIndexSort;
    }

    public void setCellIndexSort(Integer cellIndexSort) {
        this.cellIndexSort = cellIndexSort;
    }
}
