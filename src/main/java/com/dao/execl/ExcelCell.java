package com.dao.execl;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCell {

    String value();

    int cellIndexSort() default 0;

    String format() default "yyyy-MM-dd HH:mm:ss";

    String timezone() default "Asia/Shanghai";
}

