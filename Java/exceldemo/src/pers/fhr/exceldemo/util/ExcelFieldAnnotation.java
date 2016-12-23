package pers.fhr.exceldemo.util;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel对象字段注解
 * @author FHR
 *
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)  
@Target(ElementType.FIELD)
public @interface ExcelFieldAnnotation {
	//列名
	String name();
	//列号
	int id();
	//列宽
	int width();
}
