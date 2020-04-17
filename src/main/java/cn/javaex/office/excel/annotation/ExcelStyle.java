package cn.javaex.office.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelStyle {
	
	/**
	 * 样式实现类名
	 * @return
	 */
	public String value() default "cn.javaex.office.excel.style.DefaultCellStyle";
	
}
