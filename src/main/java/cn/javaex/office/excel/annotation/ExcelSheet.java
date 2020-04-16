package cn.javaex.office.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {
	
	/**
	 * Sheet页名称
	 * @return
	 */
	public String name() default "Sheet1";

	/**
	 * 顶部标题
	 * @return
	 */
	public String title() default "";
	
}
