package cn.javaex.office.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {
	
	/**
	 * 对应数据库字段的描述
	 * @return
	 */
	public String name();
	
	/**
	 * 值替换
	 *     {"1_男", "0_女"}：表示数据库值为“1”时，替换为“男”，值为“0”时，替换为“女”
	 * @return
	 */
	public String[] replace() default {};
	
	/**
	 * 排序，从 0 开始计算
	 *     如果都缺省的话，则按照成员变量的顺序自动排序
	 * @return
	 */
	public int sort() default -1;
	
	/**
	 * 导出时，每列的宽度
	 *     单位为字符。1个汉字=2个字符
	 * @return
	 */
	public int width() default 10;
	
	/**
	 * 格式化
	 *     例如：format="yyyy-MM-dd"
	 * @return
	 */
	public String format() default "";
	
	/**
	 * 类型，默认都是文本
	 *     例如：type="image"    表示该列是图片列
	 * @return
	 */
	public String type() default "";
	
}
