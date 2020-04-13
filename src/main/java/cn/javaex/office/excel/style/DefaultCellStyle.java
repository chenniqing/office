package cn.javaex.office.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 自定义样式
 * 
 * @author 陈霓清
 */
public class DefaultCellStyle implements ICellStyle {

	/**
	 * 创建头部样式
	 */
	@Override
	public CellStyle createHeaderStyle(Workbook wb) {
		// 设置字体样式
		CellStyle cellStyle = wb.createCellStyle();
		// 水平对齐方式（居中）
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		// 垂直对齐方式（居中）
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		// 字体
		Font fontHeader = wb.createFont();
		fontHeader.setFontName("等线");
		cellStyle.setFont(fontHeader);
		
		return cellStyle;
	}

	/**
	 * 创建数据样式
	 */
	@Override
	public CellStyle createDataStyle(Workbook wb) {
		// 设置字体样式
		CellStyle cellStyle = wb.createCellStyle();
		// 水平对齐方式（居中）
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		// 垂直对齐方式（居中）
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		// 字体
		Font fontHeader = wb.createFont();
		fontHeader.setFontName("等线");
		cellStyle.setFont(fontHeader);
		
		return cellStyle;
	}
	
}
