package cn.javaex.office.excel.entity;

import java.util.List;

import cn.javaex.office.excel.style.DefaultCellStyle;
import cn.javaex.office.excel.style.ICellStyle;

/**
 * Excel模板配置类
 * 
 * @author 陈霓清
 */
public class ExcelSetting {
	private String sheetName;                                 // sheet页名称
	private String title;                                     // 顶部标题/说明
	private List<String[]> headerList;                        // 表头
	private List<String[]> dataList;                          // 数据
	private Integer columnWidth = 10;                         // 列宽
	private ICellStyle cellStyle = new DefaultCellStyle();    // 单元格样式
	
	/**
	 * 得到sheet页名称
	 * @return
	 */
	public String getSheetName() {
		return sheetName;
	}
	
	/**
	 * 设置sheet页名称
	 * @param sheet1Name
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	/**
	 * 得到顶部标题/说明
	 * @return
	 */
	public String getTitle() {
		return title;
	}
	/**
	 * 设置顶部标题/说明
	 * @param title
	 */
	public void setTitle(String title) {
		this.title = title;
	}
	/**
	 * 得到表头
	 * @return
	 */
	public List<String[]> getHeaderList() {
		return headerList;
	}
	/**
	 * 设置表头
	 * @param headerList
	 */
	public void setHeaderList(List<String[]> headerList) {
		this.headerList = headerList;
	}

	/**
	 * 得到数据
	 * @return
	 */
	public List<String[]> getDataList() {
		return dataList;
	}
	/**
	 * 设置数据
	 * @param demoList
	 */
	public void setDataList(List<String[]> dataList) {
		this.dataList = dataList;
	}
	
	/**
	 * 得到列宽
	 * @return
	 */
	public Integer getColumnWidth() {
		return columnWidth;
	}
	/**
	 * 设置列宽
	 * @param columnWidth
	 */
	public void setColumnWidth(Integer columnWidth) {
		this.columnWidth = columnWidth;
	}

	/**
	 * 得到单元格样式
	 * @return
	 */
	public ICellStyle getCellStyle() {
		return cellStyle;
	}
	/**
	 * 设置单元格样式
	 * @param cellStyle
	 */
	public void setCellStyle(ICellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
}
