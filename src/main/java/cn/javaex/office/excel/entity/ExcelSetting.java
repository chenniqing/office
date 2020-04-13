package cn.javaex.office.excel.entity;

import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;

import cn.javaex.office.excel.style.DefaultCellStyle;
import cn.javaex.office.excel.style.ICellStyle;

/**
 * Excel模板配置类
 * 
 * @author 陈霓清
 */
public class ExcelSetting {
	private String sheetName;                      // sheet页名称
	private String title;                          // 顶部标题/说明
	private List<String[]> headerList;             // 表头
	private List<String[]> dataList;               // 数据
	private String selectSheetName = "下拉数据";     // 下拉数据的sheet页名称
	private List<String[]> selectDataList;         // 下拉数据
	private String[] selectColArr;                 // 指定sheet1中需要下拉的列
	private Integer columnWidth = 10;              // 列宽
	private Integer maxRow;                        // 下拉数据来源作用于sheet1的最大行
	private List<CellRangeAddress> rangeList;      // 合并单元格
	private ICellStyle cellStyle = new DefaultCellStyle();    // 单元格样式
	
	public ICellStyle getCellStyle() {
		return cellStyle;
	}
	public void setCellStyle(ICellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

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
	 * 得到下拉数据的sheet页名称
	 * @return
	 */
	public String getSelectSheetName() {
		return selectSheetName;
	}
	/**
	 * 设置下拉数据的sheet页名称
	 * @param selectSheetName
	 */
	public void setSelectSheetName(String selectSheetName) {
		this.selectSheetName = selectSheetName;
	}
	
	/**
	 * 得到下拉数据
	 * @return
	 */
	public List<String[]> getSelectDataList() {
		return selectDataList;
	}
	/**
	 * 设置下拉数据
	 * @param selectDataList
	 */
	public void setSelectDataList(List<String[]> selectDataList) {
		this.selectDataList = selectDataList;
	}
	
	/**
	 * 得到指定sheet1中需要下拉的列
	 * @return
	 */
	public String[] getSelectColArr() {
		return selectColArr;
	}
	/**
	 * 设置指定sheet1中需要下拉的列
	 * @param columnWidth
	 */
	public void setSelectColArr(String[] selectColArr) {
		this.selectColArr = selectColArr;
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
	 * 得到下拉数据来源作用于sheet1的最大行
	 * @return
	 */
	public Integer getMaxRow() {
		return maxRow;
	}
	/**
	 * 设置下拉数据来源作用于sheet1的最大行
	 * @param maxRow
	 */
	public void setMaxRow(Integer maxRow) {
		this.maxRow = maxRow;
	}
	
	/**
	 * 得到合并单元格list
	 * @return
	 */
	public List<CellRangeAddress> getRangeList() {
		return rangeList;
	}
	/**
	 * 设置合并单元格list
	 * @param rangeList
	 */
	public void setRangeList(List<CellRangeAddress> rangeList) {
		this.rangeList = rangeList;
	}
}
