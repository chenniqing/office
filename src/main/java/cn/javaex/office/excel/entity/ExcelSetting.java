package cn.javaex.office.excel.entity;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Excel模板配置类
 * 
 * @author 陈霓清
 */
public class ExcelSetting {
	private String sheet1Name;                      // sheet1名称
	private String sheet2Name;                      // sheet2名称
	private String[] headerArr;                     // 表头
	private ArrayList<String[]> demoList = new ArrayList<String[]>();        // 样例数据
	private ArrayList<String[]> selectDataList = new ArrayList<String[]>();  // 下拉数据
	private String[] selectColArr;                  // 指定sheet1中需要下拉的列
	private Integer columnWidth;                    // 列宽
	private Integer maxRow;                         // 下拉数据来源作用于sheet1的最大行
	private List<CellRangeAddress> rangeList;       // 合并单元格
	
	/**
	 * 得到sheet1名称
	 * @return
	 */
	public String getSheet1Name() {
		return sheet1Name;
	}
	
	/**
	 * 设置sheet1名称
	 * @param sheet1Name
	 */
	public void setSheet1Name(String sheet1Name) {
		this.sheet1Name = sheet1Name;
	}
	
	/**
	 * 得到sheet2名称
	 * @return
	 */
	public String getSheet2Name() {
		return sheet2Name;
	}
	/**
	 * 设置sheet2名称
	 * @param sheet2Name
	 */
	public void setSheet2Name(String sheet2Name) {
		this.sheet2Name = sheet2Name;
	}
	
	/**
	 * 得到表头
	 * @return
	 */
	public String[] getHeaderArr() {
		return headerArr;
	}
	/**
	 * 设置表头
	 * @param headerArr
	 */
	public void setHeaderArr(String[] headerArr) {
		this.headerArr = headerArr;
	}
	
	/**
	 * 得到样例数据
	 * @return
	 */
	public ArrayList<String[]> getDemoList() {
		return demoList;
	}
	/**
	 * 设置样例数据
	 * @param demoList
	 */
	public void setDemoList(ArrayList<String[]> demoList) {
		this.demoList = demoList;
	}
	
	/**
	 * 得到下拉数据
	 * @return
	 */
	public ArrayList<String[]> getSelectDataList() {
		return selectDataList;
	}
	/**
	 * 设置下拉数据
	 * @param selectDataList
	 */
	public void setSelectDataList(ArrayList<String[]> selectDataList) {
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
