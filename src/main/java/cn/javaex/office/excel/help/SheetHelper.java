package cn.javaex.office.excel.help;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import cn.javaex.office.excel.entity.ExcelSetting;

/**
 * Sheet操作
 * 
 * @author 陈霓清
 */
public class SheetHelper {
	
	/** 默认sheet页名称 */
	public static final String SHEET_NAME = "Sheet1";
	/** 行高基数 */
	public static final int BASE_ROW_HEIGHT = 20;
	/** 列宽基数 */
	public static final int BASE_COLUMN_WIDTH = 256;
	
	// 存储值替换
	public Map<String, Object> replaceMap = new HashMap<String, Object>();
	// 存储格式化
	public Map<String, Object> formatMap = new HashMap<String, Object>();
	// 存储合并多个单元格数据的成员变量
	public Map<String, String> skipMap = new HashMap<String, String>();
	
	/**
	 * 创建内容
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param title
	 * @throws Exception 
	 */
	public void exportExcel(Sheet sheet, Class<?> clazz, List<?> list, String title) throws Exception {
		
	}

	/**
	 * 创建内容
	 * @param sheet
	 * @param excelSetting
	 */
	public void exportExcel(Sheet sheet, ExcelSetting excelSetting) {
		
	}

	/**
	 * 读取sheet
	 * @param <T>
	 * @param sheet
	 * @param clazz     自定义实体类
	 * @param rowNum    从第几行开始读取（从0开始计算）
	 * @return
	 * @throws Exception 
	 */
	public <T> List<T> readSheet(Sheet sheet, Class<T> clazz, int rowNum) throws Exception {
		return null;
	}
	
	/**
	 * 设置下拉选项
	 * @param sheet
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataList  下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public void setSelect(Sheet sheet, int colNum, int startRow, int endRow, String[] selectDataList) {
		int colIndex = colNum - 1;
		int startRowIndex = startRow - 1;
		int endRowIndex = endRow - 1;
		
		// 获取单元格样式
		CellStyle cellStyle = null;
		try {
			// 获取第一个单元格的样式，用于继承
			cellStyle = sheet.getRow(startRowIndex).getCell(colIndex).getCellStyle();
		} catch (Exception e) {
			// 如果没有该单元格存在，则使用默认的样式
		}
		
		Row row = null;
		Cell cell = null;
		
		for (int i=startRowIndex; i<=endRowIndex; i++) {
			row = sheet.getRow(i);
			if (row==null) {
				cell = sheet.createRow(i).createCell(colIndex);
			} else {
				cell = row.getCell(colIndex);
				if (cell==null) {
					cell = row.createCell(colIndex);
				}
			}
			
			if (cellStyle!=null) {
				cell.setCellStyle(cellStyle);
			}
		}
		
		// 下拉的数据、起始行、终止行、起始列、终止列
		CellRangeAddressList addressList = new CellRangeAddressList(startRowIndex, endRowIndex, colIndex, colIndex);
		
		// 生成下拉框内容
		DataValidationHelper helper = sheet.getDataValidationHelper();
		DataValidationConstraint constraint = helper.createExplicitListConstraint(selectDataList); 
		DataValidation dataValidation = helper.createValidation(constraint, addressList);
		
		// 设置数据有效性
		sheet.addValidationData(dataValidation);
	}

	/**
	 * 设置合并
	 * @param wb
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算）
	 */
	public void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow-1, lastRow-1, firstCol-1, lastCol-1);
		sheet.addMergedRegion(cellRangeAddress);
	}

}
