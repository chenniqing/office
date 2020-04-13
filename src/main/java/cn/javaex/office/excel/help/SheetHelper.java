package cn.javaex.office.excel.help;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import cn.javaex.office.excel.annotation.ExcelCell;
import cn.javaex.office.excel.entity.ExcelSetting;
import cn.javaex.office.excel.style.DefaultCellStyle;

/**
 * Sheet
 * 
 * @author 陈霓清
 */
@SuppressWarnings("unchecked")
public class SheetHelper {
	
	/** 默认sheet页名称 */
	public static final String SHEET_NAME = "Sheet1";
	/** 列宽基数 */
	public static final int BASE_COLUMN_WIDTH = 256;
	/** 下拉数据来源作用于sheet的最大行 */
	public static final int MAX_SELECT_ROW = 5000;
	
	// 存储值替换
	private Map<String, Object> replaceMap = new HashMap<String, Object>();
	// 存储格式化
	private Map<String, Object> formatMap = new HashMap<String, Object>();

	/**
	 * 设置头部
	 * @param sheet
	 * @param clazz
	 * @return 
	 */
	public void createHeader(Sheet sheet, Class<?> clazz) {
		Row row = sheet.createRow(0);
		
		CellStyle cellStyle = new DefaultCellStyle().createHeaderStyle(sheet.getWorkbook());
		
		int index = 0;    // 列索引
		// 得到该类的所有成员变量
		Field[] fieldArr = clazz.getDeclaredFields();
		for (Field field : fieldArr) {
			// 设置类的私有属性可访问
			field.setAccessible(true);
			
			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			
			int sort = excelCell.sort()<0 ? index : excelCell.sort();
			
			// 设置列宽
			sheet.setColumnWidth(sort, excelCell.width() * BASE_COLUMN_WIDTH);
			
			// 设置单元格内容
			Cell cell = row.createCell(sort);
			cell.setCellValue(excelCell.name());
			
			// 设置单元格样式
			cell.setCellStyle(cellStyle);
			
			index++;
			
			// 设置值替换属性
			String[] replaceArr = excelCell.replace();
			if (replaceArr.length>0) {
				Map<String, String> map = new HashMap<String, String>();
				// {"1_男", "0_女"}
				for (String replace : replaceArr) {
					// 1_男
					String[] arr = replace.split("_");
					map.put(arr[0], arr[1]);
				}
				
				replaceMap.put(String.valueOf(sort), map);
			}
			// 设置格式化属性
			String format = excelCell.format();
			if (format.length()>0) {
				if (field.getType()==LocalDateTime.class || field.getType()==LocalDate.class) {
					DateTimeFormatter dtf = DateTimeFormatter.ofPattern(format);
					formatMap.put(String.valueOf(sort), dtf);
				}
				else if (field.getType()==Date.class) {
					SimpleDateFormat sdf = new SimpleDateFormat(format);
					formatMap.put(String.valueOf(sort), sdf);
				}
			}
		}
	}
	
	/**
	 * 设置头部
	 * @param sheet
	 * @param excelSetting
	 */
	public void createHeader(Sheet sheet, ExcelSetting excelSetting) {
		// 头部样式
		CellStyle cellStyle = excelSetting.getCellStyle().createHeaderStyle(sheet.getWorkbook());
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		
		if (headerList!=null && headerList.isEmpty()==false) {
			for (int i=0; i<headerList.size(); i++) {
				// 创建行
				Row row = sheet.createRow(i);
				
				String[] headerArr = headerList.get(i);
				for (int j=0; j<headerArr.length; j++) {
					// 创建单元格
					Cell cell = row.createCell(j);
					// 设置列宽
					sheet.setColumnWidth(j, excelSetting.getColumnWidth() * BASE_COLUMN_WIDTH);
					// 设置单元格数据
					cell.setCellValue(headerArr[j]);
					// 设置单元格样式
					cell.setCellStyle(cellStyle);
				}
			}
		}
	}
	
	/**
	 * 设置数据
	 * @param row
	 * @param clazz
	 * @param list
	 * @throws Exception 
	 */
	public void createData(Sheet sheet, Class<?> clazz, List<?> list) throws Exception {
		if (list==null || list.isEmpty()) {
			return;
		}
		
		CellStyle cellStyle = new DefaultCellStyle().createDataStyle(sheet.getWorkbook());
		
		Field[] fieldArr = clazz.getDeclaredFields();
		
		CellHelper cellHelper = new CellHelper();
		
		Row row = null;
		int rowIndex = 1;
		int len = list.size();
		for (int i=0; i<len; i++) {
			row = sheet.createRow(rowIndex);
			
			Object entity = list.get(i);
			
			int index = 0;
			for (Field field : fieldArr) {
				// 设置类的私有属性可访问
				field.setAccessible(true);
				
				// 得到每一个成员变量上的 ExcelCell 注解
				ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
				if (excelCell==null) {
					continue;
				}
				
				int sort = excelCell.sort()<0 ? index : excelCell.sort();
				
				// 创建单元格并设置值
				Cell cell = row.createCell(sort);
				Object obj = field.get(entity);
				
				if ("image".equals(excelCell.type())) {
					if (obj==null) {
						cell.setCellValue("");
					} else {
						cellHelper.setImage(cell, (String) obj);
					}
				} else {
					if (obj==null) {
						cell.setCellValue("");
					}
					else if (obj instanceof String) {
						cell.setCellValue((String) obj);
					}
					else if (obj instanceof Integer) {
						cell.setCellValue(Integer.parseInt(obj.toString()));
					}
					else if (obj instanceof Double) {
						cell.setCellValue(Double.parseDouble(obj.toString()));
					}
					else if (obj instanceof Long) {
						cell.setCellValue(Long.parseLong(obj.toString()));
					}
					else if (obj instanceof Float) {
						cell.setCellValue(Float.parseFloat(obj.toString()));
					}
					else if (obj instanceof BigDecimal) {
						cell.setCellValue(new BigDecimal(obj.toString()).doubleValue());
					}
					else if (obj instanceof LocalDateTime) {
						DateTimeFormatter dtf = (DateTimeFormatter) formatMap.get(String.valueOf(sort));
						cell.setCellValue(dtf.format((LocalDateTime) obj));
					}
					else if (obj instanceof LocalDate) {
						DateTimeFormatter dtf = (DateTimeFormatter) formatMap.get(String.valueOf(sort));
						cell.setCellValue(dtf.format((LocalDate) obj));
					}
					else if (obj instanceof Date) {
						SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(sort));
						cell.setCellValue(sdf.format((Date) obj));
					}
					else {
						cell.setCellValue(obj.toString());
					}
					
					// 值替换
					if (obj!=null && excelCell.replace().length>0) {
						Map<String, String> map = (Map<String, String>) replaceMap.get(String.valueOf(sort));
						if (map.get(obj.toString())!=null) {
							cell.setCellValue(map.get(obj.toString()));
						}
					}
				}
				
				// 设置单元格样式
				cell.setCellStyle(cellStyle);
				
				index++;
			}
			
			rowIndex++;
		}
	}

	/**
	 * 设置数据
	 * @param sheet
	 * @param excelSetting
	 */
	public void createData(Sheet sheet, ExcelSetting excelSetting) {
		int dataRowIndex = 0;    // 数据行的起始索引
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			dataRowIndex = headerList.size();
		}
		
		// 数据样式
		CellStyle cellStyle = excelSetting.getCellStyle().createDataStyle(sheet.getWorkbook());
		// 数据
		List<String[]> dataList = excelSetting.getDataList();
		
		if (dataList!=null && dataList.isEmpty()==false) {
			int len = dataList.size();
			for (int i=0; i<len; i++) {
				// 创建行
				Row row = sheet.createRow(i + dataRowIndex);
				
				// 得到每一行的数据
				String[] data = dataList.get(i);
				for (int j=0; j<data.length; j++) {
					// 创建单元格
					Cell cell = row.createCell(j);
					// 设置单元格数据
					cell.setCellValue(data[j]);
					// 设置单元格样式
					cell.setCellStyle(cellStyle);
				}
			}
		}
	}

	/**
	 * 合并单元格
	 * @param sheet
	 * @param excelSetting
	 */
	public void mergeCell(Sheet sheet, ExcelSetting excelSetting) {
		List<CellRangeAddress> rangeList = excelSetting.getRangeList();
		if (rangeList!=null && rangeList.isEmpty()==false) {
			for (CellRangeAddress cellRangeAddress : rangeList) {
				sheet.addMergedRegion(cellRangeAddress);
			}
		}
	}

	/**
	 * 创建下拉框sheet
	 * @param sheet
	 * @param excelSetting
	 */
	public void createSelectSheet(Sheet sheet, ExcelSetting excelSetting) {
		Workbook wb = sheet.getWorkbook();
		
		// 数据样式
		CellStyle cellStyle = excelSetting.getCellStyle().createDataStyle(wb);
		for (int i=0; i<26; i++) {
			sheet.setDefaultColumnStyle(i, cellStyle);
		}
		
		int dataRowIndex = 0;    // 数据行的起始索引
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			dataRowIndex = headerList.size();
		}
		
		// 下拉数据
		List<String[]> selectDataList = excelSetting.getSelectDataList();
		// 指定sheet中需要下拉的列
		String[] selectColArr = excelSetting.getSelectColArr();
		// 下拉数据来源作用于sheet的最大行
		Integer maxRow = excelSetting.getMaxRow();
		if (maxRow==null || maxRow==0) {
			maxRow = MAX_SELECT_ROW;
		}
		
		// 设置下拉框数据
		if (selectColArr!=null && selectColArr.length>0) {
			String selectSheetName = excelSetting.getSelectSheetName();
			Sheet selectSheet = wb.createSheet(selectSheetName);
			
			String[] arr = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
			int index = 0;
			for (int i=0; i<selectColArr.length; i++) {
				// 获取下拉对象
				String[] dlData = selectDataList.get(i);
				int colNum = Integer.parseInt(selectColArr[i]);
				
				// 设置有效性
				String formula = selectSheetName + "!$"+arr[index]+"$1:$"+arr[index]+"$"+dlData.length;
				// 设置数据有效性加载在哪个单元格上,参数分别是：从selectSheet获取A1到AmaxRow作为一个下拉的数据、起始行、终止行、起始列、终止列
				sheet.addValidationData(this.setDataValidation(selectSheet, formula, dataRowIndex, maxRow, colNum, colNum));
				
				// 生成selectSheet内容
				for (int j=0; j<dlData.length; j++) {
					if (index==0) {
						// 第1个下拉选项，直接创建行、列，设置对应单元格的值
						Cell createCell = selectSheet.createRow(j).createCell(0);
						createCell.setCellValue(dlData[j]);
					} else {
						// 非第1个下拉选项
						int colCount = selectSheet.getLastRowNum();
						if (j<=colCount) {
							// 前面创建过的行，直接获取行，创建列，设置对应单元格的值
							Cell createCell = selectSheet.getRow(j).createCell(index);
							createCell.setCellValue(dlData[j]);
						} else {
							// 未创建过的行，直接创建行、创建列，设置对应单元格的值
							Cell createCell = selectSheet.createRow(j).createCell(index);
							createCell.setCellValue(dlData[j]);
						}
					}
				}
				
				index++;
			}
		}
	}
	
	/**
	 * 设置数据有效性
	 * @param sheet
	 * @param formula
	 * @param startRow 起始行
	 * @param endRow 终止行
	 * @param startCol 起始列
	 * @param endCol 终止列
	 * @return
	 */
	private DataValidation setDataValidation(Sheet sheet, String formula,
			int startRow, int endRow, int startCol, int endCol) {
		
		// 设置数据有效性加载在哪个单元格上。四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList addressList = new CellRangeAddressList(startRow, endRow, startCol, endCol);
		
		DataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
		DataValidationConstraint constraint = dvHelper.createFormulaListConstraint(formula);
		DataValidation dataValidation = dvHelper.createValidation(constraint, addressList);
		
		return dataValidation;
	}
	
}
