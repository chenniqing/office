package cn.javaex.office.excel.help;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import cn.javaex.office.excel.ExcelUtils;
import cn.javaex.office.excel.annotation.ExcelCell;
import cn.javaex.office.excel.annotation.ExcelStyle;
import cn.javaex.office.excel.entity.ExcelSetting;
import cn.javaex.office.excel.style.DefaultCellStyle;
import cn.javaex.office.excel.style.ICellStyle;

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
	 * 设置标题
	 * @param sheet
	 * @param clazz
	 * @param sheetTitle 
	 * @return        返回当前写到第几行
	 * @throws Exception 
	 */
	public int createTtile(Sheet sheet, Class<?> clazz, String sheetTitle) throws Exception {
		Row row = sheet.createRow(0);
		
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createTitleStyle(sheet.getWorkbook());
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.value()).newInstance();
			cellStyle = obj.createTitleStyle(sheet.getWorkbook());
		}
		
		Cell cell = row.createCell(0);
		// 设置单元格内容
		cell.setCellValue(sheetTitle);
		// 设置单元格样式
		cell.setCellStyle(cellStyle);
		
		// 得到该类的所有成员变量，计算得到需要合并的单元格
		int length = 0;
		Field[] declaredFields = clazz.getDeclaredFields();
		for (Field field : declaredFields) {
			// 设置类的私有属性可访问
			field.setAccessible(true);
			
			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			
			length++;
		}
		
		// 设置合并
		// 四个参数分别是：起始行、终止行、起始列、终止列（从0计）
		CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, length-1);
		sheet.addMergedRegion(cellRangeAddress);
		
		return 1;
	}
	
	/**
	 * 设置标题
	 * @param sheet
	 * @param excelSetting
	 * @return
	 */
	public int createTtile(Sheet sheet, ExcelSetting excelSetting) {
		String title = excelSetting.getTitle();
		if (title==null || title.length()==0) {
			return 0;
		}
		
		Row row = sheet.createRow(0);
		
		// 标题样式
		CellStyle cellStyle = excelSetting.getCellStyle().createTitleStyle(sheet.getWorkbook());
		
		Cell cell = row.createCell(0);
		// 设置单元格内容
		cell.setCellValue(title);
		// 设置单元格样式
		cell.setCellStyle(cellStyle);
		
		int length = 0;
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			length = headerList.size();
		}
		
		// 设置合并
		// 四个参数分别是：起始行、终止行、起始列、终止列（从0计）
		CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, length-1);
		sheet.addMergedRegion(cellRangeAddress);
		
		return 1;
	}
	
	/**
	 * 设置头部
	 * @param sheet
	 * @param clazz
	 * @return 
	 * @throws Exception 
	 */
	public void createHeader(Sheet sheet, Class<?> clazz) throws Exception {
		this.createHeader(sheet, clazz, 0);
	}
	
	/**
	 * 设置头部
	 * @param sheet
	 * @param clazz
	 * @param rowIndex
	 * @return        返回当前写到第几行
	 * @throws Exception 
	 */
	public int createHeader(Sheet sheet, Class<?> clazz, int rowIndex) throws Exception {
		Row row = sheet.createRow(rowIndex);
		
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createHeaderStyle(sheet.getWorkbook());
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.value()).newInstance();
			cellStyle = obj.createHeaderStyle(sheet.getWorkbook());
		}
		
		int colIndex = 0;    // 列索引
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
			
			int sort = excelCell.sort()<0 ? colIndex : excelCell.sort();
			
			// 设置列宽
			sheet.setColumnWidth(sort, excelCell.width() * BASE_COLUMN_WIDTH);
			
			Cell cell = row.createCell(sort);
			// 设置单元格内容
			cell.setCellValue(excelCell.name());
			// 设置单元格样式
			cell.setCellStyle(cellStyle);
			
			colIndex++;
			
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
		
		return ++rowIndex;
	}
	
	/**
	 * 设置头部
	 * @param sheet
	 * @param excelSetting
	 * @return 
	 */
	public int createHeader(Sheet sheet, ExcelSetting excelSetting) {
		int rowIndex = 0;
		String title = excelSetting.getTitle();
		if (title!=null && title.length()>0) {
			rowIndex = 1;
		}
		
		int headerLen = 0;
		// 头部样式
		CellStyle cellStyle = excelSetting.getCellStyle().createHeaderStyle(sheet.getWorkbook());
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			headerLen = headerList.size();
			
			for (int i=rowIndex; i<headerLen; i++) {
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
		
		return rowIndex + headerLen;
	}
	
	/**
	 * 设置数据
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @throws Exception
	 */
	public void createData(Sheet sheet, Class<?> clazz, List<?> list) throws Exception {
		this.createData(sheet, clazz, list, 1);
	}

	/**
	 * 设置数据
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param rowIndex
	 */
	public void createData(Sheet sheet, Class<?> clazz, List<?> list, int rowIndex) throws Exception {
		if (list==null || list.isEmpty()) {
			return;
		}
		
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createDataStyle(sheet.getWorkbook());
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.value()).newInstance();
			cellStyle = obj.createDataStyle(sheet.getWorkbook());
		}
		
		Field[] fieldArr = clazz.getDeclaredFields();
		
		CellHelper cellHelper = new CellHelper();
		
		Row row = null;
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
		
		String title = excelSetting.getTitle();
		if (title!=null && title.length()>0) {
			dataRowIndex += 1;
		}
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			dataRowIndex += headerList.size();
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
	 * 读取sheet
	 * @param <T>
	 * @param sheet
	 * @param clazz     自定义实体类
	 * @param rowNum    从第几行开始读取（从0开始计算）
	 * @return
	 * @throws Exception 
	 */
	public <T> List<T> readSheet(Sheet sheet, Class<T> clazz, int rowNum) throws Exception {
		List<T> list = new ArrayList<T>();
		
		Field[] fieldArr = clazz.getDeclaredFields();
		
		// 解析注解
		readAnnotation(fieldArr);
		
		// 遍历数据
		for (Row row : sheet) {
			// 跳过第一行的表头
			if (row.getRowNum()<rowNum) {
				continue;
			}
			
			// 遍历每一列
			T entity = null;
			int len = fieldArr.length;
			for (int i=0; i<len; i++) {
				Cell cell = row.getCell(i);
				if (cell==null) {
					continue;
				}
				
				// 获取该列的值
				String cellValue = ExcelUtils.getCellValue(cell);
				if (cellValue.length()==0) {
					continue;
				}
				// 如果实例不存在则新建
				if (entity==null) {
					entity = clazz.newInstance();
				}
				
				// 根据对象类型设置值
				Field field = null;
				try {
					field = fieldArr[i];
					field.setAccessible(true);    // 设置类的私有属性可访问
					
					ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
					if (excelCell!=null) {
						// 值替换
						if (excelCell.replace().length>0) {
							Map<String, String> map = (Map<String, String>) replaceMap.get(String.valueOf(i));
							if (map.get((String) cellValue)!=null) {
								cellValue = map.get(cellValue.toString());
							}
						}
					}
				} catch (Exception e) {
					e.printStackTrace();
					throw new Exception("导入实体类成员变量的数量与Excel中的字段数量不匹配");
				}
				Class<?> fieldType = field.getType();
				if (fieldType==String.class) {
					field.set(entity, cellValue);
				}
				else if (fieldType==Integer.class || fieldType==Integer.TYPE) {
					field.set(entity, Double.valueOf(String.valueOf(cellValue)).intValue());
				}
				else if (fieldType==Long.class || fieldType==Long.TYPE) {
					field.set(entity, Long.valueOf(cellValue));
				}
				else if (fieldType==Double.class || fieldType==Double.TYPE) {
					field.set(entity, Double.valueOf(cellValue));
				}
				else if (fieldType==Float.class || fieldType==Float.TYPE) {
					field.set(entity, Float.valueOf(cellValue));
				}
				else if (fieldType==LocalDateTime.class) {
					SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(i));
					if (sdf!=null) {
						try {
							Instant instant = sdf.parse(cellValue).toInstant();
							ZoneId zone = ZoneId.systemDefault();
							LocalDateTime localDateTime = LocalDateTime.ofInstant(instant, zone);
							
							field.set(entity, localDateTime);
						} catch (Exception e) {
							field.set(entity, null);
						}
					}
				}
				else if (fieldType==LocalDate.class) {
					DateTimeFormatter dtf = (DateTimeFormatter) formatMap.get(String.valueOf(i));
					if (dtf!=null) {
						try {
							LocalDate ld = LocalDate.parse(cellValue, dtf);
							field.set(entity, ld);
						} catch (Exception e) {
							field.set(entity, null);
						}
					}
				}
				else if (fieldType==Date.class) {
					SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(i));
					if (sdf!=null) {
						try {
							field.set(entity, sdf.parse(cellValue));
						} catch (Exception e) {
							field.set(entity, null);
						}
					}
				}
				else {
					field.set(entity, null);
				}
			}
			
			// 把每一行的实体对象加入list
			if (entity!=null) {
				list.add(entity);
			}
		}
		
		return list;
	}

	/**
	 * 解析注解
	 * @param <T>
	 * @param fieldArr
	 */
	private void readAnnotation(Field[] fieldArr) {
		for (int i=0; i<fieldArr.length; i++) {
			Field field = fieldArr[i];
			// 设置类的私有属性可访问
			field.setAccessible(true);
			
			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			
			int sort = excelCell.sort()<0 ? i : excelCell.sort();
			
			// 设置值替换属性
			String[] replaceArr = excelCell.replace();
			if (replaceArr.length>0) {
				Map<String, String> map = new HashMap<String, String>();
				// {"男_1", "女_0"}
				for (String replace : replaceArr) {
					// 男_1
					String[] arr = replace.split("_");
					map.put(arr[0], arr[1]);
				}
				
				replaceMap.put(String.valueOf(sort), map);
			}
			
			// 设置格式化属性
			String format = excelCell.format();
			if (format.length()>0) {
				if (field.getType()==LocalDateTime.class || field.getType()==Date.class) {
					SimpleDateFormat sdf = new SimpleDateFormat(format);
					formatMap.put(String.valueOf(sort), sdf);
				}
				else if (field.getType()==LocalDate.class) {
					DateTimeFormatter dtf = DateTimeFormatter.ofPattern(format);
					formatMap.put(String.valueOf(sort), dtf);
				}
			}
		}
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
	 * @param firstRow    起始行（从0计）
	 * @param lastRow     终止行（从0计）
	 * @param firstCol    起始列（从0计）
	 * @param lastCol     终止列（从0计）
	 */
	public void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		sheet.addMergedRegion(cellRangeAddress);
	}

}
