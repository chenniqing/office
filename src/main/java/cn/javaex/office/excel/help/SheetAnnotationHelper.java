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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import cn.javaex.office.excel.annotation.ExcelCell;
import cn.javaex.office.excel.annotation.ExcelStyle;
import cn.javaex.office.excel.style.DefaultCellStyle;
import cn.javaex.office.excel.style.ICellStyle;

/**
 * 注解法导出Excel
 * 
 * @author 陈霓清
 */
public class SheetAnnotationHelper extends SheetHelper {

	/**
	 * 创建内容
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param title
	 * @throws Exception
	 */
	@Override
	public void exportExcel(Sheet sheet, Class<?> clazz, List<?> list, String title) throws Exception {
		// 当前写到了第几行（从1开始计算）
		int rowNum = 0;
		
		// 1.0 设置标题
		if (title!=null && title.length()>0) {
			rowNum = this.createTtile(sheet, clazz, title);
		}
		
		// 2.0 设置表头
		rowNum = this.createHeader(sheet, clazz, rowNum);
		
		// 3.0 设置数据
		this.createData(sheet, clazz, list, rowNum);
	}
	
	/**
	 * 设置标题
	 * @param sheet
	 * @param clazz
	 * @param title
	 * @return        返回当前写到第几行
	 * @throws Exception 
	 */
	private int createTtile(Sheet sheet, Class<?> clazz, String title) throws Exception {
		Row row = sheet.createRow(0);
		
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createTitleStyle(sheet.getWorkbook());
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.cellStyle()).newInstance();
			cellStyle = obj.createTitleStyle(sheet.getWorkbook());
			
			// 行高
			int height = excelStyle.titleHeight();
			if (height>0) {
				row.setHeight((short) (height * BASE_ROW_HEIGHT));
			}
		}
		
		// 设置单元格
		Cell cell = row.createCell(0);
		cell.setCellValue(title);
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
		// 四个参数分别是：起始行、终止行、起始列、终止列（从0开始计算）
		CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, length-1);
		sheet.addMergedRegion(cellRangeAddress);
		
		return 1;
	}
	
	/**
	 * 设置头部
	 * @param sheet
	 * @param clazz
	 * @param rowIndex
	 * @return        返回当前写到第几行
	 * @throws Exception 
	 */
	private int createHeader(Sheet sheet, Class<?> clazz, int rowIndex) throws Exception {
		Row row = sheet.createRow(rowIndex);
		
		Workbook workbook = sheet.getWorkbook();
		
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createHeaderStyle(workbook);
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.cellStyle()).newInstance();
			cellStyle = obj.createHeaderStyle(workbook);
			
			// 行高
			int height = excelStyle.headerHeight();
			if (height>0) {
				row.setHeight((short) (height * BASE_ROW_HEIGHT));
			}
		}
		
		skipMap.clear();
		int colIndex = 0;    // 列索引
		// 得到该类的所有成员变量
		Field[] fieldArr = clazz.getDeclaredFields();
		for (int j=0; j<fieldArr.length; j++) {
			Field field = fieldArr[j];
			
			// 设置类的私有属性可访问
			field.setAccessible(true);
			
			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			// 跳过被归入组的列
			if (skipMap.get(field.getName())!=null) {
				continue;
			}
			
			int sort = excelCell.sort()<0 ? colIndex : excelCell.sort();
			
			// 设置列宽
			sheet.setColumnWidth(sort, excelCell.width() * BASE_COLUMN_WIDTH);
			
			// 设置单元格
			Cell cell = row.createCell(sort);
			cell.setCellValue(excelCell.name());
			cell.setCellStyle(cellStyle);
			
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
			
			int mergeCol = excelCell.group();
			if (mergeCol>1 && excelCell.sort()==-1) {
				int num = 0;
				for (int k=(j+1); k<fieldArr.length; k++) {
					Field temp = fieldArr[k];
					temp.setAccessible(true);
					if (field.getAnnotation(ExcelCell.class)==null) {
						continue;
					}
					
					skipMap.put(temp.getName(), temp.getName());
					
					num++;
					
					if (num==(mergeCol-1)) {
						break;
					}
				}
			}
			
			colIndex++;
		}
		
		return ++rowIndex;
	}
	
	/**
	 * 设置数据
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param rowIndex
	 */
	@SuppressWarnings("unchecked")
	public void createData(Sheet sheet, Class<?> clazz, List<?> list, int rowIndex) throws Exception {
		if (list==null || list.isEmpty()) {
			return;
		}
		
		// 行高
		int height = 0;
		
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createDataStyle(sheet.getWorkbook());
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.cellStyle()).newInstance();
			cellStyle = obj.createDataStyle(sheet.getWorkbook());
			
			// 行高
			height = excelStyle.titleHeight();
		}
		
		Field[] fieldArr = clazz.getDeclaredFields();
		
		CellHelper cellHelper = new CellHelper();
		
		Row row = null;
		int len = list.size();
		for (int i=0; i<len; i++) {
			row = sheet.createRow(rowIndex);
			
			// 行高
			if (height>0) {
				row.setHeight((short) (height * BASE_ROW_HEIGHT));
			}
			
			Object entity = list.get(i);
			
			skipMap.clear();
			int colIndex = 0;    // 列索引
			for (int j=0; j<fieldArr.length; j++) {
				Field field = fieldArr[j];
				
				// 设置类的私有属性可访问
				field.setAccessible(true);
				
				// 得到每一个成员变量上的 ExcelCell 注解
				ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
				if (excelCell==null) {
					continue;
				}
				// 跳过被归入组的列
				if (skipMap.get(field.getName())!=null) {
					continue;
				}
				
				int sort = excelCell.sort()<0 ? colIndex : excelCell.sort();
				
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
				
				int mergeCol = excelCell.group();
				if (mergeCol>1 && excelCell.sort()==-1) {
					String mergeStr = (String) obj;
					String separator = excelCell.separator();
					
					int num = 0;
					for (int k=(j+1); k<fieldArr.length; k++) {
						Field temp = fieldArr[k];
						temp.setAccessible(true);
						if (field.getAnnotation(ExcelCell.class)==null) {
							continue;
						}
						String str = (String) temp.get(entity);
						mergeStr = mergeStr + separator + str;
						
						skipMap.put(temp.getName(), temp.getName());
						
						num++;
						
						if (num==(mergeCol-1)) {
							break;
						}
					}
					
					cell.setCellValue(mergeStr);
				}
				
				colIndex++;
			}
			
			rowIndex++;
		}
	}
}
