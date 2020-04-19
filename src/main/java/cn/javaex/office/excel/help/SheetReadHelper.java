package cn.javaex.office.excel.help;

import java.lang.reflect.Field;
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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import cn.javaex.office.excel.ExcelUtils;
import cn.javaex.office.excel.annotation.ExcelCell;

/**
 * 读取Excel
 * 
 * @author 陈霓清
 */
public class SheetReadHelper extends SheetHelper {

	/**
	 * 读取sheet
	 * @param <T>
	 * @param sheet
	 * @param clazz     自定义实体类
	 * @param rowNum    从第几行开始读取（从0开始计算）
	 * @return
	 * @throws Exception 
	 */
	@SuppressWarnings("unchecked")
	@Override
	public <T> List<T> readSheet(Sheet sheet, Class<T> clazz, int rowNum) throws Exception {
		List<T> list = new ArrayList<T>();
		
		Field[] fieldArr = clazz.getDeclaredFields();
		
		// 解析注解
		this.readAnnotation(fieldArr);
		
		// 遍历数据
		for (Row row : sheet) {
			// 跳过表头
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
	
}
