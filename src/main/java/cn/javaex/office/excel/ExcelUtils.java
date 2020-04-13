package cn.javaex.office.excel;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import cn.javaex.office.excel.annotation.ExcelSheet;
import cn.javaex.office.excel.entity.ExcelSetting;
import cn.javaex.office.excel.help.SheetHelper;
import cn.javaex.office.excel.help.WorkbookHelpler;

/**
 * Excel工具类
 * 
 * @author 陈霓清
 */
public class ExcelUtils {
	
	/**
	 * 获取单元格内容
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		if (cell==null) {
			return "";
		}
		
		String cellValue = "";
		
		switch (cell.getCellType()) {
			case STRING :
				cellValue = cell.getRichStringCellValue().getString().trim();
				break;
			case NUMERIC :
				// 判断是否为日期类型
				if (DateUtil.isCellDateFormatted(cell)) {
					// 用于转化为日期格式
					Date date = cell.getDateCellValue();
					DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
					cellValue = formater.format(date);
				} else {
					// 格式化数字
					if (cell.toString().endsWith(".0")) {
						DecimalFormat df = new DecimalFormat("#");
						cellValue = df.format(cell.getNumericCellValue()).toString();
					} else {
						cellValue = String.valueOf(cell.getNumericCellValue());
					}
				}
				break;
			case BOOLEAN :
				cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
				break;
			case FORMULA :
				cellValue = cell.getCellFormula();
				break;
			case BLANK :
				cellValue = "";
				break;
			case ERROR :
				cellValue = "";
				break;
			default :
				cellValue = "";
		}
		return cellValue;
	}
	
	/**
	 * 根据注解方式得到Workbook对象
	 * @param clazz
	 * @param list
	 * @return
	 * @throws Exception
	 */
	public static Workbook getExcel(Class<?> clazz, List<?> list) throws Exception {
		// 设置sheet名称
		String sheetName = SheetHelper.SHEET_NAME;
		ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
		if (excelSheet!=null) {
			sheetName = excelSheet.name();
		}
		
		// 得到Workbook对象
		return getExcel(null, clazz, list, sheetName);
	}
	
	/**
	 * 根据注解方式得到Workbook对象（手动指定sheet名称）
	 * @param wb
	 * @param clazz
	 * @param list
	 * @param sheetName
	 * @return
	 * @throws Exception
	 */
	public static Workbook getExcel(Workbook wb, Class<?> clazz, List<?> list, String sheetName) throws Exception {
		SheetHelper sheetHelper = new SheetHelper();
		
		// 1.0 创建 Excel
		if (wb==null) {
			int size = list==null ? 0 : list.size();
			wb = new WorkbookHelpler().createWorkbook(size);
		}
		
		// 2.0 创建sheet
		Sheet sheet = wb.createSheet(sheetName);
		
		// 3.0 设置表头
		sheetHelper.createHeader(sheet, clazz);
		
		// 4.0 设置数据体
		sheetHelper.createData(sheet, clazz, list);
		
		return wb;
	}
	
	/**
	 * 得到Workbook对象
	 * @param excelSetting
	 * @throws Exception
	 */
	public static Workbook getExcel(ExcelSetting excelSetting) throws Exception {
		SheetHelper sheetHelper = new SheetHelper();
		
		// 1.0 创建 Excel
		List<String[]> dataList = excelSetting.getDataList();
		int size = dataList==null ? 0 : dataList.size();
		Workbook wb = new WorkbookHelpler().createWorkbook(size);
		
		// 2.0 创建sheet
		String sheetName = excelSetting.getSheetName();
		if (sheetName==null || sheetName.length()==0) {
			sheetName = SheetHelper.SHEET_NAME;
		}
		Sheet sheet = wb.createSheet(sheetName);
		
		// 3.0 设置表头
		sheetHelper.createHeader(sheet, excelSetting);
		
		// 4.0 设置数据体
		sheetHelper.createData(sheet, excelSetting);
		
		// 5.0 设置合并单元格
		sheetHelper.mergeCell(sheet, excelSetting);
		
		// 6.0 设置下拉数据
		sheetHelper.createSelectSheet(sheet, excelSetting);
		
		return wb;
	}

	/**
	 * 读取将Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz 自定义实体类
	 * @param num   前几行跳过。从1开始计，0表示不跳过
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream inputStream, Class<T> clazz, int num) throws Exception {
		List<T> list = new ArrayList<T>();
		
		Workbook wb = WorkbookFactory.create(inputStream);
		Sheet sheet = wb.getSheetAt(0);
		
		Field[] fieldArr = clazz.getDeclaredFields();
		
		for (Row row : sheet) {
			// 跳过第一行的表头
			if (row.getRowNum()<num) {
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
				String cellValue = getCellValue(cell);
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
				} catch (Exception e) {
					throw new Exception("导入实体类成员变量的数量与Excel中的字段数量不匹配");
				}
				Class<?> fieldType = field.getType();
				if (fieldType==String.class) {
					field.set(entity, cellValue);
				}
				else if (fieldType==Integer.class || fieldType==Integer.TYPE) {
					field.set(entity, Integer.parseInt(cellValue));
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
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz 自定义实体类
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream inputStream, Class<T> clazz) throws Exception {
		return readExcel(inputStream, clazz, 1);
	}
	
}
