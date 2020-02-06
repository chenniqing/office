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
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.javaex.office.excel.entity.ExcelSetting;

/**
 * Excel工具类
 * 
 * @author 陈霓清
 */
public class ExcelUtils {
	
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
	private static DataValidation setDataValidation(XSSFSheet sheet, String formula,
			int startRow, int endRow, int startCol, int endCol) {
		
		// 设置数据有效性加载在哪个单元格上。四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList addressList = new CellRangeAddressList(startRow, endRow, startCol, endCol);
		
		XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		XSSFDataValidationConstraint constraint = (XSSFDataValidationConstraint)dvHelper.createFormulaListConstraint(formula);
		XSSFDataValidation dataValidation = (XSSFDataValidation)dvHelper.createValidation(constraint, addressList);
		
		return dataValidation;
	}
	
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
	 * 得到Excel对象
	 * @param excelSetting
	 * @throws Exception
	 */
	public static XSSFWorkbook getXlsx(ExcelSetting excelSetting) throws Exception {
		// sheet1名称
		String sheet1Name = excelSetting.getSheet1Name();
		if (sheet1Name==null || sheet1Name.length()==0) {
			sheet1Name = "Sheet1";
		}
		// sheet2名称
		String sheet2Name = excelSetting.getSheet2Name();
		if (sheet2Name==null || sheet2Name.length()==0) {
			sheet2Name = "Sheet2";
		}
		// 表头
		String[] headerArr = excelSetting.getHeaderArr();
		// 样例数据
		ArrayList<String[]> demoList = excelSetting.getDemoList();
		// 下拉数据
		ArrayList<String[]> selectDataList = excelSetting.getSelectDataList();
		// 指定sheet1中需要下拉的列
		String[] selectColArr = excelSetting.getSelectColArr();
		// 列宽
		Integer columnWidth = excelSetting.getColumnWidth();
		// 下拉数据来源作用于sheet1的最大行
		Integer maxRow = excelSetting.getMaxRow();
		if (maxRow==null || maxRow==0) {
			maxRow = 5000;
		}
		
		// 创建工作薄
		XSSFWorkbook xwb = new XSSFWorkbook();
		
		// 设置字体样式
		XSSFCellStyle cellStyle = xwb.createCellStyle();
		// 水平对齐方式（居中）
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		// 垂直对齐方式（居中）
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		// 设置单元格为文本格式
		XSSFDataFormat format = xwb.createDataFormat();
		cellStyle.setDataFormat(format.getFormat("@"));
		
		// 新建sheet及名称
		XSSFSheet sheet1 = xwb.createSheet(sheet1Name);
		// 设置单元格属性
		for (int i=0; i<26; i++) {
			sheet1.setDefaultColumnStyle(i, cellStyle);
		}
		
		// sheet的第一行为标题
		XSSFRow firstRow = sheet1.createRow(0);
		if (headerArr!=null && headerArr.length>0) {
			for (int i=0; i<headerArr.length; i++) {
				// 获取第一行的每个单元格
				XSSFCell cell = firstRow.createCell(i);
				// 设置列宽
				if (columnWidth!=null && columnWidth>0) {
					sheet1.setColumnWidth(i, columnWidth);
				}
				// 往单元格里写数据
				cell.setCellValue(headerArr[i]);
				
				// 设置字体样式
				XSSFCellStyle cellStyleHeader = xwb.createCellStyle();
				// 水平对齐方式（居中）
				cellStyleHeader.setAlignment(HorizontalAlignment.CENTER);
				// 垂直对齐方式（居中）
				cellStyleHeader.setVerticalAlignment(VerticalAlignment.CENTER);
				// 字体
				XSSFFont fontHeader = xwb.createFont();
				fontHeader.setFontName("等线");
				fontHeader.setFontHeightInPoints((short) 14);
				cellStyleHeader.setFont(fontHeader);
				
				// 将样式设置应用具体单元格
				cell.setCellStyle(cellStyleHeader);
			}
		}
		
		// 写数据
		if (demoList!=null && demoList.isEmpty()==false) {
			int len = demoList.size();
			for (int i=0; i<len; i++) {
				String[] data = demoList.get(i);
				if (data!=null && data.length>0) {
					XSSFRow row = sheet1.createRow(i+1);
					
					for (int j=0; j<data.length; j++) {
						// 获取行的每个单元格
						XSSFCell cell = row.createCell(j);
						// 往单元格里写数据
						cell.setCellValue(data[j]);
						// 将样式设置应用具体单元格
						cell.setCellStyle(cellStyle);
					}
				}
			}
		}
		
		// 设置合并单元格
		List<CellRangeAddress> rangeList = excelSetting.getRangeList();
		if (rangeList!=null && rangeList.isEmpty()==false) {
			for (CellRangeAddress cellRangeAddress : rangeList) {
				sheet1.addMergedRegion(cellRangeAddress);
			}
		}
		
		// 设置下拉框数据
		if (selectColArr!=null && selectColArr.length>0) {
			XSSFSheet sheet2 = xwb.createSheet(sheet2Name);
			
			String[] arr = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
			int index = 0;
			for (int i=0; i<selectColArr.length; i++) {
				// 获取下拉对象
				String[] dlData = selectDataList.get(i);
				int colNum = Integer.parseInt(selectColArr[i]);
				
				// 设置有效性
				String formula = "Sheet2!$"+arr[index]+"$1:$"+arr[index]+"$"+dlData.length; //Sheet2第A1到A500作为下拉列表来源数据
				// 设置数据有效性加载在哪个单元格上,参数分别是：从sheet2获取A1到AmaxRow作为一个下拉的数据、起始行、终止行、起始列、终止列
				sheet1.addValidationData(setDataValidation(sheet2, formula, 1, maxRow, colNum, colNum));
				
				// 生成sheet2内容
				for (int j=0; j<dlData.length; j++) {
					if (index==0) {
						// 第1个下拉选项，直接创建行、列，设置对应单元格的值
						sheet2.createRow(j).createCell(0).setCellValue(dlData[j]);
					} else {
						// 非第1个下拉选项
						int colCount = sheet2.getLastRowNum();
						if (j<=colCount) {
							// 前面创建过的行，直接获取行，创建列，设置对应单元格的值
							sheet2.getRow(j).createCell(index).setCellValue(dlData[j]);
						} else {
							// 未创建过的行，直接创建行、创建列，设置对应单元格的值
							sheet2.createRow(j).createCell(index).setCellValue(dlData[j]);
						}
					}
				}
				
				index++;
			}
		}
		
		return xwb;
	}

	/**
	 * 读取将Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz 自定义实体类
	 * @param num   前几行跳过。从1开始计，0表示不跳过
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readXlsx(InputStream inputStream, Class<T> clazz, int num) throws Exception {
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
				if ("".equals(cellValue)) {
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
	 * 读取将Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz 自定义实体类
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readXlsx(InputStream inputStream, Class<T> clazz) throws Exception {
		return readXlsx(inputStream, clazz, 1);
	}

}
