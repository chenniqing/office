package cn.javaex.office.excel;

import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
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
			case FORMULA :    // 公式
				try {
					cellValue = String.valueOf(cell.getNumericCellValue());
				} catch (IllegalStateException e) {
					cellValue = String.valueOf(cell.getRichStringCellValue());
				} catch (Exception e) {
					cellValue = cell.getCellFormula();
				}
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
	 * 根据注解方式导出Excel
	 * @param clazz 数据库查询得到的vo实体对象
	 * @param list  数据库查询得到的vo实体对象的数据集合
	 * @return
	 * @throws Exception
	 */
	public static Workbook exportExcel(Class<?> clazz, List<?> list) throws Exception {
		// 设置sheet名称
		String sheetName = SheetHelper.SHEET_NAME;
		String sheetTitle = null;
		ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
		if (excelSheet!=null) {
			sheetName = excelSheet.name();
			sheetTitle = excelSheet.title();
		}
		
		return exportExcel(null, clazz, list, sheetName, sheetTitle);
	}
	
	/**
	 * 根据注解方式导出Excel（手动指定sheet页名称）
	 * @param wb         Workbook对象
	 * @param clazz      数据库查询得到的vo实体对象
	 * @param list       数据库查询得到的vo实体对象的数据集合
	 * @param sheetName  追加创建的sheet页名称
	 * @return
	 * @throws Exception
	 */
	public static Workbook exportExcel(Workbook wb, Class<?> clazz, List<?> list, String sheetName) throws Exception {
		return exportExcel(null, clazz, list, sheetName, null);
	}
	
	/**
	 * 根据注解方式导出Excel（手动指定Sheet页名称）
	 * @param wb          Workbook对象，允许为NULL
	 * @param clazz       数据库查询得到的vo实体对象
	 * @param list        数据库查询得到的vo实体对象的数据集合
	 * @param sheetName   追加创建的sheet页名称
	 * @param sheetTitle  追加创建的sheet页顶部标题
	 * @return
	 * @throws Exception
	 */
	public static Workbook exportExcel(Workbook wb, Class<?> clazz, List<?> list, String sheetName, String sheetTitle) throws Exception {
		SheetHelper sheetHelper = new SheetHelper();
		
		// 1.0 创建 Excel
		if (wb==null) {
			int size = list==null ? 0 : list.size();
			wb = new WorkbookHelpler().createWorkbook(size);
		}
		
		// 2.0 创建sheet
		Sheet sheet = wb.createSheet(sheetName);
		
		// 当前写到了第几行（从1开始计算）
		int rowNum = 0;
		
		// 3.0 设置标题
		if (sheetTitle!=null && sheetTitle.length()>0) {
			rowNum = sheetHelper.createTtile(sheet, clazz, sheetTitle);
		}
		
		// 4.0 设置表头
		rowNum = sheetHelper.createHeader(sheet, clazz, rowNum);
		
		// 5.0 设置数据体
		sheetHelper.createData(sheet, clazz, list, rowNum);
		
		return wb;
	}

	/**
	 * 导出Excel
	 * @param excelSetting
	 * @throws Exception
	 */
	public static Workbook exportExcel(ExcelSetting excelSetting) throws Exception {
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
		
		// 3.0 设置标题
		sheetHelper.createTtile(sheet, excelSetting);
		
		// 4.0 设置表头
		sheetHelper.createHeader(sheet, excelSetting);
		
		// 5.0 设置数据体
		sheetHelper.createData(sheet, excelSetting);
		
		return wb;
	}

	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz 自定义实体类
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream inputStream, Class<T> clazz) throws Exception {
		return readExcel(inputStream, clazz, 1, 1);
	}
	
	/**
	 * 读取将Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz    自定义实体类
	 * @param rowNum   从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream inputStream, Class<T> clazz, int rowNum) throws Exception {
		return readExcel(inputStream, clazz, 1, rowNum-1);
	}
	
	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz      自定义实体类
	 * @param sheetNum   读取第几个sheet页（从1开始计算）
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream inputStream, Class<T> clazz, int sheetNum, int rowNum) throws Exception {
		SheetHelper sheetHelper = new SheetHelper();
		
		Workbook wb = WorkbookFactory.create(inputStream);
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		
		return sheetHelper.readSheet(sheet, clazz, rowNum-1);
	}
	
	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param inputStream
	 * @param clazz      自定义实体类
	 * @param sheetName  读取哪一个sheet页，填写sheet页名称
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream inputStream, Class<T> clazz, String sheetName, int rowNum) throws Exception {
		SheetHelper sheetHelper = new SheetHelper();
		
		Workbook wb = WorkbookFactory.create(inputStream);
		Sheet sheet = wb.getSheet(sheetName);
		
		return sheetHelper.readSheet(sheet, clazz, rowNum-1);
	}

	/**
	 * 设置下拉选项
	 * @param wb
	 * @param sheetNum        第几个Sheet页（从1开始计算）
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataList  下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Workbook wb, int sheetNum, int colNum, int startRow, int endRow, String[] selectDataList) {
		SheetHelper sheetHelper = new SheetHelper();
		
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		
		sheetHelper.setSelect(sheet, colNum, startRow, endRow, selectDataList);
	}
	
	/**
	 * 设置下拉选项
	 * @param wb
	 * @param sheetName       Sheet页名称
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataList  下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Workbook wb, String sheetName, int colNum, int startRow, int endRow, String[] selectDataList) {
		SheetHelper sheetHelper = new SheetHelper();
		
		Sheet sheet = wb.getSheet(sheetName);
		
		sheetHelper.setSelect(sheet, colNum, startRow, endRow, selectDataList);
	}

	/**
	 * 设置下拉选项
	 * @param sheet
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataList  下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Sheet sheet, int colNum, int startRow, int endRow, String[] selectDataList) {
		SheetHelper sheetHelper = new SheetHelper();
		
		sheetHelper.setSelect(sheet, colNum, startRow, endRow, selectDataList);
	}
	
	/**
	 * 设置合并单元格
	 * @param wb
	 * @param sheetNum    第几个Sheet页（从1开始计算）
	 * @param firstRow    起始行（从0计）
	 * @param lastRow     终止行（从0计）
	 * @param firstCol    起始列（从0计）
	 * @param lastCol     终止列（从0计）
	 */
	public static void setMerge(Workbook wb, int sheetNum, int firstRow, int lastRow, int firstCol, int lastCol) {
		SheetHelper sheetHelper = new SheetHelper();
		
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		
		sheetHelper.setMerge(sheet, firstRow, lastRow, firstCol, lastCol);
	}
	
	/**
	 * 设置合并单元格
	 * @param wb
	 * @param sheetName   Sheet页名称
	 * @param firstRow    起始行（从0计）
	 * @param lastRow     终止行（从0计）
	 * @param firstCol    起始列（从0计）
	 * @param lastCol     终止列（从0计）
	 */
	public static void setMerge(Workbook wb, String sheetName, int firstRow, int lastRow, int firstCol, int lastCol) {
		SheetHelper sheetHelper = new SheetHelper();
		
		Sheet sheet = wb.getSheet(sheetName);
		
		sheetHelper.setMerge(sheet, firstRow, lastRow, firstCol, lastCol);
	}
	
	/**
	 * 设置合并单元格
	 * @param sheet
	 * @param firstRow    起始行（从0计）
	 * @param lastRow     终止行（从0计）
	 * @param firstCol    起始列（从0计）
	 * @param lastCol     终止列（从0计）
	 */
	public static void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		SheetHelper sheetHelper = new SheetHelper();
		
		sheetHelper.setMerge(sheet, firstRow, lastRow, firstCol, lastCol);
	}

	/**
	 * 通过流读取Excel
	 * @param inputStream
	 * @return
	 * @throws Exception
	 */
	public static Workbook getExcel(InputStream inputStream) throws Exception {
		return WorkbookFactory.create(inputStream);
	}
	
	/**
	 * 读取resources文件夹下的Excel
	 * @param <T>
	 * @param t            直接写死 this
	 * @param filePath     resources文件夹下的路径，例如：template/excel/模板.xlsx
	 * @return
	 * @throws Exception
	 */
	public static <T> Workbook getExcelFromResource(T t, String filePath) throws Exception {
		if (filePath.startsWith("/")) {
			filePath = filePath.substring(1, filePath.length());
		}
		
		InputStream in = t.getClass().getClassLoader().getResourceAsStream(filePath);
		
		return WorkbookFactory.create(in);
	}
	
}
