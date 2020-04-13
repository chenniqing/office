package cn.javaex.office.excel.help;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Workbook
 * 
 * @author 陈霓清
 */
public class WorkbookHelpler {
	/** 导出超过多少条数据时，使用SXSSFWorkbook */
	public static final int MAX_SIZE = 100000;
	
	/**
	 * 创建Workbook
	 * @param size 导出条数
	 * @return
	 */
	public Workbook createWorkbook(int size) {
		if (size > MAX_SIZE) {
			return new SXSSFWorkbook();
		}
		return new XSSFWorkbook();
	}
	
}
