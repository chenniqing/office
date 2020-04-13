package cn.javaex.office.excel.help;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import cn.javaex.office.util.PathUtils;

/**
 * Cell
 * 
 * @author 陈霓清
 */
public class CellHelper {
	
	/**
	 * 设置图片
	 * @param cell
	 * @param imagePath
	 * @throws IOException 
	 */
	public void setImage(Cell cell, String imagePath) throws IOException {
		cell.getRow().setHeight((short) 1000);
		if (imagePath==null || imagePath.length()==0) {
			cell.setCellValue("");
			return;
		}
		
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		// 文件后缀
		String fileSuffix = imagePath.substring(imagePath.lastIndexOf(".") + 1).toLowerCase();
		// 传入的路径是否是绝对路径
		boolean isAbsolutePath = PathUtils.isAbsolutePath(imagePath);
		// 存储图片的物理路径
		String filePath = "";
		if (isAbsolutePath) {
			filePath = imagePath;
		} else {
			String projectPath = PathUtils.getProjectPath();
			filePath = projectPath + File.separator + imagePath;
		}
		
		BufferedImage bufferImg = ImageIO.read(new File(filePath));
		ImageIO.write(bufferImg, fileSuffix, byteArrayOut);
		
		Drawing<?> patriarch = cell.getSheet().createDrawingPatriarch();
		ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, (short) cell.getColumnIndex(), cell.getRow().getRowNum(), (short) (cell.getColumnIndex() + 1), cell.getRow().getRowNum() + 1);
		
		int imageType = Workbook.PICTURE_TYPE_JPEG;
		if ("png".equals(fileSuffix)) {
			imageType = Workbook.PICTURE_TYPE_PNG;
		}
		patriarch.createPicture(anchor, cell.getSheet().getWorkbook().addPicture(byteArrayOut.toByteArray(), imageType));
	}

}
