package cn.javaex.office.word;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc.Enum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.w3c.dom.Document;

import cn.javaex.office.common.Font;
import cn.javaex.office.common.Picture;
import cn.javaex.office.common.Table;
import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;

/**
 * Word工具类
 * 
 * @author 陈霓清
 */
public class WordUtils {
	
	/**
	 * 根据图片类型，取得对应的图片类型代码
	 * @param imageType
	 * @return int
	 */
	private static int getImageType(String imageType) {
		int res = XWPFDocument.PICTURE_TYPE_JPEG;
		
		if (imageType!=null && imageType.length()>0) {
			if (imageType.equalsIgnoreCase("png")) {
				res = XWPFDocument.PICTURE_TYPE_PNG;
			}
			else if (imageType.equalsIgnoreCase("dib")) {
				res = XWPFDocument.PICTURE_TYPE_DIB;
			}
			else if (imageType.equalsIgnoreCase("emf")) {
				res = XWPFDocument.PICTURE_TYPE_EMF;
			}
			else if (imageType.equalsIgnoreCase("jpg") || imageType.equalsIgnoreCase("jpeg")) {
				res = XWPFDocument.PICTURE_TYPE_JPEG;
			}
			else if (imageType.equalsIgnoreCase("wmf")) {
				res = XWPFDocument.PICTURE_TYPE_WMF;
			}
		}
		
		return res;
	}
	
	/**
	 * 正则匹配字符串
	 * @param str
	 * @return
	 */
	private static Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
	
	/**
	 * 在表格指定位置新增一行
	 * @param table 需要插入数据的表格
	 * @param sourceRow 复制的源行
	 * @param rowIndex 表格指定位置
	 */
	private static void createRow(XWPFTable table, XWPFTableRow sourceRow, int rowIndex) {
		// 在表格指定位置新增一行
		XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);
		
		// 复制行属性
		targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
		
		// 复制列属性
		List<XWPFTableCell> cellList = sourceRow.getTableCells();
		if (cellList!=null && cellList.isEmpty()==false) {
			XWPFTableCell targetCell = null;
			for (XWPFTableCell sourceCell : cellList) {
				targetCell = targetRow.addNewTableCell();
				targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
			}
		}
	}
	
	/**
	 * 获取单元格水平对齐方式
	 * @param cell
	 * @return
	 */
	private static String getCellTextAlign(XWPFTableCell cell) {
		CTTc cttc = cell.getCTTc();
		CTP ctp = cttc.getPList().get(0);
		CTPPr ctppr = ctp.getPPr();
		if (ctppr==null) {
			ctppr = ctp.addNewPPr();
		}
		CTJc ctjc = ctppr.getJc();
		if (ctjc==null) {
			ctjc = ctppr.addNewJc();
		}
		
		Enum enumVal = ctjc.getVal();
		if (enumVal!=null) {
			return enumVal.toString().toLowerCase();
		}
		
		return null;
	}
	
	/**
	 * 设置单元格水平对齐方式
	 * @param cell
	 * @param cellTextAlign
	 */
	private static void setCellTextAlign(XWPFTableCell cell, String textAlign) {
		if (textAlign!=null && textAlign.length()>0) {
			textAlign = textAlign.toLowerCase();
			
			CTTc cttc = cell.getCTTc();
			CTP ctp = cttc.getPList().get(0);
			CTPPr ctppr = ctp.getPPr();
			if (ctppr==null) {
				ctppr = ctp.addNewPPr();
			}
			CTJc ctjc = ctppr.getJc();
			if (ctjc==null) {
				ctjc = ctppr.addNewJc();
			}
			
			if (textAlign.equals("left")) {
				ctjc.setVal(STJc.LEFT);
			}
			else if (textAlign.equals("center")) {
				ctjc.setVal(STJc.CENTER);
			}
			else if (textAlign.equals("right")) {
				ctjc.setVal(STJc.RIGHT);
			}
		}
	}
	
	/**
	 * 获取单元格垂直对齐方式
	 * @param cell
	 * @return
	 */
	private static String getCellVertAlign(XWPFTableCell cell) {
		XWPFVertAlign verticalAlignment = cell.getVerticalAlignment();
		if (verticalAlignment!=null) {
			return verticalAlignment.toString().toLowerCase();
		}
		
		return null;
	}
	
	/**
	 * 设置单元格垂直对齐方式
	 * @param cell
	 * @param vertAlign
	 */
	private static void setCellVertAlign(XWPFTableCell cell, String vertAlign) {
		if (vertAlign!=null && vertAlign.length()>0) {
			vertAlign = vertAlign.toLowerCase();
			
			if (vertAlign.equals("top")) {
				cell.setVerticalAlignment(XWPFVertAlign.TOP);
			}
			else if (vertAlign.equals("center")) {
				cell.setVerticalAlignment(XWPFVertAlign.CENTER);
			}
			else if (vertAlign.equals("both")) {
				cell.setVerticalAlignment(XWPFVertAlign.BOTH);
			}
			else if (vertAlign.equals("bottom")) {
				cell.setVerticalAlignment(XWPFVertAlign.BOTTOM);
			}
		}
	}
	
	/**
	 * 插入动态表格数据
	 * @param table 需要插入数据的表格
	 * @param dataList 插入数据集合
	 * @param index 表头行数/第一行数据行所在的索引位置
	 */
	private static void insertTable(XWPFTable table, List<String[]> dataList, int index) {
		// 创建行，根据需要插入的数据添加新行
		int len = dataList.size() - 1;
		for (int i=0; i<len; i++) {
			createRow(table, table.getRow(index), (i+1+index));
		}
		
		// 判断是否是自定义表格
		if (index==0) {
			// 根据每一行数据的数组大小，添加列数
			for (int i=0; i<dataList.size(); i++) {
				XWPFTableRow row = table.getRow(i);      // 每一行
				int cellSize = row.getTableCells().size();   // 每一行目前的列数
				int dataSize = dataList.get(i).length;       // 传入的List数据的每一个数据的数组大小
				
				// 给该行添加新列
				int colNum = dataSize - cellSize;
				if (colNum>0) {
					// 获取当前行最后一列的列属性
					XWPFTableCell sourceCell = row.getTableCells().get(cellSize-1);
					// 记录水平对齐方式
					String cellTextAlign = getCellTextAlign(sourceCell);
					// 记录垂直对齐方式
					String cellVertAlign = getCellVertAlign(sourceCell);
					
					for (int j=0; j<colNum; j++) {
						XWPFTableCell addCell = row.addNewTableCell();
						addCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
						
						// 设置水平对齐方式
						if (cellTextAlign!=null && cellTextAlign.length()>0) {
							setCellTextAlign(addCell, cellTextAlign);
						}
						// 设置垂直对齐方式
						if (cellVertAlign!=null && cellVertAlign.length()>0) {
							setCellVertAlign(addCell, cellVertAlign);
						}
					}
				}
			}
		}
		
		// 记录每一列单元格的水平对齐方式
		List<String> textAlignList = new ArrayList<String>();
		// 记录每一列单元格的垂直对齐方式
		List<String> vertAlignList = new ArrayList<String>();
		
		// 遍历表格插入数据
		len = table.getRows().size();
		for (int i=index; i<len; i++) {
			XWPFTableRow newRow = table.getRow(i);
			List<XWPFTableCell> cellList = newRow.getTableCells();
			for (int j=0; j<cellList.size(); j++) {
				XWPFTableCell cell = cellList.get(j);
				String text = null;
				try {
					text = dataList.get(i-index)[j];
				} catch (Exception e) {
					
				} finally {
					if (text==null) {
						text = "";
					}
					else if ("BLANK_LINE".equals(text)) {
						for (XWPFTableCell xwpfTableCell : cellList) {
							xwpfTableCell.getCTTc().getTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.NIL);
							xwpfTableCell.getCTTc().getTcPr().addNewTcBorders().addNewRight().setVal(STBorder.NIL);
						}
						for (XWPFParagraph paragraph : cell.getParagraphs()) {
							replaceParagraph(paragraph, "");
						}
						continue;
					}
				}
				
				XWPFParagraph addParagraph = cell.getParagraphs().get(0);
				XWPFRun createRun = addParagraph.createRun();
				
				// 需要插入的文本
				String insertText = text.replace("<br>", "<br/>");
				
				try {
					// 字体样式反序列化
					ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(text.getBytes("ISO-8859-1"));
					ObjectInputStream objectInputStream = new ObjectInputStream(byteArrayInputStream);
					Font fontStyle = (Font) objectInputStream.readObject();
					objectInputStream.close();
					byteArrayInputStream.close();
					
					// 设置字体样式
					setFontStyle(createRun, fontStyle);
					
					// 需要插入的文本
					insertText = fontStyle.getText();
				} catch (Exception e) {
					
				} finally {
					setText(createRun, insertText);
				}
				
				// 单元格水平、垂直对齐方式
				if (i==index) {
					// 记录水平对齐方式
					String cellTextAlign = getCellTextAlign(cell);
					
					if (cellTextAlign==null || cellTextAlign.length()==0) {
						textAlignList.add(null);
						setCellTextAlign(cell, "left");
					} else {
						textAlignList.add(cellTextAlign);
					}
					
					// 记录垂直对齐方式
					String cellVertAlign = getCellVertAlign(cell);
					if (cellVertAlign==null || cellVertAlign.length()==0) {
						vertAlignList.add(null);
					} else {
						vertAlignList.add(cellVertAlign.toString());
					}
				} else {
					try {
						// 设置水平对齐方式
						setCellTextAlign(cell, textAlignList.get(j));
						// 设置垂直对齐方式
						setCellVertAlign(cell, vertAlignList.get(j));
					} catch (Exception e) {
						
					}
				}
			}
		}
	}
	
	/**
	 * 替换表格变量
	 * @param doc
	 * @param param
	 */
	private static void replaceTable(XWPFDocument doc, Map<String, Object> param) {
		Iterator<XWPFTable> iterator = doc.getTablesIterator();
		XWPFTable table = null;
		
		while (iterator.hasNext()) {
			table = iterator.next();
			if (table.getRows().size()>0) {
				if (matcher(table.getText()).find()) {
					boolean flag = true;
					String key = "";
					int index = 1;
					
					jump:
					for (int i=0; i<table.getRows().size(); i++) {
						XWPFTableRow row = table.getRows().get(i);
						for (XWPFTableCell cell : row.getTableCells()) {
							for (XWPFParagraph paragraph : cell.getParagraphs()) {
								try {
									replaceParagraph(paragraph, param);
								} catch (Exception e) {
									// 此处表示是自动插入循环表格数据
									flag = false;
									key = e.getMessage();
									
									// 设置指定变量为指定值
									replaceParagraph(paragraph, "");
									
									// 第几行开始是数据（从0开始计）
									index = i;
									
									break jump;
								}
							}
						}
					}
					
					if (!flag) {
						// 为表格插入数据
						if (param.get(key)!=null && (param.get(key) instanceof Table)) {
							Table tableSetting = (Table) param.get(key);
							
							// 1.0 插入数据
							List<String[]> dataList = tableSetting.getDataList();
							if (dataList!=null && dataList.isEmpty()==false) {
								insertTable(table, dataList, index);
							}
							
							// 2.0 单元格合并列
							List<int[]> mergeColList = tableSetting.getMergeColList();
							if (mergeColList!=null && mergeColList.isEmpty()==false) {
								for (int[] mergeColArr : mergeColList) {
									mergeCol(table, mergeColArr[0], mergeColArr[1], mergeColArr[2]);
								}
							}
							
							// 3.0 单元格合并行
							List<int[]> mergeRowList = tableSetting.getMergeRowList();
							if (mergeRowList!=null && mergeRowList.isEmpty()==false) {
								for (int[] mergeRowArr : mergeRowList) {
									mergeRow(table, mergeRowArr[0], mergeRowArr[1], mergeRowArr[2]);
								}
							}
						}
					}
				}
			}
		}
	}

	/**
	 * 替换段落变量
	 * @param doc
	 * @param param
	 * @throws Exception
	 */
	private static void replaceParagraph(XWPFDocument doc, Map<String, Object> param) throws Exception {
		List<XWPFParagraph> paragraphList = doc.getParagraphs();
		if (paragraphList!=null && paragraphList.isEmpty()==false) {
			for (XWPFParagraph paragraph : paragraphList) {
				replaceParagraph(paragraph, param);
			}
		}
	}
	
	/**
	 * 设置指定变量为指定值
	 * @param paragraph
	 * @param value
	 */
	private static void replaceParagraph(XWPFParagraph paragraph, String value) {
		List<XWPFRun> runList = paragraph.getRuns();
		for (XWPFRun run : runList) {
			run.setText(value, 0);
		}
	}
	
	/**
	 * 替换段落变量
	 * @param paragraph
	 * @param param
	 * @throws Exception
	 */
	private static void replaceParagraph(XWPFParagraph paragraph, Map<String, Object> param) throws Exception {
		String tempString = "";
		Set<XWPFRun> runSet = new HashSet<XWPFRun>();
		char lastChar = ' ';
		List<XWPFRun> runList = paragraph.getRuns();
		for (XWPFRun run : runList) {
			String text = run.getText(0);
			if (text==null) {
				continue;
			}
			
			run.setText("", 0);
			run.setText(text, 0);
			for (int i=0; i<text.length(); i++) {
				char ch = text.charAt(i);
				if (ch=='$') {
					runSet = new HashSet<XWPFRun>();
					runSet.add(run);
					tempString = text;
				}
				else if (ch=='{') {
					if (lastChar=='$') {
						if (runSet.contains(run)) {
							
						} else {
							runSet.add(run);
							tempString = tempString + text;
						}
					} else {
						runSet = new HashSet<XWPFRun>();
						tempString = "";
					}
				}
				else if (ch=='}') {
					if (tempString!=null && tempString.indexOf("${")>=0) {
						if (runSet.contains(run)) {
							
						} else {
							runSet.add(run);
							tempString = tempString + text;
						}
					} else {
						runSet = new HashSet<XWPFRun>();
						tempString = ""; 
					}
					if (runSet.size()>0) {
						String replaceContent = replaceContent(tempString, param, run);
						if (!replaceContent.equals(tempString)) {
							int index = 0;
							XWPFRun aRun = null;
							for (XWPFRun tempRun : runSet) {
								tempRun.setText("", 0);
								if (index==0) {
									aRun = tempRun;
								}
								index++;
							}
							aRun.setText(replaceContent, 0);
						}
						runSet = new HashSet<XWPFRun>();
						tempString = ""; 
					}
				}
				else {
					if (runSet.size()<=0) {
						continue;
					}
					if (runSet.contains(run)) {
						continue;
					}
					runSet.add(run);
					tempString = tempString + text;
				}
				
				lastChar = ch;
			}
		}
	}
	
	/**
	 * 设置字体样式
	 * @param run
	 * @param fontStyle
	 */
	private static void setFontStyle(XWPFRun run, Font fontStyle) {
		if (fontStyle.getColor()!=null) {
			run.setColor(fontStyle.getColor());
		}
		if (fontStyle.getFontFamily()!=null) {
			run.setFontFamily(fontStyle.getFontFamily());
		}
		if (fontStyle.getFontSize()!=null && fontStyle.getFontSize()>0) {
			run.setFontSize(fontStyle.getFontSize());
		}
		if (fontStyle.getBold()) {
			run.setBold(true);
		}
		if (fontStyle.getItalic()) {
			run.setItalic(true);
		}
		if (fontStyle.getStrike()) {
			run.setStrikeThrough(true);
		}
	}
	
	/**
	 * 设置文本
	 * @param run
	 * @param text
	 */
	private static void setText(XWPFRun run, String text) {
		if (text.indexOf("<br/>")>=0) {
			setWrapText(run, text);
		} else {
			run.setText(text);
		}
	}
	
	/**
	 * 设置换行文本
	 * @param run
	 * @param text
	 */
	private static void setWrapText(XWPFRun run, String text) {
		String[] arr = text.split("<br/>");
		for (int n=0; n<arr.length; n++) {
			if (n==0) {
				run.setText(arr[n]);
			} else {
				run.addBreak();
				run.setText(arr[n]);
			}
		}
	}
	
	/**
	 * 替换内容
	 * @param text
	 * @param param
	 * @param run
	 * @return
	 * @throws Exception 
	 */
	private static String replaceContent(String text, Map<String, Object> param, XWPFRun run) throws Exception {
		if (text!=null) {
			for (Map.Entry<String, Object> entry : param.entrySet()) {
				String key = entry.getKey();
				if (text.indexOf(key)!=-1) {
					Object value = entry.getValue();
					if (value==null) {
						value = "";
					}
					
					// 文本替换
					if (value instanceof String) {
						String str = value.toString().replace("<br>", "<br/>");
						if (str.indexOf("<br/>")>=0) {
							text = text.replace(key, "");
							setWrapText(run, str);
						} else {
							text = text.replace(key, str);
						}
					}
					// 文本替换（带字体样式）
					else if (value instanceof Font) {
						Font fontStyle = (Font) value;
						
						// 设置字体样式
						setFontStyle(run, fontStyle);
						// 设置文本
						String str = value.toString().replace("<br>", "<br/>");
						if (str.indexOf("<br/>")>=0) {
							text = text.replace(key, "");
							setWrapText(run, fontStyle.getText());
						} else {
							text = text.replace(key, fontStyle.getText());
						}
					}
					// 图片替换
					else if (value instanceof Picture) {
						FileInputStream fis = null;
						try {
							Picture wordPicture = (Picture) value;
							
							double width = wordPicture.getWidth();
							double height = wordPicture.getHeight();
							String imgUrl = wordPicture.getUrl();
							String imgType = imgUrl.substring(imgUrl.lastIndexOf(".")+1);
							int imageType = getImageType(imgType);
							
							fis = new FileInputStream(imgUrl);
							if (fis!=null) {
								run.addPicture(fis, imageType, null, Units.toEMU(width), Units.toEMU(height));
							}
							
							// 图片描述
							if (wordPicture.getDescription()!=null) {
								run.addBreak();
								run.setText(wordPicture.getDescription());
							}
							
							text = text.replace(key, "");
						} catch (Exception e) {
							e.printStackTrace();
						} finally {
							IOUtils.closeQuietly(fis);
						}
					}
					// 插入表格数据
					else if (value instanceof Table) {
						throw new Exception(key);
					}
				}
			}
		}
		
		return text;
	}
	
	/**
	 * 表格合并列
	 * @param table
	 * @param rowIndex 行的索引（从0计）
	 * @param startCol 起始列（从0计）
	 * @param endCol   终止列（从0计）
	 */
	private static void mergeCol(XWPFTable table, int rowIndex, int startCol, int endCol) {
		for (int i=startCol; i<=endCol; i++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(i);
			if (i==startCol) {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * 表格合并行
	 * @param table
	 * @param colIndex 列的索引（从0计）
	 * @param startRow 起始行（从0计）
	 * @param endRow   终止行（从0计）
	 */
	private static void mergeRow(XWPFTable table, int colIndex, int startRow, int endRow) {
		for (int i=startRow; i<=endRow; i++) {
			XWPFTableCell cell = table.getRow(i).getCell(colIndex);
			if (i==startRow) {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
			}
		}
	}
	
	/**
	 * 替换Word占位符的内容，返回一个新的XWPFDocument对象
	 * @param filePath Word模板文件路径
	 * @param param 允许为空
	 * @return
	 */
	public static XWPFDocument getDocx(String filePath, Map<String, Object> param) {
		XWPFDocument docx = null;
		
		try {
			OPCPackage pack = POIXMLDocument.openPackage(filePath);
			docx = new XWPFDocument(pack);
			
			if (param!=null && param.size()>0) {
				// 替换段落
				replaceParagraph(docx, param);
				// 替换表格
				replaceTable(docx, param);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return docx;
	}
	
	/**
	 * 替换Word占位符的内容，返回一个新的XWPFDocument对象
	 * @param docx Word文件对象
	 * @param param 允许为空
	 * @return
	 */
	public static XWPFDocument getDocx(XWPFDocument docx, Map<String, Object> param) {
		try {
			if (param!=null && param.size()>0) {
				// 替换段落
				replaceParagraph(docx, param);
				// 替换表格
				replaceTable(docx, param);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return docx;
	}
	
	/**
	 * 将html写入word（.doc），指定word存储路径
	 * @param html 开头、结尾带<html></html>标签的字符串
	 * @param path word全路径，例如：D:\\Temp\\2.doc
	 * @return
	 */
	public static boolean htmlToDoc(String html, String path) {
		boolean flag = true;
		ByteArrayInputStream bais = null;
		FileOutputStream out = null;
		POIFSFileSystem fs = null;
		
		try {
			byte bArr[] = html.getBytes("GB2312");
			bais = new ByteArrayInputStream(bArr);
			fs = new POIFSFileSystem();
			DirectoryEntry directory = fs.getRoot();
			directory.createDocument("WordDocument", bais);
			out = new FileOutputStream(path);
			fs.writeFilesystem(out);
		} catch (Exception e) {
			e.printStackTrace();
			flag = false;
		} finally {
			IOUtils.closeQuietly(fs);
			IOUtils.closeQuietly(out);
			IOUtils.closeQuietly(bais);
		}
		
		return flag;
	}
	
	/**
	 * 将word（.docx后缀）转为html
	 * @param wordPath word文件的全路径，例如：D:\\Temp\\1.docx
	 * @return 返回生成的html文件的全路径，例如：D:\Temp\1_html\1.html
	 * @throws Exception
	 */
	public static String docxToHtml(String wordPath) throws Exception {
		String htmlPath = "";
		
		InputStream in = null;
		OutputStream out = null;
		XWPFDocument doc = null;
		
		try {
			File wordFile = new File(wordPath);
			if (!wordFile.exists()) {
				throw new FileNotFoundException("指定文件不存在：" + wordPath);
			}
			
			String wordName = wordFile.getName();
			String htmlName = wordName.replace(".docx", ".html");
			String wordFolderPath = wordFile.getParent();
			String htmlFolderPath = wordFolderPath + File.separator + wordName.replace(".docx", "") + "_html";
			
			// 1.0 判断html文件是否已存在
			File htmlFile = new File(htmlFolderPath + File.separator + htmlName);
			if (htmlFile.exists()) {
				return htmlFile.getAbsolutePath();
			} else {
				// 生成html文件上级文件夹
				File folder = new File(htmlFolderPath);
				if (!folder.exists()) {
					folder.mkdirs();
				}
			}
			
			// 2.0 生成html文件
			// 2.1 读取word
			doc = new XWPFDocument(new FileInputStream(wordFile));
			// 2.2 解析 XHTML配置
			ImageManager imageManager = new ImageManager(new File(htmlFolderPath), "image");    // html中图片的路径 相对路径
			
			XHTMLOptions options = XHTMLOptions.create();
			options.setImageManager(imageManager);
			options.setIgnoreStylesIfUnused(false);
			options.setFragment(true);
			
			// 2.3 将 XWPFDocument转换成XHTML
			out = new FileOutputStream(htmlFile);
			XHTMLConverter.getInstance().convert(doc, out, options);
			
			htmlPath = htmlFile.getAbsolutePath();
			
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException("指定文件不存在：" + wordPath);
		} catch (IOException e) {
			throw new IOException(e);
		} catch (Exception e) {
			throw new Exception(e);
		} finally {
			IOUtils.closeQuietly(doc);
			IOUtils.closeQuietly(out);
			IOUtils.closeQuietly(in);
		}
		
		return htmlPath;
	}

	/**
	 * 将word（.doc后缀）转为html
	 * @param wordPath word文件的全路径，例如：D:\\Temp\\1.docx
	 * @return 返回生成的html文件的全路径，例如：D:\Temp\1_html\1.html
	 * @throws Exception
	 */
	public static String docToHtml(String wordPath) throws Exception {
		String htmlPath = "";
		
		InputStream in = null;
		OutputStream out = null;
		HWPFDocument doc = null;
		
		try {
			File wordFile = new File(wordPath);
			if (!wordFile.exists()) {
				throw new FileNotFoundException("指定文件不存在：" + wordPath);
			}
			
			String wordName = wordFile.getName();
			String htmlName = wordName.replace(".doc", ".html");
			String wordFolderPath = wordFile.getParent();
			String htmlFolderPath = wordFolderPath + File.separator + wordName.replace(".doc", "") + "_html";
			
			// 判断html文件是否已存在
			File htmlFile = new File(htmlFolderPath + File.separator + htmlName);
			if (htmlFile.exists()) {
				return htmlFile.getAbsolutePath();
			} else {
				// 生成html文件上级文件夹
				File folder = new File(htmlFolderPath);
				if (!folder.exists()) {
					folder.mkdirs();
				}
			}
			
			// 图片目录
			final String imagePath = htmlFolderPath + File.separator + "image";
			
			// 原word文档
			in = new FileInputStream(wordFile);
			doc = new HWPFDocument(in);
			WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
			// 设置图片存放的位置
			wordToHtmlConverter.setPicturesManager(new PicturesManager() {
				public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
					File imgPath = new File(imagePath);
					if (!imgPath.exists()) {
						imgPath.mkdirs();
					}
					File file = new File(imagePath + File.separator + suggestedName);
					try {
						OutputStream os = new FileOutputStream(file);
						os.write(content);
						os.close();
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
					// 图片在html文件上的相对路径
					return "image" + File.separator + suggestedName;
				}
			});
			
			// 解析word文档
			wordToHtmlConverter.processDocument(doc);
			Document htmlDocument = wordToHtmlConverter.getDocument();
			out = new FileOutputStream(htmlFile);
			DOMSource domSource = new DOMSource(htmlDocument);
			StreamResult streamResult = new StreamResult(out);
			TransformerFactory factory = TransformerFactory.newInstance();
			Transformer serializer = factory.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");    // 是否添加空格
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.transform(domSource, streamResult);
			
			htmlPath = htmlFile.getAbsolutePath();
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException("指定文件不存在：" + wordPath);
		} catch (IOException e) {
			throw new IOException(e);
		} catch (Exception e) {
			throw new Exception(e);
		} finally {
			IOUtils.closeQuietly(doc);
			IOUtils.closeQuietly(out);
			IOUtils.closeQuietly(in);
		}
		
		return htmlPath;
	}
	
}
