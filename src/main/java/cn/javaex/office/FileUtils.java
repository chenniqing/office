package cn.javaex.office;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

/**
 * 文件工具类
 * 
 * @author 陈霓清
 */
public class FileUtils {

	/**
	 * 获取项目所在磁盘的文件夹路径，并设置临时目录
	 * @return
	 */
	private static String getFolderPath() {
		HttpServletRequest request = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest();
		// 获取地址内容，原路径（项目名）
		String realPath = request.getSession().getServletContext().getRealPath("/");
		// 项目名
		String path = request.getContextPath();
		path = path.replace("/", "") + File.separator;
		String projectPath = realPath.replace(path, "");
		
		String folderPath = projectPath + File.separator + "temp_download";
		File file = new File(folderPath);
		file.mkdirs();
		
		return folderPath;
	}
	
	/**
	 * 写Word
	 * @param docx
	 * @param filePath 文件写到哪里的全路径
	 * @throws IOException
	 */
	public static void writeDocx(XWPFDocument docx, String filePath) throws IOException {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		if(!targetFile.getParentFile().exists()){
			targetFile.getParentFile().mkdirs();
		}
		
		FileOutputStream out = new FileOutputStream(targetFile);
		docx.write(out);
		out.flush();
		IOUtils.closeQuietly(out);
	}
	
	/**
	 * 下载Word
	 * @param docx
	 * @param fileName
	 * @throws IOException
	 */
	public static void downloadDocx(XWPFDocument docx, String fileName) throws IOException {
		String folderPath = getFolderPath();
		
		FileOutputStream out = new FileOutputStream(folderPath + File.separator + fileName);
		docx.write(out);
		out.flush();
		IOUtils.closeQuietly(out);
		
		downloadFile(folderPath, fileName);
	}
	
	/**
	 * 写Excel
	 * @param xlsx
	 * @param filePath 文件写到哪里的全路径
	 * @throws IOException
	 */
	public static void writeXlsx(XSSFWorkbook xlsx, String filePath) throws IOException {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		if(!targetFile.getParentFile().exists()){
			targetFile.getParentFile().mkdirs();
		}
		
		FileOutputStream out = new FileOutputStream(targetFile);
		xlsx.write(out);
		out.flush();
		IOUtils.closeQuietly(out);
	}
	
	/**
	 * 下载Excel
	 * @param xlsx
	 * @param fileName 文件名，例如：1.xlsx
	 * @throws IOException
	 */
	public static void downloadXlsx(XSSFWorkbook xlsx, String fileName) throws IOException {
		String folderPath = getFolderPath();
		
		FileOutputStream out = new FileOutputStream(folderPath + File.separator + fileName);
		xlsx.write(out);
		out.flush();
		IOUtils.closeQuietly(out);
		
		downloadFile(folderPath, fileName);
	}
	
	/**
	 * 文件下载
	 * @param filePath 文件的绝对路径（带具体的文件名）
	 * @throws IOException
	 */
	public static void downloadFile(String filePath) throws IOException {
		downloadFile(filePath, null);
	}
	
	/**
	 * 文件下载
	 * @param folderPath 文件所在文件夹路径
	 * @param fileName 文件名称（带后缀）
	 * @throws IOException
	 */
	public static void downloadFile(String folderPath, String fileName) throws IOException {
		HttpServletResponse response = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse();
		
		File file = null;
		if (fileName==null || fileName.length()==0) {
			file = new File(folderPath);
		} else {
			file = new File(folderPath, fileName);
		}
		
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		
		response.setContentType("application/octet-stream");
		response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(file.getName(), "UTF-8"));
		response.setHeader("Content-Length", String.valueOf(file.length()));
		
		bis = new BufferedInputStream(new FileInputStream(file));
		bos = new BufferedOutputStream(response.getOutputStream());
		byte[] buff = new byte[2048];
		while (true) {
			int bytesRead;
			
			if (-1 == (bytesRead=bis.read(buff, 0, buff.length))) {
				break;
			}
			
			bos.write(buff, 0, bytesRead);
		}
		
		bos.flush();
		IOUtils.closeQuietly(bis);
		IOUtils.closeQuietly(bos);
	}
	
	/**
	 * 删除文件或目录（目录本身也删除）
	 * @param path 目录或文件的全路径
	 * @return boolean
	 */
	public static boolean deleteFile(String path) {
		File file = new File(path);
		return deleteFile(file);
	}
	
	/**
	 * 删除文件或目录
	 * @param file File文件对象
	 * @return boolean
	 */
	public static boolean deleteFile(File file) {
		if (!file.exists()) {
			return false;
		}
		
		if (file.isDirectory()) {
			File[] fileArr = file.listFiles();
			for (File f : fileArr) {
				deleteFile(f);
			}
		}
		
		return file.delete();
	}

	/**
	 * 删除文件或目录
	 * @param path 目录或文件的全路径
	 * @param flag 是否删除目录本身
	 * @return boolean
	 */
	public static boolean deleteFile(String path, boolean flag) {
		File file = new File(path);
		return deleteFile(file, flag);
	}
	
	/**
	 * 删除文件或目录（是否删除目录本身）
	 * @param file File文件对象
	 * @return boolean
	 */
	private static boolean deleteFile(File file, boolean flag) {
		if (!file.exists()) {
			return false;
		}
		
		if (file.isDirectory()) {
			File[] fileArr = file.listFiles();
			for (File f : fileArr) {
				deleteFile(f);
			}
		}
		
		if (flag) {
			return file.delete();
		}
		
		return true;
	}
	
	/**
	 * 获取文件源中的所有文件
	 * @param sourcePath
	 * @return
	 * @throws FileNotFoundException 
	 */
	private static List<File> getSourceAllFiles(String sourcePath) throws FileNotFoundException {
		List<File> fileList = new ArrayList<File>();
		
		File sourceFile = new File(sourcePath);
		// 判断文件或目录是否存在
		if (!sourceFile.exists()) {
			throw new FileNotFoundException("待压缩的文件或目录不存在：" + sourcePath);
		}
		
		File[] fileArr = sourceFile.listFiles();
		for (File file : fileArr) {
			if (file.isFile()) {
				fileList.add(file);
			} else {
				// 递归，获取路径中子路径中的所有文件
				fileList.addAll(getSourceAllFiles(file.getPath()));
			}
		}
		
		return fileList;
	}
	
	/**
	 * 得到文件在文件夹中的相对路径，保持原有结构
	 * @param sourcePath
	 * @param file
	 * @return
	 */
	private static String getRealName(String sourcePath, File file) {
		return file.getAbsolutePath().replace(sourcePath + File.separator, "").replace(sourcePath, "");
	}
	
	/**
	 * 创建一个zip压缩文件，并存放到新的路径中
	 * @param sourcePath 源目录或文件的全路径，例如：D:\\Temp  或  D:\\Temp\\1.docx
	 * @param zipPath 压缩后的文件全路径，例如：D:\\Temp\\xx.zip
	 * @throws Exception 
	 */
	public static void zip(String sourcePath, String zipPath) throws Exception {
		zip(sourcePath, zipPath, true);
	}
	
	/**
	 * 创建一个zip压缩文件，并存放到新的路径中
	 * @param sourcePath 源目录或文件的全路径，例如：D:\\Temp  或  D:\\Temp\\1.docx
	 * @param zipPath 压缩后的文件全路径，例如：D:\\Temp\\xx.zip
	 * @param keepFolder 是否将目录名称也一起压缩
	 * @throws Exception 
	 */
	public static void zip(String sourcePath, String zipPath, boolean keepFolder) throws Exception {
		byte[] buffer = new byte[1024*10];
		FileInputStream fis = null;
		BufferedInputStream bis = null;
		FileOutputStream fos = null;
		ZipOutputStream zos = null;
		
		try {
			File sourceFile = new File(sourcePath);
			// 判断文件或目录是否存在
			if (!sourceFile.exists()) {
				throw new FileNotFoundException("待压缩的文件或目录不存在：" + sourcePath);
			}
			
			File zipFile = new File(zipPath);
			if (zipFile.exists()) {
				deleteFile(zipFile);    // 如果压缩包已存在，则先删除
			}
			
			// 判断是否是一个具体的文件
			if (sourceFile.isFile()) {
				fos = new FileOutputStream(zipFile);
				zos = new ZipOutputStream(new BufferedOutputStream(fos));
				
				// 创建zip实体，并添加进压缩包  
				ZipEntry zipEntry = new ZipEntry(sourceFile.getName());
				zos.putNextEntry(zipEntry);
				
				fis = new FileInputStream(sourceFile);
				bis = new BufferedInputStream(fis, 1024*10);
				int read = 0;
				while ((read=bis.read(buffer, 0, 1024*10)) != -1) {
					zos.write(buffer, 0, read);
				}
				zos.flush();
			}
			// 目录
			else if (sourceFile.isDirectory()) {
				// 获取文件源中的所有文件
				List<File> fileList = getSourceAllFiles(sourcePath);
				
				fos = new FileOutputStream(zipFile);
				zos = new ZipOutputStream(new BufferedOutputStream(fos));
				
				// 将每个文件放入zip流中
				for (File file : fileList) {
					String name = getRealName(sourcePath, file);    // 获取文件相对路径，保持文件原有结构
					ZipEntry zipEntry = null;
					if (keepFolder) {
						zipEntry = new ZipEntry(new File(sourcePath).getName() + File.separator + name); 
					} else {
						zipEntry = new ZipEntry(name); 
					}
					
					zipEntry.setSize(file.length());
					zos.putNextEntry(zipEntry);
					
					fis = new FileInputStream(file);
					bis = new BufferedInputStream(fis, 1024*10);
					int read = 0;
					while ((read=bis.read(buffer, 0, 1024*10)) != -1) {
						zos.write(buffer, 0, read);
					}
					
					zos.flush();
					bis.close();
					fis.close();
				}
				zos.close();
			}
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException("待压缩的文件或目录不存在：" + sourcePath);
		} catch (IOException e) {
			throw new IOException(e);
		} catch (Exception e) {
			throw new Exception(e);
		} finally {
			IOUtils.closeQuietly(zos);
			IOUtils.closeQuietly(fos);
			IOUtils.closeQuietly(bis);
			IOUtils.closeQuietly(fis);
		}
	}
	
	/**
	 * zip解压
	 * @param zipPath zip文件的全路径，例如：D:\\Temp\\xx.zip
	 * @param destDirPath 解压后的目标文件夹路径，例如：D:\\Tempxx
	 * @throws Exception
	 */
	public static void unZip(String zipPath, String destDirPath) throws Exception {
		byte[] buffer = new byte[1024*10];
		ZipFile zipFile = null;
		
		try {
			File file = new File(zipPath);
			if (!file.exists()) {
				throw new FileNotFoundException("待解压的文件不存在：" + zipPath);
			}
			
			zipFile = new ZipFile(zipPath);
			Enumeration<?> entries = zipFile.entries();
			while (entries.hasMoreElements()) {
				ZipEntry entry = (ZipEntry) entries.nextElement();
				// 如果是文件夹，就创建个文件夹
				if (entry.isDirectory()) {
					String dirPath = destDirPath + File.separator + entry.getName();
					File dir = new File(dirPath);
					dir.mkdirs();
				} else {
					// 如果是文件，就先创建一个文件，然后用io流把内容copy过去
					File targetFile = new File(destDirPath + File.separator + entry.getName());
					// 保证这个文件的父文件夹必须要存在
					if(!targetFile.getParentFile().exists()){
						targetFile.getParentFile().mkdirs();
					}
					targetFile.createNewFile();
					// 将压缩文件内容写入到这个文件中
					InputStream in = zipFile.getInputStream(entry);
					BufferedInputStream bis = new BufferedInputStream(in, 1024*10);
					FileOutputStream fos = new FileOutputStream(targetFile);
					
					int read;
					while ((read=bis.read(buffer, 0, 1024*10)) != -1) {
						fos.write(buffer, 0, read);
					}
					
					fos.close();
					in.close();
				}
			}
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException("待解压的文件不存在：" + zipPath);
		} catch (IOException e) {
			throw new IOException(e);
		} catch (Exception e) {
			throw new Exception(e);
		} finally {
			IOUtils.closeQuietly(zipFile);
		}
	}
	
}
