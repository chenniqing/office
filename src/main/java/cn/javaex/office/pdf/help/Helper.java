package cn.javaex.office.pdf.help;

import java.io.File;
import java.net.URL;

import cn.javaex.office.common.util.PathUtils;

/**
 * 顶级帮助类
 * 
 * @author 陈霓清
 */
public class Helper {

	/**
	 * 得到真实的字体路径
	 * @param fontFamily
	 * @return
	 */
	public String getRealPath(String fontFamily) {
		String path = fontFamily;
		
		// resources文件夹下的字体
		if (path.startsWith("resources:")) {
			path = path.replace("resources:", "");
			if (path.startsWith("/")) {
				path = path.substring(1);
			}
			
			URL fontPath = Thread.currentThread().getContextClassLoader().getResource(path);
			path = fontPath + "";
			if (path.endsWith(".ttc")) {
				path = path + ",1";
			}
		} else {
			boolean absolutePath = PathUtils.isAbsolutePath(path);
			
			if (!absolutePath) {
				String projectPath = PathUtils.getProjectPath();
				path = projectPath + File.separator + path;
			}
		}
		
		if (path.endsWith(".ttc")) {
			path = path + ",1";
		}
		
		return path;
	}

}
