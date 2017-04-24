package org.dxj.excel.util;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @author : duxiji
 * @date : 2017/4/24
 * @description :
 */
public class ExcelUtils {
    public static List<List<String>> readxlsx(String path) throws IOException {
        //获取excel文件的io流
        InputStream is = new FileInputStream(path);
        //创建一个内存中的excel文件XSSFWorkbook类型对象，这个对象代表了整个excel文件
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
    }

}
