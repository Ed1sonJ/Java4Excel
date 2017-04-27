package org.dxj.excel.util;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author : duxiji
 * @date : 2017/4/24
 * @description :
 */
public class ExcelUtils {
    /**
     * 从excel.xlsx里面读取数据，保存到 List<List<String>> 里面
     *
     * @param path
     * @return
     * @throws IOException
     */
    public static List<List<List<String>>> readxlsx(String path , int firstSheet ,int lastSheet) {
        List<List<List<String>>> result = null;
        try {
            //获取excel文件的io流
            InputStream is = new FileInputStream(path);
            //创建一个内存中的excel文件XSSFWorkbook类型对象，这个对象代表了整个excel文件
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            result = new ArrayList<List<List<String>>>();

            //循环每一页sheet，并处理当前页
            for(int i = firstSheet ; i <=lastSheet ; i ++){
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(i);
                if (xssfSheet == null) {
                    continue;
                }
                //处理当前页，循环处理每一行
                List<List<String>> sheetList = new ArrayList<List<String>>();
                //从1开始，不读取表头
                for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                    //获取当前行
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    int minColIndex = xssfRow.getFirstCellNum();
                    int maxColIndex = xssfRow.getLastCellNum();
                    List<String> rowList = new ArrayList<String>();
                    //处理当前行，循环处理每一个cell
                    for (int colIndex = minColIndex; colIndex < maxColIndex; colIndex++) {
                        //获取当前cell
                        XSSFCell cell = xssfRow.getCell(colIndex);
                        if (cell == null) {
                            continue;
                        }
                        rowList.add(cell.toString());
                    }
                    sheetList.add(rowList);
                }
                result.add(sheetList);
            }
            is.close();
        } catch (Exception e) {
        } finally {
        }
        return result;
    }

}
