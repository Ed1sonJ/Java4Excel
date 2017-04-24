package org.dxj.excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.dxj.excel.util.ExcelUtils;

import java.io.IOException;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        try {
            List<List<List<String>>> result = ExcelUtils.readxlsx("E:\\test.xlsx");
            //获取第一个sheet的数据
            List<List<String>> sheet = result.get(2);
            for(List<String> rowList : sheet){
                for (String cell : rowList){
                    System.out.print(cell+" ");
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
