package org.dxj.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.dxj.excel.biz.GenerateBiz;
import org.dxj.excel.util.ExcelUtils;

import java.io.IOException;
import java.util.List;

/**
 * Hello world!
 */
public class App {
    public static List<List<String>> sheet1;//放第一个表的数据
    public static List<List<String>> sheet2;//放第二个表的数据
    public static List<List<String>> sheet3;//放第三个表的数据

    public static void main(String[] args) {
            //获取了所有的数据，放到result里面
            List<List<List<String>>> result = ExcelUtils.readxlsx("E:\\test.xlsx");
            sheet1 = result.get(0);
            for (List<String> rowList : sheet1) {
                for (String cell : rowList) {
                    System.out.print(cell + " ");
                }
                System.out.println();
            }
            sheet2 = result.get(1);
            for (List<String> rowList : sheet2) {
                for (String cell : rowList) {
                    System.out.print(cell + " ");
                }
                System.out.println();
            }
            sheet3 = result.get(2);
            for (List<String> rowList : sheet3) {
                for (String cell : rowList) {
                    System.out.print(cell + " ");
                }
                System.out.println();
            }
        GenerateBiz.generateSheet("E:\\test.xlsx","家长版",result);
    }
}
