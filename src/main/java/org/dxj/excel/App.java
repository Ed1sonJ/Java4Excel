package org.dxj.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.regex.REUtil;
import org.dxj.excel.bean.Rank;
import org.dxj.excel.biz.BizUtils;
import org.dxj.excel.biz.GenerateAdvanceRetreatBiz;
import org.dxj.excel.biz.GenerateParentBiz;
import org.dxj.excel.util.ExcelUtils;

import java.io.*;
import java.util.List;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) {
        String path = "E:\\test.xlsx";
        //获取了所有的数据，放到result里面
        List<List<List<String>>> result = ExcelUtils.readxlsx(path);

        //保存9科的排名，方便后面计算对出率，核心数据结构
        List<Rank> rankList = BizUtils.generateListRank(result.get(2).size(), result.get(2));

//        new GenerateParentBiz().generateSheet(path, "成绩（家长版）", result, rankList);
//        new GenerateAdvanceRetreatBiz().generateSheet(path, "每科进退-期末", result.get(0),result.get(2));
        new GenerateAdvanceRetreatBiz().generateSheet(path, "每科进退-一段", result.get(1),result.get(2));


    }

//    private static void writeWorkbook(String path, XSSFWorkbook xssfWorkbook) {
//        OutputStream os = null;
//        try {
//            os = new FileOutputStream(path);
//            xssfWorkbook.write(os);
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//
//    private static XSSFWorkbook openWorkbook(String path) {
//        InputStream is = null;
//        try {
//            is = new FileInputStream(path);
//            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
//            return xssfWorkbook;
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        return null;
//    }
}
