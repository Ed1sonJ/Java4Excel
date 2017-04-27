package org.dxj.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.regex.REUtil;
import org.dxj.excel.bean.Rank;
import org.dxj.excel.biz.BizUtils;
import org.dxj.excel.biz.GenerateAdvanceRetreatBiz;
import org.dxj.excel.biz.GenerateBranchBiz;
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
        List<List<List<String>>> result = ExcelUtils.readxlsx(path, 0, 2);

        //保存9科的排名，方便后面计算对出率，核心数据结构
        List<Rank> rankList = BizUtils.generateListRank(result.get(2).size(), result.get(2));

        new GenerateParentBiz().generateSheet(path, "成绩（家长版）", result, rankList);
        new GenerateAdvanceRetreatBiz().generateSheet(path, "每科进退-期末", result.get(0), result.get(2));//表4（下标）
        new GenerateAdvanceRetreatBiz().generateSheet(path, "每科进退-一段", result.get(1), result.get(2));//表5（下标）

        //必须运行在生成4-5表之后，不然会报错
        List<List<List<String>>> advanceRetreat = ExcelUtils.readxlsx(path, 4, 5);
        //index指的是在进退表中科任-X的X所在的列的index
        new GenerateBranchBiz().generateSheet(path, "科任-语文", 3,result.get(2),
                advanceRetreat.get(0), advanceRetreat.get(1));

    }


}
