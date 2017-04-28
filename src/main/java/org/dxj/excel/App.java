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
import java.util.ArrayList;
import java.util.List;

/**
 * Excel操作
 */
public class App {
    public static void main(String[] args) {
        String path = "E:\\test.xlsx";
        //获取了所有的数据，放到result里面
        List<List<List<String>>> result = ExcelUtils.readxlsx(path, 0, 2);

        //保存9科的排名，方便后面计算对出率，核心数据结构
        List<Rank> rankList = BizUtils.generateListRank(result.get(2).size(), result.get(2));

        //开始生成表
        //1.生成家长表
        new GenerateParentBiz().generateSheet(path, "成绩（家长版）", result, rankList);

        //2.生成进退表
        GenerateAdvanceRetreatBiz jintuiSheet = new GenerateAdvanceRetreatBiz();
        jintuiSheet.generateSheet(path, "每科进退-期末", result.get(0), result.get(2));//表4（下标）
        jintuiSheet.generateSheet(path, "每科进退-一段", result.get(1), result.get(2));//表5（下标）

        //3.生成科任-X表，必须运行在生成2个进退表之后，不然会报错
        List<List<List<String>>> advanceRetreat = ExcelUtils.readxlsx(path, 4, 5);
        //index指的是在进退表中科任-X的X所在的列的index
        GenerateBranchBiz kerenSheet = new GenerateBranchBiz();
        List<String> kemuList = new ArrayList<String>();
        kemuList.add("科任-语文");
        kemuList.add("科任-数学");
        kemuList.add("科任-英文");
        kemuList.add("科任-物理");
        kemuList.add("科任-化学");
        kemuList.add("科任-生物");
        kemuList.add("科任-政治");
        kemuList.add("科任-历史");
        kemuList.add("科任-地理");
        int i = 3;
        for (String kemu : kemuList) {
            kerenSheet.generateSheet(path, kemu, i, result.get(2),
                    advanceRetreat.get(0), advanceRetreat.get(1));
            i++;
        }

        //4.生成XX-前X表


    }


}
