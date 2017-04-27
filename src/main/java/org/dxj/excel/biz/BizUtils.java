package org.dxj.excel.biz;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.dxj.excel.bean.Rank;

import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author : duxiji
 * @date : 2017/4/27
 * @description : 包含了计算业务需要的函数
 */
public class BizUtils {

    /**
     * 计算对位率
     *
     * @param target
     * @param rank
     */
    public static String computeRate(String target, Rank rank) {
        //获得平均名次
        double averageRnk = Rank.get9rank(rank);
        double targetRank = Double.parseDouble(target);
        double result = (averageRnk - targetRank) / averageRnk;
        String resultPercent = demical2Percent(result);
        return resultPercent;
    }

    /**
     * 小数转百分数
     *
     * @param result
     * @return
     */
    public static String demical2Percent(double result) {
        NumberFormat nf = NumberFormat.getPercentInstance();
        nf.setMinimumFractionDigits(2);//保留两位小数
        return nf.format(result);
    }

    /**
     * 生成对应列数学号，9科排名，9科总分排名
     *
     * @param sheetLength
     * @param sheetData
     * @return
     */
    public static List<Rank> generateListRank(int sheetLength, List<List<String>> sheetData) {
        List<Rank> rankList = new ArrayList<Rank>();
        for (int i = 0; i < sheetLength; i++) {
            Rank rank = new Rank();
            int h = 1;
            rank.setId(sheetData.get(i).get(h));
            h = h + 3;
            rank.setChinese(sheetData.get(i).get(h));
            h = h + 2;
            rank.setMath(sheetData.get(i).get(h));
            h = h + 2;
            rank.setEnglish(sheetData.get(i).get(h));
            h = h + 2;
            rank.setPhysics(sheetData.get(i).get(h));
            h = h + 2;
            rank.setChemology(sheetData.get(i).get(h));
            h = h + 2;
            rank.setBiological(sheetData.get(i).get(h));
            h = h + 2;
            rank.setPolitics(sheetData.get(i).get(h));
            h = h + 2;
            rank.setHistory(sheetData.get(i).get(h));
            h = h + 2;
            rank.setGeography(sheetData.get(i).get(h));
            h = h + 4;
            rank.setTotalRank(sheetData.get(i).get(h));
            h = h + 2;
            rank.setWenkeRank(sheetData.get(i).get(h));
            h = h + 2;
            rank.setLikeRank(sheetData.get(i).get(h));
            rankList.add(rank);
        }
        return rankList;
    }

    /**
     * 生成表头head
     */
    public static void generateHeadData(XSSFRow row, List<String> data) {
        for (int i = 0; i < data.size(); i++) {
            row.createCell(i).setCellValue(data.get(i));
        }
    }

    /**
     * 初始化家长表的所有行row
     */

    public static XSSFSheet newBody(int sheetLength, XSSFSheet sheet) {
        //有41行数据，所以i<=sheetLength
        for (int i = 1; i <= sheetLength; i++) {
            sheet.createRow(i);
        }
        return sheet;
    }

}


