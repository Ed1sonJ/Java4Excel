package org.dxj.excel;

import org.dxj.excel.bean.Rank;
import org.dxj.excel.biz.*;
import org.dxj.excel.util.ExcelUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * Excel操作
 */
public class App {
    public static void main(String[] args) {

        System.out.println("开始...");
        String path = "E:\\test.xlsx";
        //获取了所有的数据，放到result里面
        List<List<List<String>>> result = ExcelUtils.readxlsx(path, 0, 2);

        //保存9科的排名，方便后面计算对出率，核心数据结构
        List<Rank> rankList = BizUtils.generateListRank(result.get(2).size(), result.get(2));

        //开始生成表
        //1.生成家长表
        try {
            new GenerateParentBiz().generateSheet(path, "成绩（家长版）", result, rankList);

            //2.生成进退表
            GenerateAdvanceRetreatBiz jintuiSheet = new GenerateAdvanceRetreatBiz();
            jintuiSheet.generateSheet(path, "每科进退-期末", result.get(0), result.get(2));//表4（下标）
            jintuiSheet.generateSheet(path, "每科进退-一段", result.get(1), result.get(2));//表5（下标）

            //3.生成科任-X表，必须运行在生成2个进退表之后，不然会报错
            List<List<List<String>>> advanceRetreat = ExcelUtils.readxlsx(path, 4, 5);
            //准备生成表的数据
            List<List<String>> jiazhangbiao = ExcelUtils.readxlsx(path, 3);
            List<String> numList = new ArrayList<String>();//存放学号
            List<String> totalRank = new ArrayList<String>();//存放9科排名
            List<String> threeRank = new ArrayList<String>();//存放3总排名
            List<String> wenkeRank = new ArrayList<String>();//存放文科排名
            List<String> likeRank = new ArrayList<String>();//存放理科排名
            List<String> chineseRank = new ArrayList<String>();//存放语文排名
            List<String> mathRank = new ArrayList<String>();//存放数学排名
            List<String> englishRank = new ArrayList<String>();//存放英语排名
            List<String> physicsRank = new ArrayList<String>();//存放物理排名
            List<String> chemologyRank = new ArrayList<String>();//存放化学排名
            List<String> biologyRank = new ArrayList<String>();//存放生物排名
            List<String> politicsRank = new ArrayList<String>();//存放政治排名
            List<String> historyRank = new ArrayList<String>();//存放历史排名
            List<String> geographyRank = new ArrayList<String>();//存放地理排名
            List<String> classRank = new ArrayList<String>();//存放班级排名
            List<List<String>> targetRank = new ArrayList<List<String>>();
//        targetRank.add(totalRank);
            targetRank.add(threeRank);
            targetRank.add(wenkeRank);
            targetRank.add(likeRank);
            targetRank.add(chineseRank);
            targetRank.add(mathRank);
            targetRank.add(englishRank);
            targetRank.add(physicsRank);
            targetRank.add(chemologyRank);
            targetRank.add(biologyRank);
            targetRank.add(politicsRank);
            targetRank.add(historyRank);
            targetRank.add(geographyRank);
            for (List<String> list : jiazhangbiao) {
                numList.add(list.get(1));
                totalRank.add(list.get(15));
                threeRank.add(list.get(13));
                wenkeRank.add(list.get(17));
                likeRank.add(list.get(19));
                chineseRank.add(list.get(3));
                mathRank.add(list.get(4));
                englishRank.add(list.get(5));
                physicsRank.add(list.get(6));
                chemologyRank.add(list.get(7));
                biologyRank.add(list.get(8));
                politicsRank.add(list.get(9));
                historyRank.add(list.get(10));
                geographyRank.add(list.get(11));
                classRank.add(list.get(33));
            }
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
                //jiazhangbiao.get(14)为班级排名
                kerenSheet.generateSheet(path, kemu, i, result.get(2),
                        advanceRetreat.get(0), advanceRetreat.get(1), classRank);
                i++;
            }

            //4.生成XX-前X表
            List<String> paimingName = new ArrayList<String>();
            paimingName.add("24班总分前十");
            paimingName.add("24班3总前五");
            paimingName.add("24班文科前五");
            paimingName.add("24班理科前五");
            paimingName.add("24班语文前五");
            paimingName.add("24班数学前五");
            paimingName.add("24班英语前五");
            paimingName.add("24班物理前五");
            paimingName.add("24班化学前五");
            paimingName.add("24班生物前五");
            paimingName.add("24班政治前五");
            paimingName.add("24班历史前五");
            paimingName.add("24班地理前五");
            //开始生成表
            GeneratePaimingBiz paimingSheet = new GeneratePaimingBiz();
            //防止脏数据
            List<String> numListTemp = new ArrayList<String>(Arrays.asList(new String[numList
                    .size()]));
            Collections.copy(numListTemp, numList);
            List<List<String>> paimingResult;
            paimingResult = getResultFromNum(numListTemp, totalRank, jiazhangbiao, 10);
            paimingSheet.generateSheet(path, paimingName.get(0), paimingResult);
            //生成后面的表
            i = 1;
            for (List<String> rank : targetRank) {
                Collections.copy(numListTemp, numList);
                paimingResult = getResultFromNum(numListTemp, rank, jiazhangbiao, 5);
                paimingSheet.generateSheet(path, paimingName.get(i), paimingResult);
                i++;
            }
            System.out.println("成功！");
        } catch (Exception e) {
            System.out.println("生成失败！");
            e.printStackTrace();
        }
    }

    /**
     * 用排序好的学号筛选出result
     *
     * @param numList
     * @param targetList
     */
    private static List<List<String>> getResultFromNum(List<String> numList, List<String> targetList, List<List<String>> jiazhangbiao, int rankLength) {
        BizUtils.BubbleSort(numList, targetList);//冒泡排序
        List<List<String>> paimingResult = new ArrayList<List<String>>();
        //取出前 rankLength 名
        for (int i = 0; i < rankLength; i++) {
            String num = numList.get(i);
            List<String> row = getListFromNum(num, jiazhangbiao);
            paimingResult.add(row);
        }
        return paimingResult;
    }

    /**
     * 为XX-前X写的方法，根据学号获得对应的row数据
     *
     * @param num
     * @param jiazhangbiao
     * @return
     */
    public static List<String> getListFromNum(String num, List<List<String>> jiazhangbiao) {
        List<String> result = new ArrayList<String>();
        //拿到每一行
        for (List<String> row : jiazhangbiao) {
            //假如num等于行的学号
            if (num.equals(row.get(1))) {
                //遍历获取需要的信息
                for (int i = 0; i < row.size(); i++) {
                    result.add(row.get(i));
                }
                break;
            }
        }
        return result;
    }
}
