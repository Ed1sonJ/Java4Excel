package org.dxj.excel.biz;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dxj.excel.bean.Rank;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * author : duxiji
 * date : 2017/4/25
 * description : 生成成绩表（家长版）
 */
public class GenerateParentBiz extends BaseBiz{

    public void generateSheet(String path ,String name, List<List<List<String>>> data,List<Rank> rankList) {
        try {
            //拿到本次成绩
            List<List<String>> sheet3 = data.get(2);
            int sheetLength = sheet3.size();//41

            InputStream is = new FileInputStream(path);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            //生成家长表
            XSSFSheet parentSheet = xssfWorkbook.createSheet(name);
            //获取第0行
            XSSFRow row0 = parentSheet.createRow(0);
            //生成表头
            generateHeadData(row0);
            //初始化家长表的所有行
            parentSheet = BizUtils.newBody(sheetLength, parentSheet);
            //生成家长表的所有数据
            generateBodyData(parentSheet, data,rankList);

            OutputStream os = new FileOutputStream(path);
            xssfWorkbook.write(os);

            is.close();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();

        }
    }

    /**
     * 填充sheet的数据
     *
     */
    private void generateBodyData(XSSFSheet parentSheet, List<List<List<String>>> data,List<Rank> rankList) {
        List<List<String>> sheet1 = data.get(0);
        List<List<String>> sheet2 = data.get(1);
        List<List<String>> sheet3 = data.get(2);
        int sheetLength = sheet3.size();//41

        //先填充班级排名，涉及到全部学号和9科总分的排名，所以先进行
//        parentSheet = generateClassRank(rankList, parentSheet, sheetLength);

        for (int i = 1; i <= sheetLength; i++) {
            XSSFRow row = parentSheet.getRow(i);
            //生成每一行的数据
            generateRowData(row, sheet1.get(i - 1), sheet2.get(i - 1), sheet3.get(i - 1), rankList.get(i - 1));
        }
    }

    /**
     * 填充行数据
     */
    private void generateRowData(XSSFRow row, List<String> sheet1Row, List<String> sheet2Row,
                                        List<String> sheet3Row, Rank rank) {
        //填充班别、学号、姓名
        for (int i = 0; i < 3; i++) {
            row.createCell(i).setCellValue(sheet3Row.get(i));
        }

        //填充学科名次,引入了rankList之后可以修改
        int j = 2;//sheet3中的排位
        for (int i = 3; i < 12; i++) {
            j = j + 2;
            row.createCell(i).setCellValue(sheet3Row.get(j));
        }

        //填充3科名次，9科名次等
        //j=20
        for (int i = 12; i < 20; i++) {
            j++;
            row.createCell(i).setCellValue(sheet3Row.get(j));
        }

        //填充学科对位率
        List<String> rankString = generateRankList(rank);
        j = 0;
        for (int i = 20; i < 31; i++) {
            row.createCell(i).setCellValue(rankString.get(j));
            j++;
        }

        //填充期末和一段进退
        generateAdvanceAndRetreat(sheet1Row, sheet2Row, sheet3Row, row);
    }

    /**
     * 填充班级排名
     *
     */
    private XSSFSheet generateClassRank(List<Rank> rankList, XSSFSheet parentSheet, int length) {
        // TODO: 2017/4/27 算法错误
        //使用TreeMap可以进行排序
        //存放<9科排名，学号>，以9科总分排名
        Map<Float, String> idAndTotalRank = new TreeMap<Float, String>();
        for (int i = 0; i < length; i++) {
            Rank rank = rankList.get(i);
            idAndTotalRank.put(Float.parseFloat(rank.getTotalRank()), rank.getId());
        }

        //存放<学号，班级排名>
        //以学号作为排序
        Map<String, String> idAndClassRank = new TreeMap<String, String>();
        Set<Float> keySet = idAndTotalRank.keySet();
        Iterator<Float> iter = keySet.iterator();
        int i = 1;
        while (iter.hasNext()) {
            Float key = iter.next();
            //获取到学号
            System.out.println(key + ":" + idAndTotalRank.get(key));
            String num = idAndTotalRank.get(key);
            idAndClassRank.put(num, String.valueOf(i));
            i++;
        }

        //遍历Map，将班级排名写进sheet
        Set<String> keySet2 = idAndClassRank.keySet();
        Iterator<String> iter2 = keySet2.iterator();
        int j = 1;
        while (iter2.hasNext()) {
            String key = iter2.next();
            System.out.println(key + ":" + idAndClassRank.get(key));
            parentSheet.getRow(j).createCell(33).setCellValue(idAndClassRank.get(key));
            j++;
        }
        return parentSheet;
    }

    /**
     * 填充期末和一段进退
     *
     */
    private void generateAdvanceAndRetreat(List<String> sheet1Row, List<String> sheet2Row, List<String> sheet3Row, XSSFRow row) {
        String endPeriod = sheet1Row.get(24);
        String yiduan = sheet2Row.get(24);
        String orgin = sheet3Row.get(24);
        float rank1 = Float.parseFloat(endPeriod) - Float.parseFloat(orgin);
        float rank2 = Float.parseFloat(endPeriod) - Float.parseFloat(yiduan);

        row.createCell(31).setCellValue(String.valueOf(rank1));
        row.createCell(32).setCellValue(String.valueOf(rank2));

    }

    /**
     * 填充对位率
     *
     */
    private List<String> generateRankList(Rank rank) {

        String chineseRankRate = BizUtils.computeRate(rank.getChinese(), rank);
        String mathRankRate = BizUtils.computeRate(rank.getMath(), rank);
        String englishRankRate = BizUtils.computeRate(rank.getEnglish(), rank);
        String physicsRankRate = BizUtils.computeRate(rank.getPhysics(), rank);
        String chemologyRankRate = BizUtils.computeRate(rank.getChemology(), rank);
        String biologicalRankRate = BizUtils.computeRate(rank.getBiological(), rank);
        String politicsRankRate = BizUtils.computeRate(rank.getPolitics(), rank);
        String historyRankRate = BizUtils.computeRate(rank.getHistory(), rank);
        String geographyRankRate = BizUtils.computeRate(rank.getGeography(), rank);
        String wenkeRankRate = BizUtils.computeRate(rank.getWenkeRank(), rank);
        String likeRankRate = BizUtils.computeRate(rank.getLikeRank(), rank);

        List<String> rankList = new ArrayList<String>();
        rankList.add(chineseRankRate);
        rankList.add(mathRankRate);
        rankList.add(englishRankRate);
        rankList.add(physicsRankRate);
        rankList.add(chemologyRankRate);
        rankList.add(biologicalRankRate);
        rankList.add(politicsRankRate);
        rankList.add(historyRankRate);
        rankList.add(geographyRankRate);
        rankList.add(wenkeRankRate);
        rankList.add(likeRankRate);
        return rankList;
    }

    @Override
    void generateHeadData(XSSFRow row) {
        List<String> data = new ArrayList<String>();
        data.add("班级");
        data.add("学号");
        data.add("姓名");
        data.add("语名");
        data.add("数名");
        data.add("英名");
        data.add("物名");
        data.add("化名");
        data.add("生名");
        data.add("政名");
        data.add("历名");
        data.add("地名");
        data.add("三总");
        data.add("3总名");
        data.add("9总分");
        data.add("9总名");
        data.add("文科总分");
        data.add("文科排名");
        data.add("理科总分");
        data.add("理科排名");
        data.add("语文对位率");
        data.add("数学对位率");
        data.add("外语对位率");
        data.add("物理对位率");
        data.add("化学对位率");
        data.add("生物对位率");
        data.add("政治对位率");
        data.add("历史对位率");
        data.add("地理对位率");
        data.add("文科总分对位率");
        data.add("理科总分对位率");
        data.add("进退-期末");
        data.add("进退-一段");
        data.add("班级排名");
        BizUtils.generateHeadData(row,data);
    }
}
