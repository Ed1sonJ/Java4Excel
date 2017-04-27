package org.dxj.excel.biz;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * author : duxiji
 * date : 2017/4/27
 * description : 生成每科进退表
 */
public class GenerateAdvanceRetreatBiz extends BaseBiz {
    public XSSFWorkbook generateSheet(String path, String name, List<List<String>> agoSheetData, List<List<String>> targetSheetData) {
        try {

            int sheetLength = targetSheetData.size();//41
            InputStream is = new FileInputStream(path);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            //生成进退表
            XSSFSheet parentSheet = xssfWorkbook.createSheet(name);
            //获取第0行
            XSSFRow row0 = parentSheet.createRow(0);
            //生成表头
            generateHeadData(row0);
            //初始化进退表的所有行
            parentSheet = BizUtils.newBody(sheetLength, parentSheet);
            //生成进退表的所有数据
            generateBodyData(parentSheet, agoSheetData, targetSheetData);

            OutputStream os = new FileOutputStream(path);
            xssfWorkbook.write(os);

            is.close();
            os.close();
            return xssfWorkbook;
        } catch (Exception e) {
            e.printStackTrace();

        }
        return null;
    }

    /**
     * 填充sheet的数据
     */
    private void generateBodyData(XSSFSheet parentSheet, List<List<String>> agoSheetData, List<List<String>> targetSheetData) {
        int sheetLength = targetSheetData.size();//41

        for (int i = 1; i <= sheetLength; i++) {
            XSSFRow row = parentSheet.getRow(i);
            //生成每一行的数据
            generateRowData(row, agoSheetData.get(i - 1), targetSheetData.get(i - 1));
        }


    }

    /**
     * 填充行数据
     */
    private void generateRowData(XSSFRow row, List<String> agoRowData, List<String> targetRowData) {
        //填充班别、学号、姓名
        for (int i = 0; i < 3; i++) {
            row.createCell(i).setCellValue(agoRowData.get(i));
        }

        //填充名次差
        int j = 4;
        for (int i = 3; i < 16; i++) {
            //计算名次差
            String resultRank = calculateRank(agoRowData.get(j), targetRowData.get(j));
            row.createCell(i).setCellValue(resultRank);
            j = j + 2;
        }
    }

    /**
     * 计算名次差
     */
    private String calculateRank(String ago, String target) {
        Float agoRank = Float.parseFloat(ago);
        Float targetRank = Float.parseFloat(target);
        Float result = agoRank - targetRank;
        return String.valueOf(result);
    }

    @Override
    void generateHeadData(XSSFRow row) {
        List<String> data = new ArrayList<String>();
        data.add("班级");
        data.add("学号");
        data.add("姓名");
        data.add("语文-期末进退");
        data.add("数学-期末进退");
        data.add("英语-期末进退");
        data.add("物理-期末进退");
        data.add("化学-期末进退");
        data.add("生物-期末进退");
        data.add("政治-期末进退");
        data.add("历史-期末进退");
        data.add("地理-期末进退");
        data.add("3总名-期末进退");
        data.add("9总名-期末进退");
        data.add("文总名-期末进退");
        data.add("理总名-期末进退");
        BizUtils.generateHeadData(row, data);
    }

//    @Override
//    XSSFSheet newBody(int sheetLength, XSSFSheet sheet) {
//        return null;
//    }
}

