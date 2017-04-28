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
 * description : 生成科任-X的表，其中index标志X科目在sheetScoreData中的index列
 */
public class GenerateBranchBiz extends BaseBiz {
    private int index;
    public static final int CHINESE = 3;
    public static final int MATH = 4;
    public static final int ENGLISH = 5;
    public static final int PHYSICS = 6;
    public static final int CHEMISTRY = 7;
    public static final int BIOLOGICAL = 8;
    public static final int POLITICAL = 9;
    public static final int HISTORY = 10;
    public static final int GEOGRAPHY = 11;


    /**
     * ndex指的是在进退表中科任-X的X所在的列的index
     *
     * @return
     */
    public XSSFWorkbook generateSheet(String path, String name, int index, List<List<String>> sheetScoreData, List<List<String>> sheetQimoData, List<List<String>> sheetYiduanData) {
        try {
            this.index = index;
            int sheetLength = sheetScoreData.size();//41
            InputStream is = new FileInputStream(path);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            //生成科任-XX表
            XSSFSheet parentSheet = xssfWorkbook.createSheet(name);
            //获取第0行
            XSSFRow row0 = parentSheet.createRow(0);
            //生成表头
            generateHeadData(row0);
            //初始化科任-XX表的所有行
            parentSheet = BizUtils.newBody(sheetLength, parentSheet);
            //生成科任-XX表的所有数据
            generateBodyData(parentSheet, sheetScoreData, sheetQimoData, sheetYiduanData);

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
     * 填充sheet数据
     */
    private void generateBodyData(XSSFSheet parentSheet, List<List<String>> sheetScoreData, List<List<String>> sheetQimoData, List<List<String>> sheetYiduanData) {
        int sheetLength = sheetScoreData.size();//41
        for (int i = 1; i <= sheetLength; i++) {
            XSSFRow row = parentSheet.getRow(i);
            //生成每一行的数据
            generateRowData(row, sheetScoreData.get(i - 1), sheetQimoData.get(i - 1), sheetYiduanData.get(i - 1));
        }
    }

    /**
     * 填充每一行的数据
     */
    private void generateRowData(XSSFRow row, List<String> rowScoreData, List<String> rowQimoData, List<String> rowYiduanData) {
        //填充学号、姓名
        for (int i = 0; i < 2; i++) {
            row.createCell(i).setCellValue(rowScoreData.get(i + 1));
        }
        //填充当前科任科目的成绩和名次数据
        int j = (index - 3) * 2 + 3;
        for (int i = 2; i < 4; i++) {
            row.createCell(i).setCellValue(rowScoreData.get(j));
            j++;
        }

        //填充当前科任科目的期末进退，一段进退
        j = index;
        row.createCell(4).setCellValue(rowQimoData.get(index));
        row.createCell(5).setCellValue(rowYiduanData.get(index));

        //填充3总，9总，文科总，理科总进退
        j = 12;
        for (int i = 7; i < 19; i = i + 3) {
            row.createCell(i).setCellValue(rowQimoData.get(j));
            row.createCell(i + 1).setCellValue(rowYiduanData.get(j));
            j++;
        }
        //填充3总，9总，文科总，理科排名
        j = 22;
        for (int i = 6; i < 16; i = i + 3) {
            row.createCell(i).setCellValue(rowScoreData.get(j));
            j = j + 2;
        }


    }

    @Override
    void generateHeadData(XSSFRow row) {
        List<String> data = new ArrayList<String>();
        data.add("学号");
        data.add("姓名");
        switch (index) {
            case CHINESE:
                data.add("语文");
                data.add("语名");
                break;
            case MATH:
                data.add("数学");
                data.add("数名");
                break;
            case ENGLISH:
                data.add("英语");
                data.add("英名");
                break;
            case PHYSICS:
                data.add("物理");
                data.add("物名");
                break;
            case BIOLOGICAL:
                data.add("生物");
                data.add("生名");
                break;
            case POLITICAL:
                data.add("政治");
                data.add("政名");
                break;
            case HISTORY:
                data.add("历史");
                data.add("历名");
                break;
            case GEOGRAPHY:
                data.add("地理");
                data.add("地名");
                break;
        }
        data.add("进退-期末");
        data.add("进退-一段");
        data.add("3总名");
        data.add("进退-期末");
        data.add("进退-一段");
        data.add("9总名");
        data.add("进退-期末");
        data.add("进退-一段");
        data.add("文名");
        data.add("进退-期末");
        data.add("进退-一段");
        data.add("理科排名");
        data.add("进退-期末");
        data.add("进退-一段");
        data.add("班名");
        BizUtils.generateHeadData(row, data);
    }
}
