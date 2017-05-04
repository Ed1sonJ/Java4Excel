package org.dxj.excel.biz;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * author : duxiji
 * date : 2017/4/28
 * description : 生成XX-前X表，继承Base类自动生成表头
 */
public class GeneratePaimingBiz extends BaseBiz {
    public XSSFWorkbook generateSheet(String path, String name, List<List<String>> result) throws IOException {
        int sheetLength = result.size();
        InputStream is = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        //生成XX-前X表
        XSSFSheet parentSheet = xssfWorkbook.createSheet(name);
        //获取第0行
        XSSFRow row0 = parentSheet.createRow(0);
        //生成表头
        generateHeadData(row0);
        //初始化XX-前X的所有行
        parentSheet = BizUtils.newBody(sheetLength, parentSheet);
        //生成XX-前X表的所有数据
        generateBodyData(parentSheet, result);

        OutputStream os = new FileOutputStream(path);
        xssfWorkbook.write(os);

        is.close();
        os.close();
        return xssfWorkbook;

    }

    /**
     * 填充数据
     *
     * @param parentSheet
     * @param result
     */
    private void generateBodyData(XSSFSheet parentSheet, List<List<String>> result) {
        for (int i = 1; i <= result.size(); i++) {
            List<String> rowData = result.get(i - 1);
            XSSFRow row = parentSheet.getRow(i);
            for (int j = 0; j < rowData.size(); j++) {
                row.createCell(j).setCellValue(rowData.get(j));
            }
        }
    }

    @Override
    void generateHeadData(XSSFRow row) {
        List<String> data = new ArrayList<String>();
        data.add("班别");
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
        BizUtils.generateHeadData(row, data);
    }
}
