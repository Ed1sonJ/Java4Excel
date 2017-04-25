package org.dxj.excel.biz;

import com.sun.corba.se.spi.ior.Writeable;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author : duxiji
 * @date : 2017/4/25
 * @description :
 */
public class GenerateBiz {
    public static void generateSheet(String path, String name, List<List<List<String>>> data) {
        try {
            //拿到本次成绩
            List<List<String>> sheet3 = data.get(2);
            int sheetLength = sheet3.size();//41

            InputStream is = new FileInputStream(path);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            //生成表
            XSSFSheet sheet = xssfWorkbook.createSheet(name);
            //获取第0行
            XSSFRow row0 = sheet.createRow(0);
            //生成表头
            row0 = grenerateHead(row0);
            sheet = grenerateBody(sheetLength,sheet);



            OutputStream os = new FileOutputStream(path);
            xssfWorkbook.write(os);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成body
     */
    private static XSSFSheet grenerateBody(int sheetLength,XSSFSheet sheet) {
        for (int i = 1; i < sheetLength; i++){
            sheet.createRow(i);
        }
        return sheet;
    }

    /**
     * 生成head
     *
     * @param row
     * @return
     */
    public static XSSFRow grenerateHead(XSSFRow row) {
        row.createCell(0).setCellValue("班别");
        row.createCell(1).setCellValue("学号");
        row.createCell(2).setCellValue("姓名");
        row.createCell(3).setCellValue("语名");
        row.createCell(4).setCellValue("数名");
        row.createCell(5).setCellValue("英名");
        row.createCell(6).setCellValue("物名");
        row.createCell(7).setCellValue("化名");
        row.createCell(8).setCellValue("生名");
        row.createCell(9).setCellValue("政名");
        row.createCell(10).setCellValue("历名");
        row.createCell(11).setCellValue("地名");
        row.createCell(12).setCellValue("三总");
        row.createCell(13).setCellValue("3总名");
        row.createCell(14).setCellValue("9总分");
        row.createCell(15).setCellValue("9总名");
        row.createCell(16).setCellValue("文科总分");
        row.createCell(17).setCellValue("文科排名");
        row.createCell(18).setCellValue("理科总分");
        row.createCell(19).setCellValue("理科排名");
        row.createCell(20).setCellValue("语文对位率");
        row.createCell(21).setCellValue("数学对位率");
        row.createCell(22).setCellValue("外语对位率");
        row.createCell(23).setCellValue("物理对位率");
        row.createCell(24).setCellValue("化学对位率");
        row.createCell(25).setCellValue("生物对位率");
        row.createCell(26).setCellValue("政治对位率");
        row.createCell(27).setCellValue("历史对位率");
        row.createCell(28).setCellValue("地理对位率");
        row.createCell(29).setCellValue("文科总分对位率");
        row.createCell(30).setCellValue("理科总分对位率");
        row.createCell(31).setCellValue("进退-期末");
        row.createCell(32).setCellValue("进退-一段");
        row.createCell(33).setCellValue("班级排名");
        return row;
    }
}
