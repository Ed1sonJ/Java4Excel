package org.dxj.excel.biz;

import org.apache.poi.xssf.usermodel.XSSFRow;

/**
 * author : duxiji
 * date : 2017/4/27
 * description : 基本的业务
 */
public abstract class BaseBiz {
   abstract void generateHeadData(XSSFRow row);
}
