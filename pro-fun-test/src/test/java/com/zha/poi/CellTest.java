package com.zha.poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.After;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by java on 2017/1/19.
 * Excel的单元格操作
 */
public class CellTest {

    String filePath="d:\\sample.xls";//文件路径
    HSSFWorkbook workbook = new HSSFWorkbook(); //创建工作表(Sheet)
    String sheet = "Test";

    /**
     * save File to filePath
     * @throws IOException
     */
    @After
    public void save() throws IOException {
        //保存Excel文件
        FileOutputStream out = new FileOutputStream(filePath);
        workbook.write(out);
        //关闭文件流
        out.close();
        System.out.println("ok!");
    }

    /**
     * 设置格式
     */
    @Test
    public void setFormat(){
        HSSFSheet sheet = workbook.createSheet(this.sheet);
        HSSFRow row = sheet.createRow(0);


        //设置日期格式--使用Excel内嵌的格式
        HSSFCell cell = row.createCell(0);
        cell.setCellValue(new Date());
        HSSFCellStyle style = workbook.createCellStyle();
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
        cell.setCellStyle(style);

        //设置保留2位小数--使用Excel内嵌的格式
        cell=row.createCell(1);
        cell.setCellValue(12.3456789);
        style=workbook.createCellStyle();
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        cell.setCellStyle(style);

        //设置货币格式--使用自定义的格式
        cell=row.createCell(2);
        cell.setCellValue(12345.6789);
        style=workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("￥#,##0"));
        cell.setCellStyle(style);

        //设置百分比格式--使用自定义的格式
        cell=row.createCell(3);
        cell.setCellValue(0.123456789);
        style=workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        cell.setCellStyle(style);

        //设置中文大写格式--使用自定义的格式
        cell=row.createCell(4);
        cell.setCellValue(12345);
        style=workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("[DbNum2][$-804]0"));
        cell.setCellStyle(style);

        //设置科学计数法格式--使用自定义的格式
        cell=row.createCell(5);
        cell.setCellValue(12345);
        style=workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00E+00"));
        cell.setCellStyle(style);
    }


    /**
     * 合并单元格
     */
    @Test
    public void testMergedRegion(){
        HSSFSheet sheet = workbook.createSheet(this.sheet);
        HSSFRow row = sheet.createRow(0);
        //合并列
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("合并列");
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 5);
        sheet.addMergedRegion(region);

        //合并行
        cell = row.createCell(6);
        cell.setCellValue("合并行");
        region = new CellRangeAddress(0, 5, 6, 6);
        sheet.addMergedRegion(region);
    }

}
