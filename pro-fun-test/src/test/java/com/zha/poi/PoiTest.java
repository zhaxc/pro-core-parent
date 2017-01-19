package com.zha.poi;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.*;
import org.junit.After;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by java on 2017/1/18.
 */
public class PoiTest {

    String filePath="d:\\sample.xls";//文件路径
    HSSFWorkbook workbook = new HSSFWorkbook(); //创建工作表(Sheet)

    /**
     * 创建Excel文件
     */
    @Test
    public void createExcel(){


        HSSFSheet sheet0 = workbook.createSheet();//默认表明Sheet0
        HSSFSheet sheet1 = workbook.createSheet();//默认表明Sheet1
        workbook.createSheet("Test");//指定sheet名

    }

    /**
     * 创建单元格
     */
    @Test
    public void createCell()  {

        HSSFSheet sheet = workbook.createSheet("Test"); // 创建工作表(Sheet)
        HSSFRow row = sheet.createRow(0); // 创建行,从0开始
        HSSFCell cell = row.createCell(0); // 创建行的单元格,也是从0开始
        cell.setCellValue("中国");      // 设置单元格内容
        row.createCell(1).setCellValue(false);
        row.createCell(2).setCellValue(new Date());
        row.createCell(3).setCellValue(12.345);

    }

    /**
     * 创建文档摘要信息
     */
    @Test
    public void createSummaryInformation(){

        workbook.createInformationProperties();//创建文档信息
        DocumentSummaryInformation dsi = workbook.getDocumentSummaryInformation();//文档摘要信息
        dsi.setCategory("类别:Excel文件");//类别
        dsi.setManager("管理者:李志伟");//管理者
        dsi.setCompany("公司:--");//公司
        SummaryInformation si = workbook.getSummaryInformation();//摘要信息
        si.setSubject("主题:--");//主题
        si.setTitle("标题:测试文档");//标题
        si.setAuthor("作者:李志伟");//作者
        si.setComments("备注:POI测试文档");//备注
    }

    /**
     * 创建批注
     * dx1         第1个单元格中x轴的偏移量
     * dy1         第1个单元格中y轴的偏移量
     * dx2         第2个单元格中x轴的偏移量
     * dy2         第2个单元格中y轴的偏移量
     * col1        第1个单元格的列号
     * row1        第1个单元格的行号
     * col2        第2个单元格的列号
     * row2        第2个单元格的行号
     */
    @Test
    public void createDrawingPatriarch(){
        HSSFSheet sheet = workbook.createSheet("Test");// 创建工作表(Sheet)
        HSSFPatriarch patr = sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = patr.createAnchor(0, 0, 0, 0, 5, 1, 8, 3);//创建批注位置
        HSSFComment comment = patr.createCellComment(anchor);//创建批注
        comment.setString(new HSSFRichTextString("这是一个批注段落！"));//设置批注内容
        comment.setAuthor("李志伟");//设置批注作者
        comment.setVisible(true);//设置批注默认显示
        HSSFCell cell = sheet.createRow(2).createCell(1);
        cell.setCellValue("测试");
        cell.setCellComment(comment);//把批注赋值给单元格
    }

    /**
     * 创建页眉页脚
     */
    @Test
    public void createHeaderAndFooter(){
        HSSFSheet sheet = workbook.createSheet("Test");// 创建工作表(Sheet)
        HSSFHeader header =sheet.getHeader();//得到页眉
        header.setLeft("页眉左边");
        header.setRight("页眉右边");
        header.setCenter("页眉中间");
        HSSFFooter footer =sheet.getFooter();//得到页脚
        footer.setLeft("页脚左边");
        footer.setRight("页脚右边");
        footer.setCenter("页脚中间");
    }


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
}
