package com.zha.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by java on 2017/1/19.
 */
public class ParserExcel {

    String filePath="d:\\sample.xls";//文件路径
    HSSFWorkbook workbook = new HSSFWorkbook(); //创建工作表(Sheet)
    String sheet = "Test";

    /**
     * 遍历Sheet
     */
    @Test
    public void testParser() throws IOException {
        FileInputStream stream = new FileInputStream(filePath);
        HSSFWorkbook workbook = new HSSFWorkbook(stream);
        HSSFSheet sheet = workbook.getSheet(this.sheet);
        for (Row row : sheet) {
            for (Cell cell : row) {
                System.out.print(cell + "\t");
            }
            System.out.println();
        }
    }
}
