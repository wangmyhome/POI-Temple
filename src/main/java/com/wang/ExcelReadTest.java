package com.wang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * @Description:
 *              读取文件
 * @Auther: shanpeng.wang
 * @Create: 2020/8/13 10:19
 */
public class ExcelReadTest {


    final String PATH = "D:\\springcouldProjects\\POI-Temple\\";


    @Test
    public void testRead03() throws IOException {

        FileInputStream fileInputStream = new FileInputStream(PATH + "测试表03.xls");
        //1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheet("工作表人数统计");
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
        Row row1 = sheet.getRow(1);
        Cell cell1 = row1.getCell(1);
        System.out.println(cell1.getStringCellValue());


    }

    @Test
    public void testRead07() throws IOException {

        FileInputStream fileInputStream = new FileInputStream(PATH + "测试表07.xlsX");
        //1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheet("工作表人数统计");
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);

        //特别注意cell.get的值
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
        Row row1 = sheet.getRow(1);
        Cell cell1 = row1.getCell(1);
        System.out.println(cell1.getStringCellValue());

    }



}
