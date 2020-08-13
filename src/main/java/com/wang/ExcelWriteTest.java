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

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Description:
 *              03版 写入文件 时间相对快但是存取行数最多65536行 xls  HSSFWorkbook
 *              07版  时间相对慢 但是没有数量限制 xlsx   XSSFWorkbook
 *              07版s 时间最快 但是耗内存 没有数量限制 xlsx SXSSFWorkbook
 * @Auther: shanpeng.wang
 * @Create: 2020/8/13 10:19
 */
public class ExcelWriteTest {


    final String PATH = "D:\\springcouldProjects\\POI-Temple\\";


    @Test
    public void testWrite03() throws IOException {

        //1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("工作表人数统计");
        //3.创建一行
        Row row1 = sheet.createRow(0);//创建第一行
        //4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("第一行第一列");

        //（1，2）
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("第一行第二列");

        //（2，1）
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("第二行第一列");

        //（2，2）
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表（IO流）  03版本使用的是xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH+"测试表03.xls");

        //写入
        workbook.write(fileOutputStream);
        //关闭
        fileOutputStream.close();

        System.out.println("测试表03.xls 生成完毕");

    }

    @Test
    public void testWrite07() throws IOException {

        //1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("工作表人数统计");
        //3.创建一行
        Row row1 = sheet.createRow(0);//创建第一行
        //4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("第一行第一列");

        //（1，2）
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("第一行第二列");

        //（2，1）
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("第二行第一列");

        //（2，2）
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表（IO流）  07版本使用的是xlsx结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH+"测试表07.xlsx");

        //写入
        workbook.write(fileOutputStream);
        //关闭
        fileOutputStream.close();

        System.out.println("测试表07.xlsx 生成完毕");

    }

    @Test
    public void testWrite03BigData() throws IOException {
        Workbook workbook = new HSSFWorkbook();
//        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        long begin = System.currentTimeMillis();
        for(int rowNum = 0; rowNum < 65536; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream testWrite03BigData = new FileOutputStream("testWrite03BigData.xls");
        workbook.write(testWrite03BigData);
        testWrite03BigData.close();
        long timeMillis = System.currentTimeMillis()-begin;
        System.out.println("03版写入时间："+timeMillis);//03版写入时间：1890
//        System.out.println("07版写入时间："+timeMillis);//07版写入时间：11150
    }

    @Test
    public void testWrite07BigDataS() throws IOException {
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        long begin = System.currentTimeMillis();
        for(int rowNum = 0; rowNum < 65536; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream testWrite03BigData = new FileOutputStream("testWrite07BigDataS.xlsx");
        workbook.write(testWrite03BigData);
        testWrite03BigData.close();
        //关闭临时文件
        ((SXSSFWorkbook) workbook).dispose();
        long timeMillis = System.currentTimeMillis()-begin;
        System.out.println("07S版写入时间："+timeMillis);//07S版写入时间：1124
    }
}
