package cn.itcast.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

//excel报表
public class test1 {
    public static void main(String[] args) throws Exception {
       Workbook workbook=new XSSFWorkbook();     //2007版本
        Workbook workbook1=new SXSSFWorkbook();  //百万数据上传
        //创建sheet
        Sheet sheet = workbook.createSheet("123");
        //创建行：参数为行下标，从0开始
        Row row = sheet.createRow(2);
        //创建单元格 ：参数为列下标，从0开始
        Cell cell = row.createCell(2);
        //向单元格中写入数据
        cell.setCellValue("运行红河");
        //设置单元格样式
//        CellStyle cellStyle =workbook.createCellStyle();
//        cellStyle.getBorderLeft();
//        cellStyle.getBorderTop();
//        cell.setCellType(cellStyle);
        //创建文件输入流
        FileOutputStream stream = new FileOutputStream("E:\\zzz.xlsx");
        ///写入
        workbook.write(stream);
        workbook.close();
    }

}
