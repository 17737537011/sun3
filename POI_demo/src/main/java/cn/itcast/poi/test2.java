package cn.itcast.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

//解析excel报表
public class test2 {
    public static void main(String[] args) throws Exception {
        //加载工作表
        Workbook workbook=new XSSFWorkbook("D:\\黑马课程\\项目一\\day09\\03-资料\\poi资料\\demo.xlsx");
        //获取sheet
        Sheet sheet = workbook.getSheetAt(0);
        //获取 行
        for(int i=0;i<=sheet.getLastRowNum();i++){//getLastRowNum获取最后一行的索引
            Row row = sheet.getRow(i);
            for(int j=2;i<row.getLastCellNum();j++){//getLastCellNum获取最后一列的列好
                Cell cell = row.getCell(j);
                Object object=getCell(cell);
                System.out.println(object+" ");
            }
        }
        //获取单元格及内容
        //
    }
    public static Object getCell(Cell cell){
        Object obj = null;
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING: { //字符串单元
                obj = cell.getStringCellValue();
                break;
            }
            //excel默认将日志也理解为数字
            case NUMERIC:{ //数字单元格
                if(DateUtil.isCellDateFormatted(cell)) { //日期
                    obj = cell.getDateCellValue();
                }else {
                    obj = cell.getNumericCellValue();
                }
                break;
            }
            case BOOLEAN:{ //boolean
                obj = cell.getBooleanCellValue();
                break;
            }
            default:{
                break;
            }
        }

        return obj;
    }
}
