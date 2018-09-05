package com.djw.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * @Description: 导出Excel表格
 * @version	1.0		2017年5月5日
 * @author jimmy
 */
public class ExcelExport {

    public static Workbook create(String name, String[] entity, List<Map<Integer,Object>> list) throws IOException{
           // create a new workbook
           Workbook wb = new XSSFWorkbook();
           //创建页面对象
           Sheet s = wb.createSheet();
           //编辑第一行的表格头
           //合并单元格，参数分别为开始行，结束行，开始列，结束列
           s.addMergedRegion(new CellRangeAddress(0, 0, 0, entity.length-1));
           //创建行对象，参数即为哪一行
           Row firstRow = s.createRow(0);
           //使用行对象创建单元格对象
           Cell firstCell = firstRow.createCell(0);
           //创建单元格样式对象
           CellStyle firstCs = wb.createCellStyle();
           //创建字体设置对象
           Font firstFont = wb.createFont();
           firstFont.setFontHeightInPoints((short) 16);// 设置字体大小
           firstFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//设置加粗

           //给样式设置字体对象
           firstCs.setFont(firstFont);
           firstCs.setAlignment(HSSFCellStyle.ALIGN_CENTER);//设置居中
           //给单元格设置样式
           firstCell.setCellStyle(firstCs);
           firstCell.setCellValue(name);//赋值
           //表格头编辑结束
           //编辑属性栏样式
           Row entityRow = s.createRow(1);
           Font entityFont = wb.createFont();
           CellStyle entityCs = wb.createCellStyle();
           entityFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//设置加粗
           entityFont.setColor(HSSFColor.WHITE.index);
           entityCs.setAlignment(HSSFCellStyle.ALIGN_CENTER);//设置居中
           entityCs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);   
           entityCs.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);//设置背景颜色
           entityCs.setFont(entityFont);
           //样式编辑结束
           //赋值并设置样式
           for(int i=0;i<entity.length;i++){
               Cell entityCell = entityRow.createCell(i);
               s.setColumnWidth(i,entity[i].toString().length()*512);
               entityCell.setCellValue(entity[i]);
               entityCell.setCellStyle(entityCs);
           }
           //属性栏编辑结束
           //编辑内容
           CellStyle cs = wb.createCellStyle();
           cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
           for(int x=0;x<list.size();x++){
               Row row = s.createRow(x+2);
               for(int y=0;y<entity.length;y++){
                   Cell cell = row.createCell(y);
                   String value = list.get(x).get(y)==null?"":list.get(x).get(y).toString();
                   //设置列宽
                   if(list.get(x).get(y)!=null&&list.get(x).get(y)!=""&&
                           value.length()>entity[y].toString().length()){
                       s.setColumnWidth(y,list.get(x).get(y).toString().length()*512);
                   }
                   cell.setCellValue(value);
                   cell.setCellStyle(cs);
               }
           }
           //内容结束
        return wb;
       }
    
}
