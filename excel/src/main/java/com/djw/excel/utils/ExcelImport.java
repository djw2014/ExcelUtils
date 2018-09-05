package com.djw.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author djw
 * @Description //TODO
 * @Date 2018/9/5 16:16
 */
public class ExcelImport {

    /**
     * @param dest 表格文件
     * @param cellLength 一行多少个单元格
     * @return 返回list集合
     * @throws Exception
     */
    public static List<Map<Integer,String>> read(File dest, Integer cellLength) throws Exception{
        Workbook wookbook = null;
        FileInputStream fis = null;
        int cellType = 1;
        try {
            fis = new FileInputStream(dest);
            //用HSSF来处理，有异常即为xlsx格式，用XSSF处理
            wookbook = new HSSFWorkbook(fis);//得到工作簿
            cellType = HSSFCell.CELL_TYPE_STRING;
        } catch (Exception e) {
            try {
                fis = new FileInputStream(dest);//这里不创建输入流就会报错stream close
                wookbook = new XSSFWorkbook(fis);
                cellType = XSSFCell.CELL_TYPE_STRING;
            } catch (Exception e1) {
                //返回文件格式错误异常
                throw new Exception("This file is not in excel format");
            }//得到工作簿
        } finally {
            fis.close();
        }
        //得到一个工作表
        Sheet sheet = wookbook.getSheetAt(0);
        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        //要获得属性
        List<Map<Integer,String>> list = new ArrayList<Map<Integer,String>>();
        Map<Integer,String> map = null;
        //获得所有数据
        //从第x行开始获取
        for(int x = 1 ; x <= totalRowNum ; x++){
            map = new HashMap<Integer,String>();
            //获得第i行对象
            Row row = sheet.getRow(x);
            //如果一行里的所有单元格都为空则不放进list里面
            int a = 0;
            for(int y=0;y<cellLength;y++){
                Cell cell = row.getCell(y);
                if(cell == null){
                    map.put(y,"");
                }else{
                    cell.setCellType(cellType);
                    map.put(y, cell.getStringCellValue().toString());
                }
                if(map.get(y)==null||"".equals(map.get(y))){
                    a++;
                }
            }
            if(a!=cellLength){
                list.add(map);
            }
        }
        return list;
    }
}
