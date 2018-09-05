package com.djw.excel.controller;

import com.djw.excel.entiy.User;
import com.djw.excel.utils.ExcelExport;
import com.djw.excel.utils.ExcelImport;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("excel")
@Api(value = "ecxel相关操作",description = "excel相关操作")
public class ExcelController {

//    @GetMapping("getExcel")
//    @ApiOperation(value = "导出Excel",notes = "导出Excel")
//    public void getExcel(HttpServletResponse response) {
//        //excel标题
//        String[] title = {"名称", "性别", "年龄", "学校", "班级"};
//        //excel文件名
//        String fileName = "学生信息表";
//        String sheetName = "学生信息表";
//        String[][] values = {{"张三", "男", "18", "家里蹲大学", "14电信1班"}, {"张三", "男", "18", "家里蹲大学", "14电信1班"}, {"张三", "男", "18", "家里蹲大学", "14电信1班"},
//                {"张三", "男", "18", "家里蹲大学", "14电信1班"}, {"张三", "男", "18", "家里蹲大学", "14电信1班"}, {"张三", "男", "18", "家里蹲大学", "14电信1班"},
//                {"张三", "男", "18", "家里蹲大学", "14电信1班"}};
//        //创建HSSFWorkbook
//        HSSFWorkbook wb = ExcelUtil.createWorkbook(sheetName, title, values);
//        //响应到客户端
//        try {
//            //设置编码、输出文件格式
//            response.reset();
//            fileName = java.net.URLEncoder.encode(fileName, "UTF-8");
//            response.setContentType("application/vnd.ms-excel;charset=UTF-8");
//            response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xls");
//            OutputStream os = response.getOutputStream();
//            wb.write(os);
//            os.flush();
//            os.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }

    @GetMapping("createExcel")
    @ApiOperation(value = "创建excel表格",notes = "创建Excel表格")
    public void createExcel(HttpServletResponse response){
        List<User> userList = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            User user = new User();
            user.setUserId(i);
            user.setUserName("张三"+i);
            user.setSex("男");
            user.setAge(i);
            user.setSchool("学校"+i);
            userList.add(user);
        }
        //excel标题
        String[] title = {"id", "姓名", "性别", "年龄", "学校"};
        String[][] objects = new String[userList.size()][title.length];
        for (int i = 0; i < userList.size(); i++) {
            User user = userList.get(i);
            objects[i][0] = String.valueOf(user.getUserId());
            objects[i][1] = user.getUserName();
            objects[i][2] = user.getSex();
            objects[i][3] = String.valueOf(user.getAge());
            objects[i][4] = user.getSchool();
        }
        String fileName = "学生信息表";
        //创建HSSFWorkbook
        HSSFWorkbook wb = ExcelExport.createWorkbook(fileName, title, objects);
        //响应到客户端
        try {
            //设置编码、输出文件格式
            response.reset();
            fileName = java.net.URLEncoder.encode(fileName, "UTF-8");
            response.setContentType("application/vnd.ms-excel;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xls");
            OutputStream os = response.getOutputStream();
            wb.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @GetMapping("importExcel")
    @ApiOperation(value = "导入Excel",notes = "将数据以Excel的形式导入到数据库中")
    public Boolean importExcel(@RequestParam("file") MultipartFile file, HttpServletRequest request) throws Exception {
        //获取用户上传表格的地址，执行完程序后表格将会删除，避免占用内存
        String filePath = request.getSession().getServletContext().getRealPath("/");
        //根据地址和文件名准确获取用户上传的表格
        File tempFile = new File(filePath+file.getOriginalFilename());
        try {
            //将MultipartFile转换为File类型
            file.transferTo(tempFile);
            List<Map<Integer,String>> dataList = ExcelImport.read(tempFile,5);
            User user = new User();
            for (int i = 0; i < dataList.size(); i++) {
                //此处的取值顺序取决于Excel表的表头顺序，一一对应
                //学生id  姓名  性别  年龄  学校
                user.setUserId(Integer.valueOf(dataList.get(i).get(0)));
                user.setUserName(dataList.get(i).get(1));
                user.setSex(dataList.get(i).get(2));
                user.setAge(Integer.valueOf(dataList.get(i).get(3)));
                user.setSchool(dataList.get(i).get(4));
                //此处应该将user对象插入数据库中
                //UserService.insert(user);
            }
            //执行完程序后删除用户上传文件
            tempFile.delete();
        } catch (Exception e) {
            tempFile.delete();
            throw new Exception("批量录入用户失败，请检查表格中的数据是否和数据库中的数据冲突！");
        }
        return true;
    }
}
