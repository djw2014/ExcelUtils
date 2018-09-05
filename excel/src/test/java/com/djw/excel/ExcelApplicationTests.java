package com.djw.excel;

import com.djw.excel.entiy.User;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelApplicationTests {

    @Test
    public void contextLoads() {

    }

    public static void main(String[] args) {
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
        for (Object[] object : objects) {
            for (Object o : object) {
                System.out.println(o);
            }
        }
    }

}
