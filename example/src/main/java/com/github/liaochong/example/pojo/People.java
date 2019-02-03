package com.github.liaochong.example.pojo;

import com.github.liaochong.html2excel.core.annotation.ExcelColumn;

/**
 * @author liaochong
 * @version 1.0
 */
public class People {

    @ExcelColumn(order = 1, title = "姓名")
    private String name;

    @ExcelColumn(order = 2, title = "年龄")
    private Integer age;

    @ExcelColumn(order = 3, title = "性别")
    private String gender;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }
}
