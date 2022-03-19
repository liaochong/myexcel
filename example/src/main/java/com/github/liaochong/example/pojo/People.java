package com.github.liaochong.example.pojo;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;

/**
 * @author liaochong
 * @version 1.0
 */
public class People {

    @ExcelColumn(order = 0, title = "name", index = 0)
    private String name;

    @ExcelColumn(order = 1, title = "age", index = 1)
    private Integer age;

    @ExcelColumn(order = 2, title = "gender", index = 2)
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
