/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core.pojo;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.MultiColumn;

import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
public class Grade {
    @ExcelColumn(index = 1)
    String grade;
//    @MultiColumn(classType = Multi1.class)
//    List<Multi1> multi1s;
//    @ExcelColumn(index = 1)
//    String m1;
//    @ExcelColumn(index = 2)
//    String m2;
//    @ExcelColumn(index = 3)
//    String m3;
//    @ExcelColumn(index = 4)
//    String m4;

    @MultiColumn(classType = Student.class)
//    @ExcelColumn(index = 1)
    List<Student> students;

//    @MultiColumn(classType = Integer.class)
//    @ExcelColumn(index = 2)
//    List<Integer> c2;

//    @MultiColumn(classType = Integer.class)
//    @ExcelColumn(index = 3)
//    List<Integer> c3;

    public static class Student {
        @ExcelColumn(index = 2)
        private String name;

        @ExcelColumn(index = 3)
        private Integer age;
    }
}
