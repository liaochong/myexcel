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
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@ExcelTable(sheetName = "人员信息", rowHeight = 50, titleRowHeight = 80)
public class Person {

    @ExcelColumn(title = "基本信息->姓名", index = 0)
    String name;

    @ExcelColumn(title = "基本信息->年龄", index = 1)
    Integer age;

    @ExcelColumn(title = "是否会跳舞", groups = CommonPeople.class, index = 2, mapping = "true:是,false:否")
    boolean dance;

    @ExcelColumn(title = "金钱", format = "#,000.00", index = 3)
    BigDecimal money;

    @ExcelColumn(title = "生日", format = "yyyy-MM-dd HH:mm:ss", index = 4)
    Date birthday;

    @ExcelColumn(title = "当前日期", format = "yyyy/MM/dd", index = 5)
    LocalDate localDate;

    @ExcelColumn(title = "当前时间", format = "yyyy/MM/dd HH:mm:ss", index = 6)
    LocalDateTime localDateTime;
}
