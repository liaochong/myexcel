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
import com.github.liaochong.myexcel.core.annotation.Prompt;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class CommonPeople extends Person {

    @ExcelColumn(title = "当前日期", format = "yyyy/MM/dd", index = 5)
    LocalDate localDate;

    @ExcelColumn(title = "当前时间", index = 6)
    LocalDateTime localDateTime;

    @ExcelColumn(prompt = @Prompt(title = "提示", text = "这是我的提示哦"))
    Long cats;
}
