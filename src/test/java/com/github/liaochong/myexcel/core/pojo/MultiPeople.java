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
 import lombok.Getter;
 import lombok.Setter;

 import java.util.Date;
 import java.util.List;

 /**
  * @author liaochong
  * @version 1.0
  */
 @Setter
 @Getter
 public class MultiPeople {

     @ExcelColumn(title = "姓名")
     String name;

     @MultiColumn(classType = String.class)
     @ExcelColumn(title = "兴趣")
     List<String> tastes;

     @MultiColumn(classType = Date.class)
     @ExcelColumn(title = "时间", format = "yyyy-MM-dd HH:mm:ss")
     List<Date> dates;
 }