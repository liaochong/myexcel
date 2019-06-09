/*
 * Copyright 2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author liaochong
 * @version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
@Documented
public @interface ExcelColumn {
    /**
     * 列标题
     *
     * @return 标题
     */
    String title() default "";

    /**
     * 顺序，数值越大越靠后
     *
     * @return int
     */
    int order() default 0;

    /**
     * 列索引，从零开始，不允许重复
     *
     * @return int
     */
    int index() default -1;

    /**
     * 时间格式化，如yyyy-MM-dd HH:mm:ss
     *
     * @return 时间格式化
     */
    String dateFormatPattern() default "";

    /**
     * 分组
     *
     * @return 分组类类型集合
     */
    Class<?>[] groups() default {};

    /**
     * 为null时默认值
     *
     * @return 默认值
     */
    String defaultValue() default "";

    /**
     * 宽度
     *
     * @return 宽度
     */
    int width() default -1;

    /**
     * 是否强制转换成字符串
     *
     * @return 是否强制转换成字符串
     */
    boolean convertToString() default false;

    /**
     * 自动换行
     *
     * @return 是否自动换行
     */
    boolean wrapText() default false;
}
