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

import com.github.liaochong.myexcel.core.constant.FileType;
import com.github.liaochong.myexcel.core.constant.LinkType;
import com.github.liaochong.myexcel.core.converter.CustomWriteConverter;
import com.github.liaochong.myexcel.core.converter.DefaultCustomWriteConverter;

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
     * 小数格式化，
     * 已过期，请使用format代替
     *
     * @return 格式化
     */
    @Deprecated
    String decimalFormat() default "";

    /**
     * 时间格式化，如yyyy-MM-dd HH:mm:ss，
     * 已过期，请使用format代替
     *
     * @return 时间格式化
     */
    @Deprecated
    String dateFormatPattern() default "";

    /**
     * 格式化，时间、金额等
     *
     * @return 格式化
     */
    String format() default "";

    /**
     * 样式
     *
     * @return 样式集合
     */
    String[] style() default {};

    /**
     * 链接
     *
     * @return linkType
     */
    LinkType linkType() default LinkType.NONE;

    /**
     * 简单映射，如"1:男,2:女"
     *
     * @return String
     */
    String mapping() default "";

    /**
     * 写转化器
     *
     * @return 映射提供者
     */
    Class<? extends CustomWriteConverter> writeConverter() default DefaultCustomWriteConverter.class;

    /**
     * 文件类型
     *
     * @return 文件类型
     */
    FileType fileType() default FileType.NONE;

    /**
     * 是否为公式
     *
     * @return true/false
     */
    boolean formula() default false;

    /**
     * 提示语
     *
     * @return 提示语
     */
    Prompt prompt() default @Prompt();

    /**
     * 图片相关配置
     *
     * @return 图片配置
     */
    Image image() default @Image();

    /**
     * 下拉列表配置
     *
     * @return 下拉列表配置
     */
    DropdownList dropdownList() default @DropdownList;
}
