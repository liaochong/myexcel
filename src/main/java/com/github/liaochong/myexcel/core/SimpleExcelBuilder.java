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
package com.github.liaochong.myexcel.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

/**
 * 简单excel构建器
 *
 * @author liaochong
 * @version 1.0
 */
interface SimpleExcelBuilder extends ExcelBuilder {

    /**
     * 设置sheet名称
     *
     * @param sheetName sheet名称
     * @return DefaultExcelBuilder
     */
    SimpleExcelBuilder sheetName(String sheetName);

    /**
     * 标题设置
     *
     * @param titles 标题集合
     * @return DefaultExcelBuilder
     */
    SimpleExcelBuilder titles(List<String> titles);

    /**
     * 设置字段的展示顺序
     *
     * @param fieldDisplayOrder 展示的字段集合
     * @return DefaultExcelBuilder
     */
    SimpleExcelBuilder fieldDisplayOrder(List<String> fieldDisplayOrder);

    /**
     * 无样式
     *
     * @return SimpleExcelBuilder
     */
    SimpleExcelBuilder noStyle();

    /**
     * 固定标题
     *
     * @return SimpleExcelBuilder
     */
    SimpleExcelBuilder fixedTitles();

    /**
     * 设置宽度
     *
     * @param widths 宽度集合
     * @return SimpleExcelBuilder
     */
    SimpleExcelBuilder widths(int... widths);

    /**
     * 根据指定的数据集合构建，需指明数据集合数据的类类型，使用该方法，如设定了标题但无数据，则标题行也不展示
     *
     * @param data   数据列表
     * @param groups 分组
     * @return Workbook
     */
    Workbook build(List<?> data, Class<?>... groups);

    @Override
    default ExcelBuilder useDefaultStyle() {
        throw new UnsupportedOperationException();
    }

    @Override
    default ExcelBuilder freezePanes(FreezePane... freezePanes) {
        throw new UnsupportedOperationException();
    }

    @Override
    default ExcelBuilder template(String path) {
        throw new UnsupportedOperationException();
    }

    @Override
    default <T> Workbook build(Map<String, T> renderData) {
        throw new UnsupportedOperationException();
    }
}
