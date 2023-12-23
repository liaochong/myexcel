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

import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.core.strategy.SheetStrategy;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.util.List;
import java.util.Map;

/**
 * excel构建器
 *
 * @author liaochong
 * @version 1.0
 */
public interface ExcelBuilder extends Closeable {

    /**
     * excel类型
     *
     * @param workbookType workbookType
     * @return ExcelBuilder
     */
    ExcelBuilder workbookType(WorkbookType workbookType);

    /**
     * 使用默认样式
     *
     * @return ExcelBuilder
     */
    ExcelBuilder useDefaultStyle();

    /**
     * 应用默认样式
     *
     * @return ExcelBuilder
     */
    ExcelBuilder applyDefaultStyle();

    /**
     * 自动宽度策略
     *
     * @param autoWidthStrategy 策略
     * @return ExcelBuilder
     */
    @Deprecated
    ExcelBuilder autoWidthStrategy(AutoWidthStrategy autoWidthStrategy);

    /**
     * 宽度策略
     *
     * @param widthStrategy 策略
     * @return ExcelBuilder
     */
    ExcelBuilder widthStrategy(WidthStrategy widthStrategy);

    /**
     * Sheet 策略
     *
     * @param sheetStrategy 策略
     * @return ExcelBuilder
     */
    ExcelBuilder sheetStrategy(SheetStrategy sheetStrategy);

    /**
     * 选择固定区域
     *
     * @param freezePanes 固定区域
     * @return ExcelBuilder
     */
    ExcelBuilder freezePanes(FreezePane... freezePanes);

    /**
     * 设置模板，过期，使用classTemplate代替
     *
     * @param path 模板路径
     * @return ExcelBuilder
     */
    @Deprecated
    ExcelBuilder template(String path);

    /**
     * 类路径模板
     *
     * @param path 类路径
     * @return ExcelBuilder
     */
    ExcelBuilder classpathTemplate(String path);

    /**
     * 文件路径模板
     *
     * @param dirPath  文件夹路径
     * @param fileName 模板名称
     * @return ExcelBuilder
     */
    ExcelBuilder fileTemplate(String dirPath, String fileName);

    /**
     * 指定名称管理器
     *
     * @param nameMapping 名称映射
     * @return ExcelFactory
     */
    default ExcelBuilder nameManager(Map<String, List<?>> nameMapping) {
        return this;
    }

    /**
     * 构建
     *
     * @param renderData 渲染数据
     * @param <T>        值类型
     * @return Workbook
     */
    <T> Workbook build(Map<String, T> renderData);
}
