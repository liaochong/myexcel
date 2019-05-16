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
package com.github.liaochong.myexcel.core;

import lombok.NonNull;

import java.io.File;
import java.io.InputStream;
import java.util.List;
import java.util.function.Consumer;

/**
 * excel读取接口
 *
 * @author liaochong
 * @version 1.0
 */
public interface ExcelReader<T> {

    /**
     * 指定从那个sheet中导入
     *
     * @param index sheet索引
     * @return ExcelReader<T>
     */
    ExcelReader<T> sheet(int index);

    /**
     * 从文件流中读取
     *
     * @param fileInputStream 文件流
     * @return 结果集
     */
    List<T> read(@NonNull InputStream fileInputStream);

    /**
     * 从文件读取
     *
     * @param file 文件
     * @return 结果集
     */
    List<T> read(@NonNull File file);

    /**
     * 从文件流中读取
     *
     * @param fileInputStream 文件流
     * @param consumer        消费者
     */
    void readThen(@NonNull InputStream fileInputStream, Consumer<T> consumer);

    /**
     * 从文件流中读取
     *
     * @param file     文件
     * @param consumer 消费者
     */
    void readThen(@NonNull File file, Consumer<T> consumer);
}
