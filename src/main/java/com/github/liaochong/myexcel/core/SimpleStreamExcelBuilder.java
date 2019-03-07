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
import java.util.concurrent.ExecutorService;

/**
 * 简单的流式excel构建器
 *
 * @author liaochong
 * @version 1.0
 */
interface SimpleStreamExcelBuilder {

    /**
     * 线程池设置
     *
     * @param executorService 线程池
     * @return SimpleStreamExcelBuilder
     */
    SimpleStreamExcelBuilder threadPool(ExecutorService executorService);

    /**
     * 流式构建启动，包含一些初始化操作
     *
     * @param waitQueueSize 等待队列容量
     * @param groups        分组
     * @return SimpleStreamExcelBuilder
     */
    SimpleStreamExcelBuilder start(int waitQueueSize, Class<?>... groups);

    /**
     * 数据追加
     *
     * @param data 需要追加的数据
     */
    void append(List<?> data);

    /**
     * 停止追加数据，开始构建
     *
     * @return Workbook
     */
    Workbook build();
}
