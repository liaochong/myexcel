<p align="center">
    <img src="https://user-images.githubusercontent.com/8674986/154786667-599f1a18-707f-4a08-857f-de97924401ea.png" width="300">
</p>

# MyExcel--Excel操作新方式
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.liaochong/myexcel/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.github.liaochong/myexcel)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)
[![Average time to resolve an issue](http://isitmaintained.com/badge/resolution/liaochong/myexcel.svg)](http://isitmaintained.com/project/liaochong/myexcel "Average time to resolve an issue")
[![Percentage of issues still open](http://isitmaintained.com/badge/open/liaochong/myexcel.svg)](http://isitmaintained.com/project/liaochong/myexcel "Percentage of issues still open")
<img src="https://img.shields.io/badge/JDK-1.8+-green.svg" ></img>
[![GitHub Contributors](https://img.shields.io/github/contributors/liaochong/myexcel)](https://github.com/liaochong/myexcel/graphs/contributors)

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=liaochong/myexcel&type=Date)](https://star-history.com/#liaochong/myexcel&Date)


> 使用示例参考请移步：[示例](https://github.com/liaochong/myexcel/tree/master/example/src/main/java/com/github/liaochong/example/controller)

简介 | Brief introduction
------------------------
MyExcel，是一个集导入、导出、加密Excel等多项功能的工具包。

优点 | Advantages
-----------------
- **可生成任意复杂表格**：本工具使用迭代单元格方式进行excel绘制，可生成任意复杂度excel，自适应宽度、高度；
- **零学习成本**：使用html作为模板，学习成本几乎为零；
- **支持常用背景色、边框、字体等样式设置**：具体参见文档-样式支持部分；
- **支持.xls、.xlsx、.csv**：支持生成.xls、.xlsx后缀的Excel以及.csv文件；
- **支持公式导出**：支持Excel模板中设置公式，降低服务端的计算量；
- **支持低内存SXSSF模式**：支持低内存的SXSSF模式，可利用极低的内存生成.xlsx；
- **支持生产者消费者模式导出**：支持生产者消费者模式导出，无需一次性获取所有数据，分批获取数据配合SXSSF模式实现真正意义上海量数据导出；
- **支持多种模板引擎**：已内置Freemarker、Groovy、Beetl、Thymeleaf等常用模板引擎Excel构建器（详情参见文档[Getting started](https://github.com/liaochong/MyExcel/wiki/Getting-started)），推荐使用Beetl模板引擎（[Beetl文档](http://ibeetl.com/guide/#beetl)）；
- **提供默认Excel构建器，直接输出简单Excel**：无需编写任何html，已内置默认模板，可直接根据POJO数据列表输出；
- **支持一次生成多sheet**：以table作为sheet单元，支持一份excel文档中多sheet导出；
- **支持Excel容量设定**：支持设定Excel容量，到达容量后自动新建Excel，可构建成zip压缩包导出；

文档 | Document
--------------
https://github.com/liaochong/myexcel/wiki

国内用户建议关注作者头条号访问文档，避免github访问不通畅

![å¿æ¬å¨å¨](https://github.com/liaochong/myexcel/assets/8674986/247a4e93-7ff9-488b-a902-893b6aa2fcd5)

或者直接访问文档：

[国内文档地址](https://m.toutiao.com/article_series/7303094494256235008?app=news_article&group_id=7303068015670264320&pseries_style_type=2&pseries_type=0&share_token=150DF85F-4B6F-493E-B0CF-39F9477501F7&tt_from=copy_link&utm_campaign=client_share&utm_medium=toutiao_ios&utm_source=copy_link)

**升级4.x版本注意事项**

 因POI 4.x与5.x版本存在部分不兼容情况，升级MyExcel为4.x（POI 5.x）时，需要注意以下事项：

1. POI版本必须为5.x
2. 排除掉poi-ooxml-schemas依赖（POI 5.x以poi-ooxml-full作为代替）
3. commons-io版本为2.11.0

**Velocity模板引擎注意事项**

自MyExcel 4.0.2版本开始，Velocity依赖修改如下：
```xml
<dependency>
    <groupId>org.apache.velocity</groupId>
    <artifactId>velocity-engine-core</artifactId>
    <version>2.3</version>
</dependency>
```
如仍然使用旧依赖，可能会有编码错乱问题。

联系以及问题反馈 | Contact me
--------------------------
* Email：liaochong8950@163.com
* QQ：1131988645
* Issue：[issues](https://github.com/liaochong/myexcel/issues)

> 如本项目对您有所帮助，烦请点击上方star或者在下方扫码支付任意金额以鼓励作者更好地开发，十分感谢您的关注！

<p>
    <img width="320" alt="image" src="https://user-images.githubusercontent.com/8674986/154786358-8e6c0d45-4a40-45f0-a7ad-6041ada3882e.JPG">
</p>
<p>
    <img width="320" margin-left alt="image" src="https://user-images.githubusercontent.com/8674986/154786504-23538aa4-6ba8-4a4f-a8bd-aed50763d873.JPG">
</p>
