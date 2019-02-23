# Html2excel--Excel导出新方式
[![Build Status](https://travis-ci.org/liaochong/html2excel.svg?branch=master)](https://travis-ci.org/liaochong/html2excel)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.liaochong/html2excel/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.github.liaochong/html2excel)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)
[![Average time to resolve an issue](http://isitmaintained.com/badge/resolution/liaochong/html2excel.svg)](http://isitmaintained.com/project/liaochong/html2excel "Average time to resolve an issue")
[![Percentage of issues still open](http://isitmaintained.com/badge/open/liaochong/html2excel.svg)](http://isitmaintained.com/project/liaochong/html2excel "Percentage of issues still open")

> 使用示例参考请移步：[示例](https://github.com/liaochong/html2excel/tree/master/example/src/main/java/com/github/liaochong/example/controller)

简介 | Brief introduction
------------------------
Html2excel，是一个可直接使用Html文件，或者使用内置的Freemarker、Groovy、Beetl等模板引擎Excel构建器生成的Html文件，以Html文件中的Table作为Excel模板来生成任意复杂布局的Excel的工具包，支持.xls、.xlsx格式，支持对背景色、边框、字体等进行个性化设置，支持合并单元格。

Html2excel, is a toolkit that can directly use Html files, or use the built-in Freemarker, Groovy, Beetl and other template engine Excel builder to generate Html files, and use the Table in the Html file as an Excel template to generate Excel of any complex layout. Supports .xls and .xlsx formats, supports personalization of background colors, borders, fonts, etc., and supports merging of cells.

优点 | Advantages
-----------------
- **可生成任意复杂表格**：本工具使用迭代单元格方式进行excel绘制，可生成任意复杂度excel，自适应宽度、高度；
- **零学习成本**：使用html作为模板，学习成本几乎为零；
- **支持常用背景色、边框、字体等样式设置**：具体参见文档-Style-support（样式支持）部分；
- **支持.XLS、.XLSX**：支持生成.xls、.xlsx后缀的excel；
- **支持低内存SXSSF模式**：支持低内存的SXSSF模式，可利用极低的内存生成.xlsx；
- **支持生产者消费者模式导出**：支持生产者消费者模式导出，无需一次性获取所有数据，分批获取数据配合SXSSF模式实现真正意义上海量数据导出；
- **支持多种模板引擎**：已内置Freemarker、Groovy、Beetl等常用模板引擎Excel构建器（详情参见文档[Getting started](https://github.com/liaochong/html2excel/wiki/Getting-started)），默认内置Beetl模板引擎（推荐引擎，[Beetl文档](http://ibeetl.com/guide/#beetl)）；
- **提供默认Excel构建器，直接输出简单Excel**：无需编写任何html，已内置默认模板，可直接根据POJO数据列表输出；
- **支持一次生成多sheet**：以table作为sheet单元，支持一份excel文档中多sheet导出；

文档 | Document
--------------
https://github.com/liaochong/html2excel/wiki

联系以及问题反馈 | Contact me
--------------------------
* QQ群： 135520370
* Email：liaochong8950@163.com
* Issue：[issues](https://github.com/liaochong/html2excel/issues)

> 如本项目对您有所帮助，烦请点击上方star或者在下方扫码支付任意金额以鼓励作者更好地开发，十分感谢您的关注！

<p>
    <img src="https://www.liaochong.site/images/alipay.jpg" height="250"/>
    <img src="https://www.liaochong.site/images/weixin_pay.jpg"  height="250" >
</p>
