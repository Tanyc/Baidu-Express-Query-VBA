# Baidu-Express-Query-VBA
百度快递运单号查询接口-Excel通过VBA调用
------
参数来源:[https://github.com/mo10/Baidu-Express-Query](https://github.com/mo10/Baidu-Express-Query)

用Excel表格来管理订单物流数据，还是比较简单便捷普遍的日常应用场景。
这里通过宏实现实时更新物流信息到Excel中，方便跟踪订单物流最新状态，简化仓管及售后人员对订单物流数据跟踪工作。
借助百度搜索的统一物流查询接口，实现小批量、小并发、整合多个物流公司的统一的物流信息更新功能。

同时借助Excel强大的宏处理能力，实现对物流的分类筛选分组导出到独立Sheet或独立Excel文件的管理功能。对有类似分组分类导出需求的，提供一个小小的参考。

# 功能演示
![功能演示](https://raw.githubusercontent.com/Tanyc/Baidu-Express-Query-VBA/master/ExpressQueryOnline.gif)

# WPS 启动宏
WPS默认情况下，“宏”按钮是灰色的，没有启用。可以下载Baidu-Express-Query-VBA/tools目录下的VBA7.0.1590ForWPS.exe安装，开启WPS宏模块功能。