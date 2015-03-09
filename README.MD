AutoIt开发的小应用
===================================
[威格灵博客：www.cuiweiyou.com](http://www.cuiweiyou.com)<br />


一.中关村在线-投影机数据采集系统
-----------------------------------
* 目录：www_zol_com_cn
* 功能：遍历指定品牌的链接地址，提取指定参数
* 说明：
        1）Barand.txt是品牌集合；Collection.txt是品牌对应网络地址集合；BarandShowArg.txt是要采集的参数集合
		2）本例使用 v3.3.12.0 
		3）Excel2.au3是一个较早版本中的电子表格API；较之于当前版本提供的笔者感觉更好用；tracker.au3是主程序。对于仅安装Office或WPS的环境都适用
		4）核心技术为 ObjCreate("MSXML2.XMLHTTP.3.0") 的使用和 _ExcelXXX 操作Excel的API