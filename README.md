# TJU_E-learning_Script
同济大学线上教学中，帮助网络助教工作的一些脚本

另外测试比较少，有Bug请及时联系我

本页面中涵盖了几个Python脚本分别是：
- A:从Course导出课程报表自动生成考勤记录
- B:Course中课程报表过期，找管理员要的.csv格式原始数据刷格式 支持作为A脚本的输入


老本行不是写程序，代码垃圾的一匹，但是应该能用吧(╯‵□′)╯︵┻━┻

*如果有大手子原因搞一搞好欢迎联系我（对不起我实在是太菜了.jpg）*

## Course导出课程报表自动生成考勤记录

暂时只有.py，因为我的pyinstaller不知道为啥抽风了，没办法打包成exe
试图补完中

包：
**import  os 
from openpyxl import load_workbook,Workbook**


### 使用方法
- Step1 新建一个文件夹 把一门课要统计的考勤报表都粘贴进去（只支持新版本的课程报表 大概3/10号之后的报表）
- Step2 把“a姓名学号模板.xlsx”下载到文件夹里，输入这门课的姓名与学号（因为核对过名单，这个应该难度不大，手头都有文件库存）
- Step3 脚本本体Counter.py粘贴到这个文件夹里
- Step4 双击运行 生成一个新的工作簿，里面将有统计信息

### 功能
- 生成报表中包含出勤次数、迟到次数、缺勤次数
- 根据听课时间是否大于91分钟判断是否迟到
- 将迟到的记录输出在备注行
- 按照姓名文字匹配学生 即“123 李狗蛋” “狗蛋 李” “狗 李 蛋”等都被识别为名叫“李狗蛋”的学生
- 自动识别老版本报表，将不计入统计
- 文件名格式化

### 有关老版本报表
老版本报表有如下特征：
- 一位同学掉线了会在表格内记录为两条
- 没有时长这一列
- 表头为：日期	与会者	email	加入时间	离开时间
(因为我懒，于是没有写怎么处理这个东西)

## Course中课程报表过期，找管理员要的.csv格式原始数据刷格式

暂时只有.py，因为我的pyinstaller不知道为啥抽风了，没办法打包成exe
试图补完中

包：
**import  os
import csv
import openpyxl**

### 使用方法
- Step1 新建一个文件夹 把.CSV粘贴进去（只支持新版本的课程报表 大概3/10号之后的报表）
- Step2 脚本本体CSV_formater.py粘贴到这个文件夹里
- Step3 双击运行 生成一个新的工作簿，里面会有格式化好的信息

### 功能
- 生成报表为匹配新版本报表 为：上课时间	参与者 【空列】 时间	加入时间	离开时间
- 以外退出导致的多行数据已经可以处理，删除多余条目的同时，参会时间将会累加
- 生成的报表名称格式化，变成时间+课程名称
- .cvs转存.xlsx
