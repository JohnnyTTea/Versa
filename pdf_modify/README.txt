-------------------------------------------------------------------------------------------------------------------
Setup：
1. 在文件夹中找到dist文件，在此目录下找到 AIS Report.exe，此文件复制到单独的文件夹中。
3. 在新的文件夹中，创建一个名为Data的目录，用于储存需要分析的数据。

-------------------------------------------------------------------------------------------------------------------
Implementation：
1. 点击Open CSV Files, 在data目录下找到所需要的csv文件，注意这里必须是csv文件，按ctrl可多项选择。将选定的csv文件导入程序中，点击create，即可生成文档。
2. 生成的文档在本文件夹中my_report目录下。

Sale Report:
1. 生成sale report需要一张主表和若干张月表（任意几个月都可）。 
2. 每一张csv表的名字都必须遵从一下规定：
   1）主表名： Item_Qty.CSV(CSV大写)
   2)副表名： Item_Xxx.CSV, Xxx:每个月份的前三个字母，开头大写，example： Item_Jan.CSV, Item_Feb.CSV,etc..

Inventory Report:
0. Inventory的时间范围可在AIS系统中导出的时候定义。
1. 生成inventory report需要一张主表和四张副表，一共需要导入五张表格.
2. 每一张CSV表的名字都必须遵从一下规定：
   1）主表名： Item_Qty_WhseQty.CSV(CSV大写)
   2)副表名： Item_Ww.CSV,Item_Wm.CSV, Item_En.CSV, Item_Es.CSV

Ww (CA) 10 States :Washington, Oregon, Idaho, California, Nevada, Utah, Arizona, Alaska, Hawaii, Outside USA

Wm (TX) 10 States :Montana, Wyoming, Colorado, New Mexico, North Dakota, South Dakota, Nebraska, Kansas, Oklahoma, Texas

En (NJ) 20 States :Minnesota, Iowa, Missouri, Wisconsin, Illinois, Michigan, Indiana, Ohio, Massachusetts, Vermont, Maine, 
New Hampshire, Rhode Island, Connecticut, New Jersey, Delaware, New York, Pennsylvania, Maryland, West Virginia

Es (GA) 11 States :Kentucky, Virginia, Arkansas, Louisiana, Mississippi, Tennessee, Alabama, North Carolina, South Carolina, Georgia, Florida

表命名方式请务必遵循以上原则。

Reference：
Report_Project中data目录下有examples可供参考

-------------------------------------------------------------------------------------------------------------------
How to Export Data:

注：所有的导出文件都将自动储存在本地C磁盘中，Data文件夹中，若没有，请先创建Data文件夹

Item_Qty.CSV: 在AIS系统中点击Data Exchange--> AIS Data Export-->Products(Qty)
请先双击要导出的文件，右边Export Description显示字段后，点击Export-->O.K.

Item_Xxx.CSV： Data Exchange-->Custom Exports-->Item Sales Performance(A,B,C,D).CSV-->Export-->Select Date-->O.K

Item_Qty_WhseQty.CSV: Data Exchange-->AIS Data Export-->双击Products(Qty+Whse Qty)-->Export-->O.K

Item_Xx.CSV: Sales Assistant-->Items Sold by US States--> Sales Dates From/To -->Select ALL:Include Origin 1,DSHIP,EBAY,AMZN,ALL OTHERS--> 
Select States corresponded with CSV Files-->Display-->File

------------------------------------------------SQL Query-----------------------------------------------------------

Select target data--> click Submit--> file will be exported to my_report directory








