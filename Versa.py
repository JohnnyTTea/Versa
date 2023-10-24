#!/usr/bin/env python
# coding: utf-8

# In[13]:


import os
from tkinter import *
from tkinter import filedialog,messagebox,ttk,simpledialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageTk
import pandas as pd
from datetime import date,timedelta,datetime
import csv
import mysql.connector
import numpy as np
from sqlalchemy import create_engine


# In[14]:


VERSION='2.61'
#FONT
TITLE_FONT=("Helvetica bold", 20)
HEADING1=("Helvetica bold", 18)
NORM_FONT = ("Helvetica", 15)
FOOTER= ("Helvetica", 10)
#COLOR
TITLE_COLOR='#999EA2'
## 6A685D #53565C #7A848D #D6DCDB #82898D #C1CCC7 *#9DA8AF
#POSITION
POSITION="1068x681+400+200"
#BACKGROUND IMAGE:
# BG_IMAGE='Y:\IT Related\Versa\\dist\\light-grey-04.jpg'
# SUB_BG_IMAGE='Y:\IT Related\Versa\\dist\\light-grey-18.jpg'
BG_IMAGE='light-grey-04.jpg'
SUB_BG_IMAGE='light-grey-18.jpg'

#CONSTANT
ALL_WEST= '("WA", "OR", "ID", "CA", "NV", "UT", "AZ", "AK", "HI","MT","WY", "CO", "NM", "ND", "SD", "NE", "KS", "OK", "TX")'
ALL_EAST = '("MN", "IA", "MO", "WI", "IL", "MI", "IN", "OH", "MA", "VT","ME", "NH", "RI", "CT", "NJ", "DE", "NY", "PA", "MD", "WV",  "KY", "VA", "AR", "LA", "MS", "TN", "AL", "NC", "SC", "GA", "FL")'
WW = '("WA", "OR", "ID", "CA", "NV", "UT", "AZ", "AK", "HI")'
WM = '("MT", "WY", "CO", "NM", "ND", "SD", "NE", "KS", "OK", "TX")'
EN = '("MN", "IA", "MO", "WI", "IL", "MI", "IN", "OH", "MA", "VT", "ME", "NH", "RI", "CT", "NJ", "DE", "NY", "PA", "MD", "WV")'
ES = '("KY", "VA", "AR", "LA", "MS", "TN", "AL", "NC", "SC", "GA", "FL")'
STATES=[('Ww',WW),('Wm',WM),('En',EN),('Es',ES),('All_West',ALL_WEST),('All_East',ALL_EAST)]
West=["WA", "OR", "ID", "CA", "NV", "UT", "AZ", "AK", "HI","MT","WY", "CO", "NM", "ND", "SD", "NE", "KS", "OK", "TX"]
East=["MN", "IA", "MO", "WI", "IL", "MI", "IN", "OH", "MA", "VT","ME", "NH", "RI", "CT", "NJ", "DE", "NY", "PA", "MD", "WV",  "KY", "VA", "AR", "LA", "MS", "TN", "AL", "NC", "SC", "GA", "FL"]

#DIR
MYREPORT = "MyReport"
# Get the path to the desktop
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
CURRENT_DIR = os.getcwd()

#CURRENT DATE
TODAY_DATE = datetime.now().strftime('%Y.%m.%d')

#DROP DOWN BOX
MONTH_OPTIONS=['1','2','3','4','5','6','7','8','9','10','11','12','24']

#DATA EXPORT
DATA_EXPORT_LIST={'Sales Report':'Generate a monthly sales report as a CSV File for past 13 months.',
                  'Inventory Report':'Generate a monthly inventory of all Items as a CSV file for past 13 months.',
                  'Shippingfee Report': 'Generate a shippingfee report based on FDW/FEDEX Invoice Hitory.',
                  'RMA Report':'Generate a Return Merchandise Authorization report for past 13 months. ',
                  'Order Processing':'Generate a sales order report from last month to present.',
                  'AIS Item Bin':'Return a report contains bin locations for all AIS items where Onhand > 0.',
                  'DTO Item Bin':'Return a report contains bin locations for all DTO items where Onhand > 0.'
                  }


# In[15]:


###DATABASE CONFIGURATION
#AIS Database
config_aisdata={'host':'192.168.26.50',
                'database':'aisdata',
                'user':'root',
                'password':'ais123'
                }
config_aisdata0={'host':'192.168.26.50',
                'database':'aisdata0',
                'user':'root',
                'password':'ais123'
                }
config_aisdata1={'host':'192.168.26.50',
                'database':'aisdata1',
                'user':'root',
                'password':'ais123'
                }
config_aisdata5={'host':'192.168.26.50',
                'database':'aisdata5',
                'user':'root',
                'password':'ais123'
                }
#VERSA Database
config_versa={'host':'192.168.16.31',
              'database':'data1',
              'user':'root',
              'password':'hello123'
             }

##DATABASE QUERY
ORDER_PROCESSING='''
SELECT
saord.Trno AS SalesOrder_No, saord.Trdate AS SalesOrder_Date, saord.Trorig1 AS Sell_Market, saord.Trorig2 AS Sell_Acc, saord.Ordno AS SHIPPER_REF,
saord.Compno AS Cust_ID, saord.Scompany AS ShipTo_CustomerName, saord.Cpono AS Cust_PO, saord.Shipvia AS Ship_VIA, saord.Freight,
saord.Sstatus AS Ship_Status, saord.Sdate AS Ship_Date, saord.Stracno AS Tracking_No,
saordl.Itemno AS Order_Item_SKU, saordl.Ordqty AS Order_QTY,
saord.Company AS BuyerName, saord.Email1 AS BuyerEmail, saord.addr1 AS BuyerAdd1, saord.addr2 AS BuyerAdd2, saord.City AS BuyerCity, saord.State AS BuyerState, saord.Zip AS BuyerZip, saord.Country AS BuyerCountry,
saord.Scompany AS ST_Name, saord.Sphone1 AS ST_Phone, saord.Saddr1 AS ST_Add1, saord.Saddr2 AS ST_Add2, saord.Scity AS ST_City, saord.Sstate AS ST_State, saord.Szip AS ST_Zip, saord.Scountry AS ST_Country,
saord.Subamt AS SUBAMT, saord.Taxamt1 AS TAX, saord.Totamt AS TOTAMT, saordl.Lnno AS LN
FROM saord, saordl
WHERE saord.Trno = saordl.Trno
AND saord.Trdate >= %s
AND saord.Trdate <= %s
ORDER BY saord.Trdate DESC
'''

#REPORT
AIS_ITEM_BIN_REPORT='''
SELECT *
FROM itemwhse
WHERE Whse in('1','2')
AND (Bin LIKE 'Y%' OR Bin LIKE '0%')
AND Onhand >0
ORDER BY
Itemno,
  CASE 
    WHEN Bin LIKE 'Y%' THEN 0 
    WHEN Bin LIKE '0%' THEN 1
    WHEN RIGHT(Bin, 2) IS NOT NULL THEN 2
  END,
Onhand
'''

DTO_ITEM_BIN_REPORT='''
SELECT * 
FROM itemwhse
WHERE Whse ='1'
ORDER BY
Itemno,
  CASE 
    WHEN RIGHT(Bin, 2) IS NOT NULL THEN 2
  END,
Onhand
'''


AIS_ITEM_BIN='''
SELECT Bin,Onhand
FROM itemwhse
WHERE Itemno=%s
AND Whse in('1','2')
AND (Bin LIKE 'Y%' OR Bin LIKE '0%')
AND Onhand >0
ORDER BY
Itemno,
  CASE 
    WHEN Bin LIKE 'Y%' THEN 0 
    WHEN Bin LIKE '0%' THEN 1
    WHEN RIGHT(Bin, 2) IS NOT NULL THEN 2
  END,
Onhand
'''

DTO_ITEM_BIN='''
SELECT Bin,Onhand
FROM itemwhse
WHERE Itemno=%s
AND Whse ='1'
AND Onhand >0
ORDER BY
Itemno,
  CASE 
    WHEN RIGHT(Bin, 2) IS NOT NULL THEN 2
  END,
Onhand
'''



##RMA 
RMA = '''
        SELECT samemo.*, sainvl.*
        FROM samemo
        JOIN sainvl ON samemo.Invno = sainvl.Trno
        WHERE samemo.Trdate > %s
        and samemo.Trdate <  %s
        Order by samemo.Trdate desc
        '''

SAINVL ='''
        select Itemno, count(sainvl.Itemno) AS 'Total Sale' from aisdata1.sainvl
        where Trdate >  %s
        and Trdate < %s
        group by Itemno
        ORDER BY 'Total Sale' DESC
        '''

SAMEMO ='''
        SELECT sainvl.Itemno, COUNT(*) AS 'Total Return'
        FROM aisdata1.samemo
        JOIN aisdata1.sainvl ON samemo.Invno = sainvl.Trno
        where samemo.Trdate > %s
        and samemo.Trdate < %s
        GROUP BY sainvl.Itemno
        ORDER BY 'Total Return' DESC
        '''

#Products
ITEM='''
select Desc1,Desc2,Sboxlen,Sboxwid,Sboxhei,Sboxwei from item
where Itemno =%s
'''
PRICE='''
select Itemno,round(Avg(Price),2) as Price from sainvl
where Itemno=%s
and Trdate >= %s
and Trdate <%s
'''
SHIPPINGFEE='''
SELECT
    Itemno,round(AVG(NetCharge),2) AS NetCharge 
FROM shippingfee
WHERE Itemno=%s
and InvDate >=%s
and InvDate <%s
'''
INVENTORY='''
SELECT
  Itemno,
  SUM(CASE WHEN Whse = '1' THEN Onhand ELSE 0 END) AS CA1,
  SUM(CASE WHEN Whse = '2' THEN Onhand ELSE 0 END) AS GA2,
  SUM(CASE WHEN Whse = '3' THEN Onhand ELSE 0 END) AS FBA
FROM itemwhse
GROUP BY Itemno;
'''
SHIPPING='''
SELECT * FROM shippingfee
'''

#INVOICE
SAINV_SQL='''
SELECT Trno,Ordno,Trdate,Trtime,Trref,Trorig1,Trorig2,Trorigno,Orddate,Totamt,Compno,Company,Addr1,City,State,Zip,Country,Phone1,Shdate,Sstatus,Sdate,Stracno,Ename,Jouno
FROM sainv
WHERE Trdate>=%s
'''
SHIPPING_RECORD='''
SELECT Ordno
FROM shipping_record
WHERE Trdate>=%s
'''

#Daily Processing
FULL_QUERY='''
SELECT
    saordl.Trno AS Ordno,
    saordl.Trdate AS Date,
    saordl.Lnno AS Lno,
    saordl.Itemno AS OrigOrder,
    Sitem1 AS Alt1,
    Sitem2 AS Alt2,
    Sitem3 AS Alt3,
    item1.Adino,
    saordl.Ordqty as Qty,
    saord.Sstate AS State,
    whse1.Bin AS D1Bin,
    whse1.Whse AS D1Whse,
    whse1.Onhand AS D1,
    whse2.Bin AS D5Bin,
    whse2.Onhand AS D5
FROM saordl
LEFT JOIN saord ON saordl.Trno = saord.Trno
LEFT JOIN (
    SELECT Sitem1,Sitem2,Sitem3,Itemno,Adino
    FROM aisdata0.item
) item1 on item1.Itemno=saordl.Itemno
LEFT JOIN(
	SELECT Itemno,Whse,Bin,Onhand
  FROM aisdata0.itemwhse
  WHERE
	Whse in('1','2')
	AND (Bin LIKE 'Y%' OR Bin LIKE '0%')
	AND Onhand >0
	ORDER BY
	Itemno,
	  CASE 
		WHEN Bin LIKE 'Y%' THEN 0 
		WHEN Bin LIKE '0%' THEN 1
		WHEN RIGHT(Bin, 2) IS NOT NULL THEN 2
	  END,
	Onhand
) whse1 on whse1.Itemno=saordl.Itemno
LEFT JOIN(
	SELECT Itemno,Bin,Onhand
	FROM aisdata5.itemwhse
	WHERE Whse ='1'
  AND Onhand >0
	ORDER BY
	Itemno,
	  CASE 
		WHEN RIGHT(Bin, 2) IS NOT NULL THEN 2
	  END,
	Onhand
) whse2 on whse2.Itemno=saordl.Itemno
WHERE saordl.Trno=%s
AND saordl.Itemno not like 'CB%'
'''

DTO_ITEM='''
SELECT Itemno, Cost 
FROM item
WHERE Itemno=%s
'''

#AISDATA1
DDS_FORM='''
SELECT
Trno AS 'Sales Record Number',
Trorigno AS 'Order Number',
concat(Trorig2,'_',Trno) AS 'Buyer Username',
Company AS 'Buyer Name',
Email1 AS 'Buyer Email',
Addr1 AS 'Buyer Address 1',
Addr2 AS 'Buyer Address 2',
City AS 'Buyer City',
State AS 'Buyer State',
Zip AS 'Buyer Zip',
countrycode.Country AS 'Buyer Country',
Scompany AS 'Ship To Name',
Sphone1 AS 'Ship To Phone',
Saddr1 AS 'Ship To Address 1',
Saddr2 AS 'Ship To Address 2',
Scity AS 'Ship To City',
Sstate AS 'Ship To State',
Szip AS 'Ship To Zip',
countrycode.Country AS 'Ship To Country',
Trorigno AS 'Item Number'
FROM saord
LEFT JOIN countrycode 
ON saord.Country=countrycode.Code
WHERE
Trno=%s
'''


# In[16]:


class Create():
    def data(self,db_config, query, params=None):
        try:
            conn = mysql.connector.connect(**db_config)
            with conn.cursor() as cursor:
                if params is not None:
                    cursor.execute(query, params)
                else:
                    cursor.execute(query)
                result = cursor.fetchall()
                df = pd.DataFrame(result, columns=cursor.column_names)
                return df  # Return the result
        except mysql.connector.Error as err:
            messagebox.showinfo('Error',"Unable to connect to the database !")
        finally:
            conn.close()  # Ensure the connection is closed

    def month(self,num):
        months = []
        end_date1 = datetime.now().replace(day=1).strftime('%Y.%m.%d')
        start_date = (datetime.now() - timedelta(days=30*num)).replace(day=1).strftime('%Y.%m.%d')
        current_date = datetime.strptime(start_date, '%Y.%m.%d')
        end_date2 = datetime.strptime(end_date1, '%Y.%m.%d')
        while current_date < end_date2:
            months.append(current_date.strftime('%Y.%m'))
            current_date += timedelta(days=31)
        return months,start_date,end_date1
    
    def file(self,file_name='',final_table=pd.DataFrame()):
        # Create the folder path
        folder_path = os.path.join(DESKTOP, MYREPORT)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        #Save the dataframe in CSV File
        FILE_PATH=folder_path+"/"+ file_name
        #Check if the file exists
        try:
            final_table.to_csv(FILE_PATH,index=False)
        except PermissionError as e:
            messagebox.showerror('Permission Denied', f"Permission denied, please close the file '{file_name}' !")
            return    
        if file_name not in os.listdir(folder_path):
            messagebox.showinfo('Failed',f"'{file_name}' Failed to create !")
        else: 
            messagebox.showinfo('Succeeded',f"'{file_name}' has created ! Please check '{MYREPORT}' directory.")
            
class Report():
    def __init__(self):
        self.selected_report=PRODUCT_GUI.selected_report
        self.create=Create()

    def sales_report(self):
        #set variables
        months,start_date,end_date = self.create.month(13)
        file_name=f'AIS SLS Report {start_date}-{end_date}_{TODAY_DATE}.csv'
        #set params        
        query_params = ', '.join([f"SUM(CASE WHEN DATE_FORMAT(sainvl.Trdate, '%Y.%m') = '{month}' THEN sainvl.Ordqty ELSE 0 END) AS '{month}'" for month in months])
        query = f'''
        SELECT 
            sainvl.Itemno, aisdata0.item.Desc1, aisdata0.item.Desc2,aisdata0.item.Adino, 
            {query_params},
            SUM(sainvl.Ordqty) AS TotalSales,
            AVG(sainvl.price) AS 'Avg(SalePrice)', 
            aisdata0.item.Vendno,aisdata0.item.Bin, aisdata0.item.Onhand,aisdata0.item.Oncom,aisdata0.item.Onord,
            aisdata0.item.Price1,aisdata0.item.Price2,aisdata0.item.Price3,aisdata0.item.Price4,aisdata0.item.Price5
        FROM sainvl
        LEFT JOIN aisdata0.item ON sainvl.Itemno = aisdata0.item.Itemno
        WHERE sainvl.Trdate >= %s 
        AND sainvl.Trdate < %s
        GROUP BY Itemno
        ORDER BY Totalsales DESC
        '''
        #Create Data
        sales_df = self.create.data(config_aisdata1,query,(start_date,end_date))
        inventory_df = self.create.data(config_aisdata0,INVENTORY)
        final_table= pd.merge(sales_df, inventory_df, on='Itemno', how='left')
        #Create File
        self.create.file(file_name,final_table)

    def whse_report(self):
        #set variables
        months,start_date,end_date = self.create.month(13)
        file_name=f'Whse SSI Report {start_date}-{end_date}_{TODAY_DATE}.csv'
        #set params      
        query_params = ', '.join([f"SUM(CASE WHEN sainv.Sstate in {STATES[i][1]} THEN sainvl.Ordqty ELSE 0 END) AS {STATES[i][0]} "for i in range(len(STATES))])
        query=f'''
        select
        aisdata1.sainvl.Itemno,
        item.Desc1, item.Desc2, item.Adino, item.Clength, item.Cwidth, item.Ulength, item.Uwidth, 
        item.Uheight, item.Onhand, item.Oncom, item.Onord,item.Bin,
        {query_params},
        SUM(sainvl.Ordqty) AS 'All'
        FROM aisdata1.sainvl
        LEFT JOIN item On aisdata1.sainvl.Itemno=item.Itemno
        LEFT JOIN aisdata1.sainv On aisdata1.sainv.Trno = aisdata1.sainvl.Trno
        Where aisdata1.sainvl.Trdate >=%s
        and aisdata1.sainvl.Trdate < %s
        group by aisdata1.sainvl.Itemno
        '''
        whse_df = self.create.data(config_aisdata0,query,(start_date,end_date))
        for i in range(len(STATES)):
            whse_df[f'{STATES[i][0]}/All %'] = whse_df.apply(lambda row: str(round(row[f'{STATES[i][0]}'] / row['All']*100,2))+'%' if row['All'] != 0 else 0, axis=1)
        inventory_df = self.create.data(config_aisdata0,INVENTORY)#inventory table
        final_table = pd.merge(whse_df, inventory_df, on='Itemno', how='left')
        final_table=final_table.sort_values(by='All',ascending=False)
        #Create File
        self.create.file(file_name,final_table)

    def shipping_report(self):
        file_name=f'Shippingfee Report 2021-06-2023-07_{TODAY_DATE}.csv'
        final_table=self.create.data(config_versa,SHIPPING)
        self.create.file(file_name,final_table)

    def RMA_report(self):
        months,start_date,end_date = self.create.month(13)
        file_name = f'RMA Report {start_date}-{end_date}_{TODAY_DATE}.xlsx'
        RMA_table = self.create.data(config_aisdata1,RMA,(start_date,end_date))
        sainvl_table = self.create.data(config_aisdata1,SAINVL,(start_date,end_date))
        samemo_table = self.create.data(config_aisdata1,SAMEMO,(start_date,end_date))
        #returned rate table
        returned_rate=pd.merge(left=sainvl_table,right=samemo_table,how='left')
        returned_rate['Total Sale'] = returned_rate['Total Sale'].fillna(0).astype(int)
        returned_rate['Total Return'] = returned_rate['Total Return'].fillna(0).astype(int)
        returned_rate['Percentage'] = (returned_rate['Total Return'] / returned_rate['Total Sale'] * 100).round(2)
        returned_rate = returned_rate.sort_values(by=['Percentage','Total Sale'],ascending=False)
        returned_rate['Percentage'] =returned_rate['Percentage'].astype(str)+'%'
        returned_rate = returned_rate.reset_index(drop=True)

        folder_path = os.path.join(DESKTOP, MYREPORT)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        FILE_PATH=folder_path+"/"+ file_name
        #Check if the file exists
        try:
            with pd.ExcelWriter(FILE_PATH, engine='xlsxwriter') as writer:
                # Save each DataFrame to a separate sheet
                RMA_table.to_excel(writer, sheet_name='Returned Products', index=False)
                returned_rate.to_excel(writer, sheet_name='Return Rate', index=False)
                messagebox.showinfo('Succeeded',f"'{file_name}' has created ! Please check <MyReport> directory.")
        except PermissionError as e:
            messagebox.showerror('Permission Denied', f"Permission denied, please close the file '{file_name}' !")
        return    


    def order_processing(self):
        # Custom formatting function for phone numbers
        end_date=TODAY_DATE
        one_month_ago = datetime.now() - timedelta(days=31)
        start_date = one_month_ago.strftime('%Y.%m.%d')
        file_name=f'Order Processing {start_date}-{end_date}_{TODAY_DATE}.csv'
        final_table=self.create.data(config_aisdata1,ORDER_PROCESSING,(start_date,end_date))
        final_table['ST_Phone'] = final_table['ST_Phone'].apply(lambda x: x[1:] if x.startswith('+') else x)
        self.create.file(file_name,final_table)

    def ais_item_bin(self):
        file_name=f'AIS Item Bin_{TODAY_DATE}.csv'
        final_table=self.create.data(config_aisdata0,AIS_ITEM_BIN_REPORT)
        self.create.file(file_name,final_table)

    def dto_item_bin(self):
        file_name=f'DTO Item Bin_{TODAY_DATE}.csv'
        final_table=self.create.data(config_aisdata5,DTO_ITEM_BIN_REPORT)
        self.create.file(file_name,final_table)

    def generate_report(self):
        if self.selected_report:
            report_name=self.selected_report.pop()
            result = messagebox.askquestion('Warning','This action may take a while, are you sure to proceed?')

            if result=='yes' and report_name=='Sales Report':#Sales Report
                self.sales_report()
            elif result=='yes' and report_name=='Inventory Report':#Whse Report
                self.whse_report()
            elif result=='yes' and report_name=='Shippingfee Report':#Shipping fee
                self.shipping_report()
            elif result=='yes' and report_name=='RMA Report':
                self.RMA_report()
            elif result=='yes' and report_name=='Order Processing':
                self.order_processing()
            elif result=='yes' and report_name=='AIS Item Bin':
                self.ais_item_bin()
            elif result=='yes' and report_name=='DTO Item Bin':
                self.dto_item_bin()
            else:
                messagebox.showwarning('Warning','The selected option can not be found.')

class Product():
    def __init__(self):
        self.input_value=PRODUCT_GUI.input_value
        self.combo_value=PRODUCT_GUI.combobox_value
        self.create=Create()

    def update_value(self):
        if self.combo_value==[]:
            selected_months=12
        else:
            selected_months=int(self.combo_value[-1])
        #Parameters
        months,start_date,end_date =self.create.month(selected_months)
        return_params = ', '.join([f"SUM(CASE WHEN DATE_FORMAT(sainvl.Trdate, '%Y.%m') = '{month}' THEN sainvl.Ordqty ELSE 0 END) AS '{month} return'" for month in months])
        sales_params = ', '.join([f"SUM(CASE WHEN DATE_FORMAT(sainvl.Trdate, '%Y.%m') = '{month}' THEN sainvl.Ordqty ELSE 0 END) AS '{month} sales'" for month in months])
        RETURNS=f'''
        SELECT 
        sainvl.Itemno,
        {return_params},
        SUM(sainvl.Ordqty) AS TotalReturn
        FROM samemo
        LEFT JOIN 
        sainvl ON samemo.Invno = sainvl.Trno
        WHERE sainvl.Trdate >= %s 
        AND sainvl.Trdate < %s
        and Itemno= %s
        GROUP BY sainvl.Itemno
        '''
        SALES=f'''
        SELECT 
        sainvl.Itemno,
        {sales_params},
        SUM(sainvl.Ordqty) AS TotalSales
        FROM sainvl
        WHERE sainvl.Trdate >= %s 
        AND sainvl.Trdate < %s
        and Itemno= %s
        GROUP BY sainvl.Itemno
        '''
        
        item=self.create.data(config_aisdata0,ITEM,(self.input_value[-1],))
        price=self.create.data(config_aisdata1,PRICE,(self.input_value[-1],start_date,end_date))
        shippingfee=self.create.data(config_versa,SHIPPINGFEE,(self.input_value[-1],start_date,end_date))
        sales=self.create.data(config_aisdata1,SALES,(start_date,end_date,self.input_value[-1],))
        returns=self.create.data(config_aisdata1,RETURNS,(start_date,end_date,self.input_value[-1],))
        def update_chart(months,item,sales,returns):
            if len(item)!=0:
                months=[i[-5:] for i in months]
                # Create a figure and a subplot
                fig, ax = plt.subplots(figsize=(6,4.5))
                bar_width = 0.4
                x = np.arange(len(months))
                if len(sales)!=0:
                    sales=[int(i) for i in sales.iloc[0,1:-1].tolist()]
                    sales_bars=ax.bar(x, sales, bar_width, label='Sales', color='lightblue')
                if len(returns)!=0:
                    returns=[int(i) for i in returns.iloc[0,1:-1].tolist()]
                    returns_bars=ax.bar(x, returns, bar_width, label='Returns', color='darkblue')
               
                # Set labels and title
                ax.set_xlabel('Months')
                ax.set_ylabel('Amount')
                ax.set_title('Monthly Sales and Returns')
                ax.set_xticks(x)
                ax.set_xticklabels(months,fontsize=10, rotation=45)
                ax.legend()
                # Add numeric labels above each bar
                if (len(sales)!=0 and len(returns)!=0):
                    for bar in sales_bars + returns_bars:
                        height = bar.get_height()
                        ax.annotate(f'{int(height)}',  # Format the label as desired (e.g., with two decimal places)
                                    xy=(bar.get_x() + bar.get_width() / 2, height),
                                    xytext=(0, 3),  # Adjust the vertical offset of the label
                                    textcoords='offset points',
                                    ha='center', va='bottom')
                        
                return fig
            else:
                return False
        fig=update_chart(months,item,sales,returns)
        return item,price,shippingfee,sales,returns,fig

class Bin():
    def __init__(self):
        self.value=PRODUCT_GUI.bin_value
        self.create=Create()
    def generate_query(self):
        ##search the query based on the input
        value=self.value[-1]
        BIN_SEARCH= f'''
        SELECT * FROM itemwhse
        Where itemwhse.Bin like '{value}%'
        ORDER BY itemwhse.itemno ASC;
        '''
        if len(value)>6:
            messagebox.showerror('Failed','Invalid Bin Value! Keep your input length less than 6!')
            return
        else:
            final_table=self.create.data(config_aisdata0, BIN_SEARCH,)
            file_name=f'Bin of {value}_{TODAY_DATE}.csv'
            self.create.file(file_name,final_table)

class Invoice():
    def __init__(self):
        self.path_list=PRODUCT_GUI.path_list
        self.create=Create()

    def generate_report(self):
        trdate=date.today()-timedelta(14)
        current_time = datetime.now()
        formatted=current_time.strftime('%Y-%m-%d-%H%M')
        tbs = pd.DataFrame()
        filtered_table=pd.DataFrame()
        COLUMNS = ['Ordno', 'Price', 'Tracno','Shipvia','Status']

        for file in self.path_list:
            df = pd.read_csv(file)
            base_name = os.path.basename(file)
            base_name=base_name.upper()
            print(base_name)
            if not any(keyword in base_name for keyword in ['FDW', 'USPS', 'TMS']):
                messagebox.showerror('Error',"Incorrect file! Please choose a 'FDW'/'USPS'/'TMS' file !")
                return
            elif 'FDW' in base_name:
                df['Shipvia']='FDW'
            elif 'USPS' in base_name:
                df['Shipvia']='USPS'
            elif 'TMS' in base_name:
                df['Shipvia']='TMS'
            df['Status']='Unchecked'
            df.columns = COLUMNS
            tbs = pd.concat([tbs, df], ignore_index=True)
            
        tbs = tbs.astype(str)
        tbs = tbs[tbs['Ordno'].apply(lambda x: x.isnumeric() and len(x) == 7)]

        create=Create()
        existing_record=create.data(config_versa,SHIPPING_RECORD,params=(trdate,))
        sql_data=create.data(config_aisdata1,SAINV_SQL,params=(trdate,))#abstract sainv from two weeks ago
        sql_data[['Ordno','Trno','Jouno']] = sql_data[['Ordno','Trno','Jouno']].astype(str)
        final = pd.merge(tbs, sql_data, on='Ordno', how='left')
        final.loc[final['Tracno'] == final['Stracno'], 'Status'] = 'Checked'
        final['Price'].replace('nan','',inplace=True)
        final['UploadDate']=date.today().strftime("%Y-%m-%d")#final table
        unchecked=final[final['Status']=='Unchecked']#unchecked table
        filtered_table= final[~final['Ordno'].isin(existing_record['Ordno'].astype(str))]#import table
        if len(filtered_table)!=0:
            # #Update Database
            db_url = f"mysql://{config_versa['user']}:{config_versa['password']}@{config_versa['host']}/{config_versa['database']}"
            engine = create_engine(db_url)
            table_name = 'shipping_record'
            filtered_table.to_sql(table_name, engine, if_exists='append', index=False, method='multi', chunksize=1000)    

        answer=messagebox.askquestion('Invoice Check',f'Unchecked Items: {len(unchecked)} \nWould you like to continue to export the report?')
        if answer=='yes':
            #Save as report
            folder_path = os.path.join(DESKTOP, MYREPORT)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file_name=f'Uninvoiced Report {formatted}.csv'
            FILE_PATH=folder_path+"/"+ file_name
            try:
                unchecked.to_csv(FILE_PATH,index=False)
            except PermissionError as e:
                messagebox.showerror('Permission Denied', f"Permission denied, please close the file '{file_name}' !")
                return    
            if file_name not in os.listdir(folder_path):
                messagebox.showinfo('Failed',f"'{file_name}' Failed to create !")
            else: 
                messagebox.showinfo('Succeeded',f"'{file_name}' has created ! Please check '{MYREPORT}' directory.")

class DTO_Processing():
    def __init__(self):
        self.path_list=PRODUCT_GUI.path_list
        self.create=Create()

    def create_table(self):
        input_CSV=pd.read_csv(self.path_list.pop())
        if input_CSV.columns[0] !='Ordno':
            messagebox.showinfo('Error',"Unable to find the column: [Ordno] !")
            return
        input_CSV = input_CSV.drop_duplicates()
        columns=['Ordno','Itemno','Qty','Category','PickBin','Stock','Lno','OrigOrder','Date','Alt1','Alt2','Alt3','Adino','State','D1Bin','D1','D5Bin','D5','W/E','AltBin']
        main=pd.DataFrame()

        for i in input_CSV['Ordno'].tolist():
            sales_orders=self.create.data(config_aisdata1,FULL_QUERY,params=(i,))
            sales_orders['W/E'] = ['East' if state in East else 'West' for state in sales_orders['State']]
            if 'West' in sales_orders['W/E'].tolist():
                sales_orders = sales_orders.sort_values(by='D1Whse',ascending=True)
            else:
                sales_orders = sales_orders.sort_values(by='D1Whse',ascending=False)
            sales_orders=sales_orders.drop_duplicates(subset='Lno')
            sales_orders=sales_orders.sort_values(by=['Lno'])
            main=pd.concat([main,sales_orders],axis=0,ignore_index=True)
        
        main['AltBin']=None
        main['Category']=''
        main['PickBin']=''
        main['Stock']=''
        main['Itemno']=''
        for index, row in main.iterrows():
            #D1
            if row['D1Bin']!=None:
                if row['D1Bin'][0]!='Y':
                    main.at[index,'Category']='D1E'
                else:
                    main.at[index,'Category']='D1'
                main.at[index,'PickBin']=row['D1Bin']
                main.at[index,'Stock']=row['D1']
                main.at[index,'Itemno']=row['OrigOrder']
            else :
                if row['Alt1']!=row['Alt2']!=None:
                    for alt in ['Alt1','Alt2']:
                        if row[alt]!='':
                            bin1=self.create.data(config_aisdata0,AIS_ITEM_BIN,(row[alt],))
                            if len(bin1)==0:
                                bin5=self.create.data(config_aisdata5,DTO_ITEM_BIN,(row[alt],))
                                if len(bin5)==0:
                                    if row['D5Bin']!=None and row['Category']=='':#D5
                                        main.at[index,'Category']='D5'
                                        main.at[index,'PickBin']=row['D5Bin']
                                        main.at[index,'Stock']=row['D5']
                                        main.at[index,'Itemno']=row['OrigOrder']
                                    else:
                                        if main.at[index,'Category']=='':
                                            main.at[index,'Itemno']=row['Adino']
                                            main.at[index,'Category']='DDS'
                                else:
                                    if main.at[index,'Category']=='':
                                        #DTO ALT
                                        main.at[index,'AltBin']=alt + ',' + bin5['Bin'][0]+ ',5'
                                        main.at[index,'Category']='D5'+alt
                                        main.at[index,'PickBin']=bin5['Bin'][0]
                                        main.at[index,'Stock']=bin5['Onhand'][0]
                                        main.at[index,'Itemno']=row[alt]
                            else:
                                #AIS ALT
                                if main.at[index,'AltBin']==None:
                                    main.at[index,'AltBin']=alt + ',' + bin1['Bin'][0]+ ',1'
                                    main.at[index,'Category']=('D1'+alt) if bin1['Bin'][0][0]=='Y' else ('D1E'+alt) 
                                    main.at[index,'PickBin']=bin1['Bin'][0]
                                    main.at[index,'Stock']=bin1['Onhand'][0]
                                    main.at[index,'Itemno']=row[alt]
                #D5
                elif row['D5Bin']!=None:
                    main.at[index,'Category']='D5'
                    main.at[index,'PickBin']=row['D5Bin']
                    main.at[index,'Stock']=row['D5']
                    main.at[index,'Itemno']=row['OrigOrder']
                else:
                    main.at[index,'Itemno']=row['Adino']
                    main.at[index,'Category']='DDS'

        main.loc[main['Category'].str[-4:] == 'Alt2', 'Category'] = ''
        main=main[columns]
        main=main.sort_values(by=['Category','Ordno'])
        main.reset_index(drop=True, inplace=True)
        return main

    def create_OPO(self,table):
        #Modify table
        table['Qty']=table['Qty'].astype(int)
        #DTO_USA
        table1=table[table['Category'].str.contains('D5')]
        DTO_USA_INV=pd.DataFrame({'Item No.':table1['Itemno'],'Order Qty':table1['Qty'], 'Whse No.':'PC','Unit Price':''})
        DTO_USA_INV['Unit Price']=[self.create.data(config_aisdata5,DTO_ITEM,params=(i,))['Cost'][0] for i in DTO_USA_INV['Item No.'].tolist()]
        DTO_USA_INV=DTO_USA_INV.groupby('Item No.').agg({'Order Qty': 'sum', 'Whse No.': 'first', 'Unit Price': 'first'}).reset_index()

        #D1OPO
        D1OPO=DTO_USA_INV[['Item No.','Order Qty','Unit Price']]

        #DTO
        table2=table[table['Category'].str.contains('DDS')]
        DDSOPO=pd.DataFrame({'Item No.':table2['OrigOrder'],'Order Qty':table2['Qty'],'Unit Price':''})
        DDSOPO['Unit Price']=[self.create.data(config_aisdata0,DTO_ITEM,params=(i,))['Cost'][0] for i in DDSOPO['Item No.'].tolist()]
        DDSOPO=DDSOPO.groupby('Item No.').agg({'Order Qty': 'sum','Unit Price': 'first'}).reset_index()

        #DDS
        table3=table[table['Category'].str.contains('DDS')]
        table3=table3[['Ordno','OrigOrder','Itemno','Qty']]
        table3['Sale Date']=date.today().strftime('%m/%d/%Y')
        table3['Shipping Service']='DHL'
        new_column_names = {'Ordno': 'Sales Record Number', 'OrigOrder': 'Item Title','Itemno':'Custom Label','Qty':'Quantity'}
        table3.rename(columns=new_column_names, inplace=True)
        table3 = table3.reset_index(drop=True)

        DDS_COLUMN=['Sales Record Number','Order Number','Buyer Username','Buyer Name','Buyer Email','Buyer Note','Buyer Address 1','Buyer Address 2','Buyer City','Buyer State','Buyer Zip','Buyer Country',
    'Ship To Name','Ship To Phone','Ship To Address 1','Ship To Address 2','Ship To City','Ship To State','Ship To Zip','Ship To Country','Item Number','Item Title','Custom Label',
    'Sold Via Promoted Listings','Quantity','Sold For','Shipping And Handling','eBay Collect And Remit Tax Rate','eBay Collect And Remit Tax Type','Seller Collected Tax','eBay Collected Tax',
    'Electronic Waste Recycling Fee','Mattress Recycling Fee','Battery Recycling Fee','White Goods Disposal Tax','Tire Recycling Fee','Additional Fee','Total Price','eBay Collected Tax and Fees Included in Total',
    'Payment Method','Sale Date','Paid On Date','Ship By Date','Minimum Estimated Delivery Date','Maximum Estimated Delivery Date','Shipped On Date','Feedback Left','Feedback Received','My Item Note',
    'PayPal Transaction ID','Shipping Service','Tracking Number','Transaction ID','Variation Details','Global Shipping Program','Global Shipping Reference ID','Click And Collect','Click And Collect Reference Number',
    'eBay Plus','Authenticity Verification Program','Authenticity Verification Status','Authenticity Verification Outcome Reason','Tax City','Tax State','Tax Zip','Tax Country']
        DDS = pd.DataFrame(columns=DDS_COLUMN)
        for i in table3['Sales Record Number'].tolist():
            query=self.create.data(config_aisdata1,DDS_FORM,params=(i,))
            DDS=pd.concat([DDS,query],axis=0,ignore_index=True)

        for col in table3.columns:
            DDS[col]=table3[col]
        DDS['Ship To Phone'] = DDS['Ship To Phone'].str.replace(r'^\+1\s*|\s*ext\.\s*\d*', '', regex=True)
        
        return D1OPO,DTO_USA_INV,DDSOPO,DDS
    
    def save_file(self,path=True,table=pd.DataFrame(), D1OPO=pd.DataFrame(),DTO_USA_SO=pd.DataFrame(),DDSOPO=pd.DataFrame(),DDS=pd.DataFrame()):
        folder_path = os.path.join(DESKTOP, MYREPORT)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        #Save the new file
        if path==True:
            save_path = filedialog.asksaveasfilename(initialdir=folder_path,defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx"),])
            try:
                with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                    # Save each DataFrame to a separate sheet
                    table.to_excel(writer, sheet_name='DTO Order Processing', index=False)
                    D1OPO.to_excel(writer, sheet_name='D1 OPO DTO_USA', index=False)
                    DTO_USA_SO.to_excel(writer, sheet_name='D5 SO', index=False)
                    DDSOPO.to_excel(writer, sheet_name='D1 OPO DTO(DDS)', index=False)
                    DDS.to_excel(writer, sheet_name='DDS', index=False)
                messagebox.showinfo('Succeeded',f"'{save_path}' has created ! Please check <MyReport> directory.")
            except PermissionError as e:
                messagebox.showerror('Permission Denied', f"Permission denied, please close the file '{save_path}' !")
            return    
            
        else:#Save the original file
            save_path=folder_path+'/Daily Process History'
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            name=f'DTO Daily Process Original {TODAY_DATE}.xlsx'
            save_path=save_path+'/'+name
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                # Save each DataFrame to a separate sheet
                table.to_excel(writer, sheet_name='DTO Order Processing', index=False)
                D1OPO.to_excel(writer, sheet_name='D1 OPO DTO_USA', index=False)
                DTO_USA_SO.to_excel(writer, sheet_name='D5 SO', index=False)
                DDSOPO.to_excel(writer, sheet_name='D1 OPO DTO(DDS)', index=False)
                DDS.to_excel(writer, sheet_name='DDS', index=False) 
            
class PRODUCT_GUI():
    path_list=[]
    combobox_value=[]
    bin_value=[]
    input_value=[]
    selected_report=[]#Data report
    
    def __init__(self,init_window_name):
        self.init_window_name=init_window_name
        self.report = Report()
        self.bin=Bin()
        self.invoice=Invoice()
        self.product=Product()
        self.process=DTO_Processing()

    def update_combobox_value(self,event):
        self.combobox_value.append(self.drop.get())

    def on_select(self,event):#Report
        self.selected_index = self.list_box.curselection()
        self.selected_report.append(self.list_box.get(self.selected_index[0]))
        if self.selected_index:
            selected_option = self.list_box.get(self.selected_index[0])
            description = DATA_EXPORT_LIST.get(selected_option, 'No description available')
            self.desc_text.delete(1.0, END)  # Clear previous description
            self.desc_text.insert(1.0, selected_option+':\n'+description)
    
    def setup_window(self,window,name,img):
        #background
        img = Image.open(img)
        img = ImageTk.PhotoImage(img)
        panel = Label(window,image=img)
        panel.image = img
        panel.place(x=0,y=0)
        #Initial Setting
        window.title(name)
        window.geometry(POSITION)
        self.header_label=Label(window,text=name,font=HEADING1,bg=TITLE_COLOR,width=100,height=2)        #Hearder
        self.header_label.pack(side=TOP)
        self.footer_label=Label(window,text='Copyright © 2002-2025 Versa Tech. All Rights Reserved. Version '+VERSION,font=FOOTER,bg=TITLE_COLOR,width=200,height=2)
        self.footer_label.pack(side=BOTTOM)#Footer
        
    def nevigation(self,window,msg='Please ask Administrator for assistance.'):
        def help_msg():
            messagebox.showinfo("Help", msg)
        self.nevig_frame=Frame(window,width=90,height=560, background='#82898D')#Nevigation Frame  relief=RAISED,
        self.nevig_frame.place(relx=0.05, rely=0.515, anchor='center')
        self.help_btn=Button(self.nevig_frame,text='Help',command=help_msg)#exit button
        self.help_btn.place(relx=0.5, rely=0.85, width=80, anchor='center')      
        self.exit_btn=Button(self.nevig_frame,text='Exit',command=window.destroy)#exit button
        self.exit_btn.place(relx=0.5, rely=0.9, width=80, anchor='center')
        return self.nevig_frame
    
    def message(self,window,title='',msg=''):
        self.msg=Toplevel(window)
        self.msg.geometry("400x100+750+500")
        self.msg.grab_set()
        self.msg.title(title)
        label = Label(self.msg, text=msg,font=10)
        label.pack(pady=5)
        button=Button(self.msg,text='OK',width=10,command=self.msg.destroy)
        button.pack(pady=5,side='bottom')

    def set_init_window(self):
        #CONSTANT
        BUTTON_DICT={'Data Report':self.report_window,'Product':self.product_window,
                      'Bin Location':self.bin_location_window,'Invoice Check':self.invoice_window,'Daily Process':self.process_window,'Order Forecast':self.forcast_window,}
        #Set up window:
        self.setup_window(self.init_window_name,'Ikon Report System',BG_IMAGE)
        self.init_window_name.title('Ikon_Report_V'+VERSION)#窗口名
        #Frame for Button
        self.menu_frame=Frame(self.init_window_name,width=400,height=400,background='#82898D')
        self.menu_frame.place(relx=0.5, rely=0.5, anchor=CENTER)
        def create_buttons(frame):
            items_list=list(BUTTON_DICT.items())
            for i in range(3):
                for j in range(3):
                    index = i * 3 + j
                    if index < len(BUTTON_DICT):
                        key, value = items_list[index]
                        button = Button(frame, text=key,command=value,width=20,height=3)
                        button.grid(row=i, column=j, padx=10, pady=10)
        create_buttons(self.menu_frame)

    def report_window(self):
        self.rpt_window=Toplevel(self.init_window_name)
        self.setup_window(self.rpt_window,'Data Report',SUB_BG_IMAGE)
        self.nevig_frame=self.nevigation(self.rpt_window)
        #Navigation:
        self.export_btn=Button(self.nevig_frame,text='Export',command=self.report.generate_report)#exit button
        self.export_btn.place(relx=0.5, rely=0.1, width=80, anchor='center') 
        #Body
        self.body_frame=Frame(self.rpt_window,width=950,height=560,background='#82898D')#Body Frame
        self.body_frame.place(relx=0.55, rely=0.515, anchor='center')
        self.separator = ttk.Separator(self.body_frame, orient='vertical')
        self.separator.place(relx=0.31, rely=0, relwidth=0.005, relheight=1)
        self.list_label=Label(self.body_frame, text = "AIS Data Report List" ,font = NORM_FONT,background='#82898D').place(relx=0.05,rely=0.02)
        self.list_box=Listbox(self.body_frame,width=28,height=25,font=11)
        self.list_box.place(relx=0.02,rely=0.08)
        self.desc_label=Label(self.body_frame, text = "AIS Data Report Description" ,font = NORM_FONT,background='#82898D').place(relx=0.5,rely=0.02)
        self.desc_text=Text(self.body_frame,width=60,height=20,font=11)
        self.desc_text.place(relx=0.37,rely=0.08)
        for i,key in enumerate(DATA_EXPORT_LIST):
            self.list_box.insert(i,key)
        self.list_box.bind("<<ListboxSelect>>", self.on_select)
        self.rpt_window.grab_set()

    def product_window(self):
        #Setup virtualization window
        self.prod_window=Toplevel(self.init_window_name)
        self.setup_window(self.prod_window,'Products',SUB_BG_IMAGE)
        self.nevigation(self.prod_window)
        
        def create_field(frame):        
            FIELD=['Length(in):','Width(in):','Height(in):','Weight(lb):','Price(avg):','Shipfee(avg):','Total Sales:','Total Returns:','Return Rate:']
            text_array=[]
            for i in range((len(FIELD)//2)+1):
                for j in range(2):
                    index = i * 2 + j
                    if index < len(FIELD):
                        label = Label(frame, text=FIELD[index],width=10,height=1)
                        label.grid(row=i, column=2*j,pady=5)
                        FIELD[index] = Text(frame, width=6,height=1)
                        FIELD[index].grid(row=i, column=2*j+1,padx=2,pady=2)
                        # FIELD[index].grid(row=i, column=2*j+1, padx=6, pady=5)
                        text_array.append(FIELD[index])
            return text_array
                                
        def update_product_info():
            #get input
            self.input_value.append(self.input.get().strip())
            item,price,shippingfee,sales,returns,fig=self.product.update_value()
            text_field=[self.text_desc1,self.text_desc2]+self.text_filed
            #warning message
            if len(item)==0:
                self.warning.place(relx=0.09,rely=0.18)
                for i in text_field:
                    i.delete('1.0','end')
            else:
                self.warning.place_forget()
                for i in text_field:
                    i.delete('1.0','end')
                self.text_desc1.insert(END,item['Desc1'][0])#desc1
                self.text_desc2.insert(END,item['Desc2'][0])#desc2
                self.text_filed[0].insert(END,item['Sboxlen'][0])#length
                self.text_filed[1].insert(END,item['Sboxwid'][0])#width
                self.text_filed[2].insert(END,item['Sboxhei'][0])#height
                self.text_filed[3].insert(END,item['Sboxwei'][0])#weight
                self.text_filed[4].insert(END,0 if price['Price'][0]==None else price['Price'][0])#price
                self.text_filed[5].insert(END,0 if shippingfee['Itemno'][0]==None else shippingfee['NetCharge'][0])#shipfee
                self.text_filed[6].insert(END,0 if len(sales)==0 else sales['TotalSales'][0])#Total Sale
                self.text_filed[7].insert(END,0 if len(returns)==0 else returns['TotalReturn'][0])#Total Returns
                self.text_filed[8].insert(END,0 if (len(sales)==0 or len(returns)==0) else str(round((returns['TotalReturn'][0] / sales['TotalSales'][0] * 100),2))+"%")#Return Rate
                #display chart
                self.canvas = FigureCanvasTkAgg(fig, master=self.body_frame)
                self.canvas=self.canvas.get_tk_widget()
                self.canvas.place(relx=0.66, rely=0.515, anchor='center')

        #Frame
        self.body_frame=Frame(self.prod_window,width=950,height=560,background='#82898D')#Body Frame
        self.body_frame.place(relx=0.55, rely=0.515, anchor='center')
        #Separator object
        self.separator = ttk.Separator(self.body_frame, orient='vertical')
        self.separator.place(relx=0.31, rely=0, relwidth=0.005, relheight=1)
        #search bar:
        self.frm_search=Frame(self.body_frame,width=300,height=60)
        self.frm_search.place(relx=0.005,rely=.11)
        self.label=Label(self.frm_search,text='Item ID:').pack(side='left') #Item ID
        entry_text =StringVar()
        self.input=Entry(self.frm_search,width=24,font=('Arial 10'), textvariable=entry_text)
        self.input.pack(side='left',padx=4,pady=5)#Input
        # Function to update the variable with uppercase text
        def update_entry_text(*args):
            entry_text.set(entry_text.get().upper())
        entry_text.trace_add("write", update_entry_text)
        self.search_btn=Button(self.frm_search,width=6,text='search',command=update_product_info)#Button
        self.search_btn.pack(side='right',padx=4,pady=5)
        self.prod_window.bind('<Return>',lambda event:update_product_info())
        self.warning=Label(self.body_frame,text='Invalid Item ID !',font='9',foreground='red',background='#82898D')#warning message
        #Info panel
        self.frm_des=Frame(self.body_frame)
        self.frm_des.place(relx=0.005,rely=.24)
        self.lab_desc1=Label(self.frm_des,text='Description 1:').grid(row=0, column=0)  #DESC1
        self.text_desc1=Text(self.frm_des,width=25,height=2)
        self.text_desc1.grid(row=0, column=1, padx=3, pady=2)
        self.lab_desc2=Label(self.frm_des,text='Description 2:').grid(row=1, column=0) #DESC2
        self.text_desc2=Text(self.frm_des,width=25,height=2)
        self.text_desc2.grid(row=1, column=1, padx=3, pady=2)
        #Dimension
        self.frm_dim=Frame(self.body_frame,highlightthickness=11.5)
        self.frm_dim.place(relx=0.005,rely=0.39)
        #Fill all input
        self.text_filed=create_field(self.frm_dim)
        self.drop_label=Label(self.frm_dim, text = "Select month:" )#Drop down box
        self.drop_label.place(relx=0.64, rely=0.89,anchor='center')
        self.drop= ttk.Combobox(self.frm_dim,values =MONTH_OPTIONS)
        self.drop.current(11)
        self.drop.place(relx=0.89, rely=0.89,width=53,height=22, anchor='center')
        self.drop.bind("<<ComboboxSelected>>", self.update_combobox_value)   

        self.prod_window.grab_set()     

    def bin_location_window(self):
        self.bin_window=Toplevel(self.init_window_name)
        self.setup_window(self.bin_window,'Bin Location Search',SUB_BG_IMAGE)
        self.nevigation(self.bin_window)

        self.dialog_value = simpledialog.askstring("Input", "Enter a Bin Location:")
        if self.dialog_value is None:  # User clicked "Cancel"
            self.bin_window.destroy()
            return
 
        #Text Window
        self.bin_value.append(self.dialog_value)
        self.text_label=Label(self.bin_window,text='Selected Files：',font=NORM_FONT)# Selected Files
        self.text_label.place(relx=0.5, rely=0.2, anchor=CENTER)
        self.my_text=Text(self.bin_window,width=40,height=10,font=("Helvetica", 16), selectbackground="yellow", selectforeground="black", undo=True)
        self.my_text.place(relx=0.5, rely=0.4, anchor=CENTER)
        self.my_text.insert(END,self.dialog_value+'\n')

        #Button
        self.create_btn=Button(self.bin_window,text='Create',command=self.bin.generate_query,width=10)
        self.create_btn.place(relx=0.45, rely=0.6, anchor=CENTER)
        self.exit_btn=Button(self.bin_window,text='Exit',command=self.bin_window.destroy,width=10)
        self.exit_btn.place(relx=0.55, rely=0.6, anchor=CENTER)
        self.bin_window.grab_set()
        
    def invoice_window(self):
        #Setup Invoice window
        self.inv_window=Toplevel(self.init_window_name)
        self.setup_window(self.inv_window,'Invoice Check',SUB_BG_IMAGE)
        msg='Import File:[TMS/USPS/FDW] \nColumns:[Ordno,Price,Tracking]'
        self.nevigation(self.inv_window,msg=msg)

        def open_invoice_csv():
            csv_file= filedialog.askopenfilenames(initialdir=CURRENT_DIR,title='Open CSV File',filetypes=(('CSV Files','*.csv'),))
            for i in csv_file:
                filepath=os.path.abspath(i)
                filename=os.path.basename(filepath)
                if filename[-3:].upper()!='CSV':
                    messagebox.showerror('Error','Incorrect File Names, Please upload CSV files!')
                else:
                    self.path_list.append(filepath)
                    self.my_text.insert(END,filename+'\n')
        def clear():
            self.path_list.clear()
            self.my_text.delete(1.0, END)  # Clear previous description
 
        #Text Window
        self.text_label=Label(self.inv_window,text='Selected Files：',font=NORM_FONT)# Selected Files
        self.text_label.place(relx=0.5, rely=0.2, anchor=CENTER)
        self.my_text=Text(self.inv_window,width=40,height=10,font=("Helvetica", 16), selectbackground="yellow", selectforeground="black", undo=True)
        self.my_text.place(relx=0.5, rely=0.4, anchor=CENTER)

        #Button
        self.open_btn=Button(self.inv_window,text='Open CSV Files',command=open_invoice_csv,width=15)
        self.open_btn.place(relx=0.42, rely=0.6, anchor=CENTER)
        self.clear_btn=Button(self.inv_window,text='Clear',command=clear,width=15)
        self.clear_btn.place(relx=0.58, rely=0.6, anchor=CENTER)

        self.create_btn=Button(self.inv_window,text='Check',command=self.invoice.generate_report,width=10)
        self.create_btn.place(relx=0.45, rely=0.7, anchor=CENTER)
        self.exit_btn=Button(self.inv_window,text='Exit',command=self.inv_window.destroy,width=10)
        self.exit_btn.place(relx=0.55, rely=0.7, anchor=CENTER)
        self.inv_window.grab_set()

    def process_window(self):
        self.procs_window=Toplevel(self.init_window_name)
        self.setup_window(self.procs_window,'Daily Processing',SUB_BG_IMAGE)
        msg='''
        Import File: [Ordno]
        \n
        <Column Names>
        Ordno: Order Number
        Itemno: Suggest Item
        Qty: Order Quantity
        Category: Suggest Company
        PickBin: Suggest Bin
        Lno: List of Each Order
        OrigOrder: Original Order SKU 
        ALt1: Alternative Item 1
        ALt2: Alternative Item 2
        ALt3: Alternative Item 3 
        D1Bin: AIS Bin
        D1: AIS Onhand
        D5Bin: DTO Bin
        D5: DTO Onhand
        AltBin: Available AltBin
        DDS: Taiwan Dropship
        '''
        self.nevigation(self.procs_window,msg=msg)
        self.display_table=[]# Display Data
        self.updated_data=[]#Store updated data
        self.file_path = ""
        self.selected_cols={0: 2,1: 8, 2: 15, 3: 3, 4: 8, 5: 7, 6: 5, 7: 3, 8: 13, 10: 15, 11: 12, 12: 9, 13: 11, 15: 7, 16: 4, 17: 7, 18: 4, 20: 12}

        def on_configure(event):
            # Update the scroll region to cover the entire canvas
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def load_csv():
            csv_file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
            if csv_file_path:
                self.path_list.append(csv_file_path)
                table=self.process.create_table()
                if table is not None:
                    table.insert(0, "0", range(1, len(table) + 1))
                    self.display_table=[table.columns.tolist()] + table.fillna('').values.tolist()
                    # self.display_table=[table.columns.astype(str).tolist()] + table.fillna('').astype(str).values.tolist()
                    display_table()
                    self.download.config(state='normal')
                self.message(self.procs_window,title='Warning',msg='Please review each order before download !')

            #Save the original table
            download_data=self.display_table
            df = pd.DataFrame(download_data[1:], columns=download_data[0])
            if df.columns[0] == '0':
                df = df.drop(df.columns[0], axis=1)
            try:
                D1OPO,DTO_USA_SO,DDSOPO,DDS=self.process.create_OPO(df)
            except IndexError as e:
                messagebox.showerror('Index Error','Please double check your Category !')
            D1OPO,DTO_USA_SO,DDSOPO,DDS=self.process.create_OPO(df)
            self.process.save_file(path=False,table=df,D1OPO=D1OPO,DTO_USA_SO=DTO_USA_SO,DDSOPO=DDSOPO,DDS=DDS)

        def display_table():
            for widget in self.main_frame.winfo_children():
                widget.destroy()

            self.main_frame=Frame(self.procs_window, width=950, height=450,padx=7, pady=12)
            self.main_frame.place(relx=0.55, rely=0.52, anchor='center') 

            # Create a canvas with a specified size
            self.canvas = Canvas(self.main_frame, width=940, height=450)
            self.canvas.pack(side=LEFT, fill=BOTH, expand=True)

            # Create a scrollbar for the canvas
            self.scrollbar = Scrollbar(self.main_frame, command=self.canvas.yview)
            self.scrollbar.pack(side=RIGHT, fill=Y)

            # Bind the canvas to the configure event to update the scroll region
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            self.canvas.bind("<Configure>", on_configure)
            self.canvas.bind("<MouseWheel>", lambda event: self.canvas.yview_scroll(-1 * (event.delta // 120), "units"))

            # Create a frame to contain the widgets
            self.table_frame = Frame(self.canvas)
            self.canvas.create_window((0, 0), window=self.table_frame, anchor="nw")
        
            for i, row in enumerate(self.display_table):
                for j, value in enumerate(row):
                    if j in self.selected_cols:
                        cell = Entry(self.table_frame, width=self.selected_cols[j])
                        cell.insert(0, value)
                        cell.grid(row=i, column=j)
                        cell.config(state='disabled',disabledforeground ='black')
            
            self.button_save.config(state='disabled')
            self.button_modify.config(state='normal')
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            self.canvas.bind("<Configure>", on_configure)

        def modify_table():
            for widget in self.table_frame.winfo_children():
                widget.config(state='normal')

            #change state
            self.button_modify.config(state='disabled')
            self.download.config(state='disabled')
            self.button_save.config(state='normal')
            self.button_cancel.config(state='normal')

            #Clear updated_data
            self.updated_data=[]
        
        def cancel_table():
            display_table()
            self.button_modify.config(state='normal')
            self.button_cancel.config(state='disabled')
            self.button_save.config(state='normal')
            self.download.config(state='normal')

        def save_table():
            for i in range(len(self.display_table)):
                row = []
                for j in range(len(self.display_table[0])):
                    cell_list = self.table_frame.grid_slaves(row=i, column=j)
                    if cell_list:
                        cell = cell_list[0]
                        row.append(cell.get())
                self.updated_data.append(row)
  
            self.message(self.procs_window,title='Success',msg='The table is saved !')
            #change buttons' state
            self.button_modify.config(state='normal')
            self.button_save.config(state='disabled')
            self.download.config(state='normal')
            for widget in self.table_frame.winfo_children():
                widget.config(state='disabled')     

        def download_csv():
            download_data=self.updated_data if self.updated_data!=[] else self.display_table
            df = pd.DataFrame(download_data[1:], columns=download_data[0])
            if any(df['Category'] == ''):
                self.message(self.procs_window,title='Failed',msg='You cannot leave <Category> empty!')
            else:
                if df.columns[0] == '0':
                    df = df.drop(df.columns[0], axis=1)
                try:
                    D1OPO,DTO_USA_SO,DDSOPO,DDS=self.process.create_OPO(df)
                except IndexError as e:
                    messagebox.showerror('Index Error','Please double check your Category !')
                D1OPO,DTO_USA_SO,DDSOPO,DDS=self.process.create_OPO(df)
                self.process.save_file(table=df, D1OPO=D1OPO,DTO_USA_SO=DTO_USA_SO,DDSOPO=DDSOPO,DDS=DDS)

        self.main_frame=Frame(self.procs_window, width=950, height=450)
        self.main_frame.place(relx=0.55, rely=0.52,anchor='center') 

        # Create a canvas with a specified size
        self.canvas = Canvas(self.main_frame, width=940, height=450)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)

        # Create a scrollbar for the canvas
        self.scrollbar = Scrollbar(self.main_frame, command=self.canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        # Bind the canvas to the configure event to update the scroll region
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", on_configure)
        self.canvas.bind("<MouseWheel>", lambda event: self.canvas.yview_scroll(-1 * (event.delta // 120), "units"))
        
        #Create title
        self.title=Label(self.procs_window,text='Daily Sales Order Table Display',font=NORM_FONT)
        self.title.place(relx=0.5,rely=0.13,anchor='center')
        # Create a frame to contain the widgets
        self.table_frame = Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.table_frame, anchor="nw")
        #Impor the table:
        self.import_btn=Button(self.nevig_frame,text='Import',command=load_csv)
        self.import_btn.place(relx=0.5, rely=0.1, width=80, anchor='center') 
        # Modify current table
        self.button_modify = Button(self.nevig_frame, text="Modify",state='disabled',command=modify_table)
        self.button_modify.place(relx=0.5, rely=0.15,width=80,anchor='center') 
        # Cancel current table
        self.button_cancel = Button(self.nevig_frame, text="Cancel",state='disabled',command=cancel_table)
        self.button_cancel.place(relx=0.5, rely=0.20,width=80,anchor='center') 
        # Save current table
        self.button_save = Button(self.nevig_frame, text="Save",state='disabled',command=save_table)
        self.button_save.place(relx=0.5, rely=0.25,width=80,anchor='center') 
        # download the fil0e
        self.download = Button(self.nevig_frame, text="Download",state='disabled',command=download_csv)
        self.download.place(relx=0.5, rely=0.30,width=80,anchor='center') 
        
        self.procs_window.grab_set()

    def forcast_window(self):
        # self.frcs_window=Toplevel(self.init_window_name)
        # self.setup_window(self.frcs_window,'Order Forcasting',SUB_BG_IMAGE)
        # self.nevigation(self.frcs_window)
        messagebox.showinfo("Information", "This feature is not currently available. Please stay tuned.")

def product_start():
    init_window=Tk()
    ZMJ_PROTAL = PRODUCT_GUI(init_window)
    #设置根窗口默认属性
    ZMJ_PROTAL.set_init_window()
    init_window.mainloop()

product_start()

