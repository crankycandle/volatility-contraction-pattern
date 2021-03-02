
Save New Duplicate & Edit Just Text Twitter
>
from pandas_datareader import data as pdr
from yahoo_fin import stock_info as si
from pandas import ExcelWriter
import yfinance as yf
import pandas as pd
import requests
from datetime import datetime, timedelta, date
import time
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import seaborn
import mplfinance as mpf
import xlrd
import img2pdf
import os, fnmatch
from pathlib import Path
from PIL import Image
import shutil
from PyPDF2 import PdfFileMerger, PdfFileReader
from fpdf import FPDF
import os.path
import gc
import sys
import subprocess
import json
from simplejson import JSONDecodeError
import math
from matplotlib.backends.backend_pdf import PdfPages
import glob

def spawn_program_and_die(program, exit_code=0):    # 自殺再啟動用
    """
    Start an external program and exit the script
    with the specified return code.

    Takes the parameter program, which is a list
    that corresponds to the argv of your command.
    """
    # Start the external program
    subprocess.Popen(program)
    # We have started the program, and can suspend this interpreter
    sys.exit(exit_code)

os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/')   # 轉Directory 去有pdf既位置
listOfFiles = []

Mon = 0
Tue = 0
Wed = 0
Thu = 0
Fri = 0

try:
    for name in glob.glob("????-??-??.pdf"):     # 將file名成日期, 再轉成星期一二三四五
        filename = os.path.splitext(name)[0]
        weekday = datetime.strptime(filename, '%Y-%m-%d').strftime('%a')
        if weekday == "Mon":
            Mon += 1
        elif weekday == "Tue":
            Tue += 1
        elif weekday == "Wed":
            Wed += 1
        elif weekday == "Thu":
            Thu += 1
        elif weekday == "Fri":
            Fri += 1
except Exception as e:
    print(e)
print('Mon:', Mon, '  Tue:', Tue, '  Wed:', Wed,'  Thu:', Thu, '  Fri:', Fri )
week_day_number = -4
if Mon == Tue == Wed == Thu == Fri:            #找出那一個weekday 最少, 由佢開始做
    week_day_number = -4
elif Fri == min(Mon,Tue, Wed, Thu, Fri):
    week_day_number = -4
elif Thu == min(Mon,Tue, Wed, Thu, Fri):
    week_day_number = -3
elif Wed == min(Mon,Tue, Wed, Thu, Fri):
    week_day_number = -2
elif Tue == min(Mon,Tue, Wed, Thu, Fri):
    week_day_number = -1
elif Mon == min(Mon,Tue, Wed, Thu, Fri):
    week_day_number = 0


last_trade_day = date.today() - timedelta(hours = 9)      #定義上一個交易日既日期, 最少要過左上一個香港日期凌晨9個鐘
offset = (last_trade_day.weekday() + week_day_number)%7
last_weekday = last_trade_day - timedelta(days=offset)      #定義最少weekday 既對上一個既日期

working = True
i = 0
while working == True:
    try:
        date_study = last_weekday - i*timedelta(days = 7)    #向前找出未下載資料既weekday, 最初值 i=0
        i += 1
        yf.pdr_override()
        outputPath = Path("/Users/XXXXX/Dropbox/VCP/Screen_data/")
        test = os.listdir(outputPath)
        for item in test:
            if item.endswith(".png"):
                os.remove(os.path.join(outputPath, item))
            if item.endswith(".jpg"):
                os.remove(os.path.join(outputPath, item))
        os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/')
        if os.path.isfile(f'{date_study}.pdf'):              #萬一出現重複值, 自動continue 跳出迥圈, 回到上一個try,向前找出未下載資料既weekday
            continue
        else:
            pass                                          #如果沒有, 就繼續
        os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/python')
        try:
            os.remove("stocks.csv")
            time.sleep(1)
        except Exception as e:
            print(e)
        data = pd.read_csv("companylist.csv", header=0)
        stocklist = list(data.Symbol)
        final = []
        index = []
        rs = []
        n = -1
        adv = 0
        decl = 0
        adv_w = 0
        decl_w = 0
        c_20 = 0
        c_50 = 0
        s_50_200 = 0
        s_200_200_20 = 0
        s_50_150_200 = 0
        index_list = []
        stocks_fit_condition = 0
        stock_name = []
        ipo_name = []
        ipo_date_list = []
        ipo_list = []
        exportList = pd.DataFrame(columns=['Stock', "RS_Rating", "50 Day MA", "150 Day Ma", "200 Day MA", "52 Week Low", "52 week High"])

        for stock in stocklist:
            n += 1
            start_date = date_study - timedelta(days=365)
            end_date = date_study
            df = pdr.get_data_yahoo(stock, start=start_date, end=end_date)
            time.sleep(0.01)

            try:
                currentClose = df["Adj Close"][-1]
                lastweekClose = df["Adj Close"][-5]
                if (currentClose > lastweekClose):
                    adv_w += 1
                else:
                    decl_w +=1
            except Exception as e:
                print(e)
            try:
                currentClose = df["Adj Close"][-1]
                yesterdayClose = df["Adj Close"][-2]
                if (currentClose > yesterdayClose):
                    adv += 1
                else:
                    decl +=1
            except Exception as e:
                print(e)
            try:
                close_3m = df["Adj Close"][-63]
            except Exception as e:
                print(e)
            try:
                close_6m = df["Adj Close"][-126]
            except Exception as e:
                print(e)
            try:
                close_9m = df["Adj Close"][-189]
            except Exception as e:
                print(e)
            try:
                close_12m = df["Adj Close"][-250]
                condition_ipo = False
            except Exception as e:
                condition_ipo = True                # Consider it is an IPO as no price can be found 1 year ago
                print(e)
            try:
                turnover = df["Volume"][-1]*df["Adj Close"][-1]
            except Exception as e:
                print(e)
            try:
                true_range_10d = (max(df["Adj Close"][-10:-1])-min(df["Adj Close"][-10:-1]))
            except Exception as e:
                print(e)
            try:
                RS_Rating = (((currentClose - close_3m)/close_3m) * 40 + ((currentClose - close_6m)/close_6m) * 20 + ((currentClose - close_9m)/close_9m) * 20 +((currentClose - close_12m)/close_12m) * 20)
            except Exception as e:
                print(e)
            try:
                sma = [20, 50, 150, 200]
                for x in sma:
                    df["SMA_"+str(x)] = round(df.iloc[:,4].rolling(window=x).mean(), 2)

                moving_average_20 = df["SMA_20"][-1]
                moving_average_50 = df["SMA_50"][-1]
                moving_average_150 = df["SMA_150"][-1]
                moving_average_200 = df["SMA_200"][-1]
                low_of_52week = min(df["Adj Close"][-260:])
                high_of_52week = max(df["Adj Close"][-260:])

                try:
                    moving_average_200_20 = df["SMA_200"][-32]

                except Exception:
                    moving_average_200_20 = 0

                # Condition 1: Current Price > 150 SMA and > 200 SMA
                if(currentClose > moving_average_150 > moving_average_200):
                    condition_1 = True
                else:
                    condition_1 = False
                # Condition 2: 50 SMA and > 200 SMA
                if(moving_average_50 > moving_average_200):
                    condition_2 = True
                    s_50_200 += 1
                else:
                    condition_2 = False
                # Condition 3: 200 SMA trending up for at least 1 month (ideally 4-5 months)
                if(moving_average_200 > moving_average_200_20):
                    condition_3 = True
                    s_200_200_20 += 1
                else:
                    condition_3 = False
                # Condition 4: 50 SMA> 150 SMA and 50 SMA> 200 SMA
                if(moving_average_50 > moving_average_150 > moving_average_200):
                    #print("Condition 4 met")
                    condition_4 = True
                    s_50_150_200 += 1
                else:
                    #print("Condition 4 not met")
                    condition_4 = False
                # Condition 5: Current Price > 50 SMA
                if(currentClose > moving_average_50):
                    condition_5 = True
                    c_50 += 1
                else:
                    condition_5 = False
                # Condition 6: Current Price is at least 40% above 52 week low (Many of the best are up 100-300% before coming out of consolidation)
                if(currentClose >= (1.4*low_of_52week)):
                    condition_6 = True
                else:
                    condition_6 = False
                # Condition 7: Current Price is within 25% of 52 week high
                if(currentClose >= (.75*high_of_52week)):
                    condition_7 = True
                else:
                    condition_7 = False
                # Condition 8: Turnover is larger than 1.5 million
                if(turnover >= 1500000):
                    condition_8 = True
                else:
                    condition_8 = False
                # Condition 9: true range in the last 10 days is less than 10% of current price
                if(true_range_10d < currentClose*0.08):
                    condition_9 = True
                else:
                    condition_9 = False
                # Condition 10: Close above 20 days moving average
                if(currentClose > moving_average_20):
                    c_20 += 1
                else:
                    condition_10 = False

                # Condition 11: true range in the last 5 days is less than 6% of current price
                #if(true_range_5d < currentClose*6):
                    #condition_11 = True
                #else:
                    #condition_11 = False

                # Condition 12: Current price > 10
                if(10 < currentClose):
                    condition_12 = True
                else:
                    condition_12 = False


                if(condition_1 and condition_2 and condition_3 and condition_4 and condition_5 and condition_6 and condition_7 and condition_8 and condition_9 and condition_12):
                    final.append(stock)
                    index.append(n)
                    rs.append(RS_Rating)
                    stocks_fit_condition += 1
                    dataframe = pd.DataFrame(list(zip(final, index, rs)), columns =['Company', 'Index', 'RS_Rating'])

                    exportList = exportList.append({'Stock': stock, "RS_Rating": RS_Rating ,"50 Day MA": moving_average_50, "150 Day Ma": moving_average_150, "200 Day MA": moving_average_200, "52 Week Low": low_of_52week, "52 week High": high_of_52week}, ignore_index=True)
                    print(stock + " made the requirements")
                    print(date_study)

            except Exception as e:
                print(e)
                print("No data on "+ stock)

        try:
            dataframe.to_csv('stocks.csv')
        except Exception as e:
            print(e)
            print("No stock fit the condition!")

        print(exportList)

        if not os.path.exists('stocks.csv'):
            dataframe = pd.DataFrame(list(zip(final, index, rs)), columns =['Company', 'Index', 'RS_Rating'])
            dataframe.to_csv('stocks.csv')

        df = pd.read_csv (open('stocks.csv'),index_col=0)
        df['RS Rank'] = round((df.RS_Rating.rank(pct = True)*1),3)-0.0167
        df.sort_values(by=['RS Rank'], axis=0, inplace=True, ascending=False)
        print(df)
        try:
            df.to_csv('stocks.csv', sep=',', encoding='utf-8')
            df = pd.read_csv("stocks.csv", usecols=[1,4],names=['Company','RS Rank'],header = 0)
        except:
            df = pd.DataFrame(columns=['Company','RS Rank'])
        time.sleep(2)

        print(df)

        for index, row in df.iterrows():
            time.sleep(0.01)
            try:
                shares = yf.Ticker(row['Company'])
                name = row['Company']
                RS_rank = round(1 - row['RS Rank'],3)
                RS_rank_title = round(row['RS Rank'],3)
                hist = shares.history(start = start_date, end = end_date)
                time.sleep(0.1)
                filename = f"{RS_rank}_{name}"
                titlename = f"{RS_rank_title}_{name}"
                kwargs = dict(type='candle',mav=(20,50,200),volume=True,figratio=(20,12),figscale=0.61)
                if (name == 'VOO'):
                    mpf.plot(hist,**kwargs,style='brasil',title=titlename, savefig=dict(fname='/Users/XXXXX/Dropbox/VCP/Screen_data/'+"{}.png".format(filename),dpi=100,pad_inches=0.25))
                    index_list.append(name)
                    plt.close()
                elif (name == 'QQQ'):
                    mpf.plot(hist,**kwargs,style='blueskies',title=titlename, savefig=dict(fname='/Users/XXXXX/Dropbox/VCP/Screen_data/'+"{}.png".format(filename),dpi=100,pad_inches=0.25))
                    index_list.append(name)
                    plt.close()
                elif (name == 'DIA'):
                    mpf.plot(hist,**kwargs,style='mike',title=titlename, savefig=dict(fname='/Users/XXXXX/Dropbox/VCP/Screen_data/'+"{}.png".format(filename),dpi=100,pad_inches=0.25))
                    index_list.append(name)
                    plt.close()
                elif (name == 'IWM'):
                    mpf.plot(hist,**kwargs,style='classic',title=titlename, savefig=dict(fname='/Users/XXXXX/Dropbox/VCP/Screen_data/'+"{}.png".format(filename),dpi=100,pad_inches=0.25))
                    index_list.append(name)
                    plt.close()
                elif (name == 'FFTY'):
                    mpf.plot(hist,**kwargs,style='classic',title=titlename, savefig=dict(fname='/Users/XXXXX/Dropbox/VCP/Screen_data/'+"{}.png".format(filename),dpi=100,pad_inches=0.25))
                    index_list.append(name)
                    plt.close()
                elif (RS_rank_title > 0.6831):
                    mpf.plot(hist,**kwargs,style='starsandstripes',title=titlename, savefig=dict(fname='/Users/XXXXX/Dropbox/VCP/Screen_data/'+"{}.png".format(filename),dpi=100,pad_inches=0.25))
                    stock_name.append(name)
                    plt.close()
                else:
                    print("{} is too weak!".format(name))

            except:
                print("{} is not found!".format(name))
                plt.close()

        os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/')
        inputPath = Path("/Users/XXXXX/Dropbox/VCP/Screen_data/")
        inputFiles = inputPath.glob("**/*.png")
        outputPath = Path("/Users/XXXXX/Dropbox/VCP/Screen_data/")
        for f in inputFiles:
            outputFile = outputPath / Path(f.stem + ".jpg")
            im = Image.open(f)
            im = im.convert('RGB')
            im.save(outputFile)
            time.sleep(0.01)
            print('JPG {} is created'.format(f))

        def first_4chars(x):
            return(x[:4])

        jpg_list = []
        for file in os.listdir("/Users/XXXXX/Dropbox/VCP/Screen_data/"):
            if file.endswith(".jpg"):
                jpg_list.append(file)
        #sorted(jpg_list, key = first_4chars)

        try:
            print('Compling into the output document.')
            with open("output_.pdf", "wb") as f:
                f.write(img2pdf.convert([filename for filename in sorted(jpg_list, key = first_4chars)]))
                time.sleep(0.1)
        except:
            print("No stock fit our requirement!")
            pdf = FPDF()
            pdf.add_page()
            pdf.output("output_.pdf")
        time.sleep(2.5)

        f.close()

        total = adv + decl
        stock_name.sort()

        print("--------------------------------------------------------------------------------")
        a1 = (("Advance:Decline = " + str(adv) +" : " + str(decl)))
        a2 = (("Advance:Decline(Weekly) = " + str(adv_w) +" : " + str(decl_w)))
        a3 = ("Stock above its 20-DMA: " + str(round(c_20/total*100,3)))
        a4 = ("Stock above its 50-DMA: " + str(round(c_50/total*100,3)))
        a5 = ("Stock that its 50-DMA > 200-DMA: " + str(round(s_50_200/total*100,3)))
        a6 = ("Stock that its 50 > 150 > 200-DMA: " + str(round(s_50_150_200/total*100,3)))
        a7 = ("Stock that its 200-DMA is rising: " + str(round(s_200_200_20/total*100,3)))
        a8 = ("Index in strength: " + str(index_list))
        a9 = ("Number of Stock that fit condition: " + str(stocks_fit_condition))

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Times", size = 15)
        pdf.cell(200, 10, txt = a1,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a2,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a3,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a4,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a5,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a6,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a7,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a8,
                 ln = 1, align = 'L')
        pdf.cell(200, 10, txt = a9,
                 ln = 1, align = 'L')
        pdf.multi_cell(200,10, str(stock_name), border = 0,
                        align='L', fill= False)

        pdf.output("front.pdf")

        import sys
        if not sys.warnoptions:
            import warnings
            warnings.simplefilter("ignore")

        f_file = open("front.pdf", 'rb')
        o_file = open("output_.pdf", 'rb')

        merger = PdfFileMerger()
        merger.append(PdfFileReader(f_file,strict=False))
        merger.append(PdfFileReader(o_file,strict=False))
        merger.write("output.pdf")

        f_file.close()
        o_file.close()

        try:
            os.rename("output.pdf","{}.pdf".format(end_date))
        except WindowsError:
            os.remove("{}.pdf".format(end_date))
            os.rename('output.pdf',"{}.pdf".format(end_date))

        print("Completed!")

        time.sleep(20)

        test = os.listdir(outputPath)
        for item in test:
            if item.endswith(".jpg"):
                os.remove(os.path.join(outputPath, item))

        time.sleep(20)

        #duration = 1000  # milliseconds
        #freq = 440  # Hz
        #winsound.Beep(freq, duration)

        time.sleep(2.5)

        outputPath = Path("/Users/XXXXX/Dropbox/VCP/Screen_data/")
        test = os.listdir(outputPath)
        for item in test:
            if item.endswith(".png"):
                os.remove(os.path.join(outputPath, item))

        os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/')
        df = pd.DataFrame(columns=['Date','Breadth'])

        listOfFiles = os.listdir('.')
        try:
            listOfFiles.remove('front.pdf')
        except Exception as e:
            print(e)
        try:
            listOfFiles.remove('output_.pdf')
        except Exception as e:
            print(e)
        try:
            listOfFiles.remove('breadth.pdf')
        except Exception as e:
            print(e)
        try:
            listOfFiles.remove('FFTY.pdf')
        except Exception as e:
            print(e)
        try:
            listOfFiles.remove('SPY.pdf')
        except Exception as e:
            print(e)
        try:
            listOfFiles.remove('DIA.pdf')
        except Exception as e:
            print(e)
        try:
            listOfFiles.remove('VOO.pdf')
        except Exception as e:
            print(e)
        try:
            listOfFiles.remove('IWM.pdf')
        except Exception as e:
            print(e)

        pattern = "*.pdf"
        date_list = []
        for entry in listOfFiles:
            if fnmatch.fnmatch(entry, pattern):
                    file_size = os.stat(entry).st_size
                    file_size_log = math.log(file_size)
                    entry = entry[:-4]
                    entry = datetime.strptime(entry, '%Y-%m-%d')
                    real_date = entry - timedelta(days=1)
                    date_list.append(real_date)
                    df = df.append({'Date' : real_date , 'Breadth' : file_size_log} , ignore_index=True)

        df = df.set_index('Date')

        start_date = date_list[1]
        end_date = date_list[-1]
        stock_list = ["FFTY","SPY","DIA","VOO","IWM"]

        for stock in stock_list:
            df_stock = pdr.get_data_yahoo(stock, start=start_date, end=end_date)
            stock_hist = df_stock["Close"]
            df = df.join(stock_hist)
            df.rename(columns={"Close": stock},inplace = True)

        # create figure and axis objects with subplots()
        fig,ax1 = plt.subplots(figsize=(15,8))
        # make a plot
        ax1.plot(df.index, df["Breadth"], color="red")
        # set x-axis label
        ax1.set_xlabel("time",fontsize=14)
        # set y-axis label
        ax1.set_ylabel("Breadth",color="red",fontsize=14)

        # twin object for two different y-axis on the sample plot
        ax2=ax1.twinx()
        # make a plot with different y-axis using second axis object
        ax2.plot(df.index, df["FFTY"],color="blue")
        ax2.set_ylabel("FFTY",color="blue",fontsize=14)
        plt.savefig("FFTY.pdf")

        # create figure and axis objects with subplots()
        fig,ax1 = plt.subplots(figsize=(15,8))
        # make a plot
        ax1.plot(df.index, df["Breadth"], color="red")
        # set x-axis label
        ax1.set_xlabel("time",fontsize=14)
        # set y-axis label
        ax1.set_ylabel("Breadth",color="red",fontsize=14)
        # twin object for two different y-axis on the sample plot
        ax3=ax1.twinx()
        # make a plot with different y-axis using second axis object
        ax3.plot(df.index, df["SPY"],color="green")
        ax3.set_ylabel("SPY",color="green",fontsize=14)
        plt.savefig("SPY.pdf")

        # create figure and axis objects with subplots()
        fig,ax1 = plt.subplots(figsize=(15,8))
        # make a plot
        ax1.plot(df.index, df["Breadth"], color="red")
        # set x-axis label
        ax1.set_xlabel("time",fontsize=14)
        # set y-axis label
        ax1.set_ylabel("Breadth",color="red",fontsize=14)
        # twin object for two different y-axis on the sample plot
        ax3=ax1.twinx()
        # make a plot with different y-axis using second axis object
        ax3.plot(df.index, df["DIA"],color="black")
        ax3.set_ylabel("DIA",color="black",fontsize=14)
        plt.savefig("DIA.pdf")

        # create figure and axis objects with subplots()
        fig,ax1 = plt.subplots(figsize=(15,8))
        # make a plot
        ax1.plot(df.index, df["Breadth"], color="red")
        # set x-axis label
        ax1.set_xlabel("time",fontsize=14)
        # set y-axis label
        ax1.set_ylabel("Breadth",color="red",fontsize=14)
        # twin object for two different y-axis on the sample plot
        ax3=ax1.twinx()
        # make a plot with different y-axis using second axis object
        ax3.plot(df.index, df["VOO"],color="orange")
        ax3.set_ylabel("VOO",color="orange",fontsize=14)
        plt.savefig("VOO.pdf")

        # create figure and axis objects with subplots()
        fig,ax1 = plt.subplots(figsize=(15,8))
        # make a plot
        ax1.plot(df.index, df["Breadth"], color="red")
        # set x-axis label
        ax1.set_xlabel("time",fontsize=14)
        # set y-axis label
        ax1.set_ylabel("Breadth",color="red",fontsize=14)
        # twin object for two different y-axis on the sample plot
        ax3=ax1.twinx()
        # make a plot with different y-axis using second axis object
        ax3.plot(df.index, df["IWM"],color="purple")
        ax3.set_ylabel("IWM",color="purple",fontsize=14)
        plt.savefig("IWM.pdf")

        import PyPDF2

        # Open the files that have to be merged one by one
        pdf1File = open('FFTY.pdf', 'rb')
        pdf2File = open('SPY.pdf', 'rb')
        pdf3File = open('DIA.pdf', 'rb')
        pdf4File = open('VOO.pdf', 'rb')
        pdf5File = open('IWM.pdf', 'rb')

        # Read the files that you have opened
        pdf1Reader = PyPDF2.PdfFileReader(pdf1File)
        pdf2Reader = PyPDF2.PdfFileReader(pdf2File)
        pdf3Reader = PyPDF2.PdfFileReader(pdf3File)
        pdf4Reader = PyPDF2.PdfFileReader(pdf4File)
        pdf5Reader = PyPDF2.PdfFileReader(pdf5File)

        # Create a new PdfFileWriter object which represents a blank PDF document
        pdfWriter = PyPDF2.PdfFileWriter()

        # Loop through all the pagenumbers for the first document
        for pageNum in range(pdf1Reader.numPages):
            pageObj = pdf1Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        # Loop through all the pagenumbers for the second document
        for pageNum in range(pdf2Reader.numPages):
            pageObj = pdf2Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        # Loop through all the pagenumbers for the third document
        for pageNum in range(pdf3Reader.numPages):
            pageObj = pdf3Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        # Loop through all the pagenumbers for the fourth document
        for pageNum in range(pdf4Reader.numPages):
            pageObj = pdf4Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        # Loop through all the pagenumbers for the fourth document
        for pageNum in range(pdf5Reader.numPages):
            pageObj = pdf5Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        # Now that you have copied all the pages in both the documents, write them into the a new document
        pdfOutputFile = open('breadth.pdf', 'wb')
        pdfWriter.write(pdfOutputFile)

        # Close all the files - Created as well as opened
        pdfOutputFile.close()
        pdf1File.close()
        pdf2File.close()
        pdf3File.close()
        pdf4File.close()
        pdf5File.close()

        os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/python')

        try:
            os.remove("ScreenOutput.xlsx")
            os.remove("stocks.csv")
            time.sleep(3)
        except Exception as e:
            print(e)
        working == False
        os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/python')
        #os.system("python VCP_Screener.py")
        print("Restarting...")
        spawn_program_and_die(['python', 'vcp_screener.py'])
    except (RuntimeError, TypeError, NameError, PermissionError,ValueError,OSError,json.decoder.JSONDecodeError, JSONDecodeError):
        print("Runtime Error: Will restart the Program after 2 min.")
        time.sleep(20)
        os.chdir('/Users/XXXXX/Dropbox/VCP/Screen_data/python')
        spawn_program_and_die(['python', 'vcp_screener.py'])
