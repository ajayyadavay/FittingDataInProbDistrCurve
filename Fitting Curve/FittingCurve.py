#  pip install xlrd

import xlrd
import math
import ProbabilityDistribution as pd
import PyMathAy as pmath
import matplotlib.pyplot as plt 
import tkinter as tk 
from tkinter import filedialog
import os
from tkinter import messagebox

# import xlwt
import xlsxwriter


root = tk.Tk()
# root.configure(background = 'teal')
root.resizable(0,0) # disalbe resizing of window and maximization button
root.title("Fitting Data in Prob. Distr. Curve")
 
# path of file
global loc

def ImportFromExcel():
    global Q_data, Sorted_Q_data, Monthly_mean, Monthyly_stddev, logMonthly_mean, logMonthyly_stddev

    Q_data = []
    Sorted_Q_data = []
    Monthly_mean = []
    Monthyly_stddev = []

    Monthly_M100 = []
    Monthly_M110 = []
    Monthly_M120 = []
    Monthly_M130 = []

    L1 = []
    L2 = []
    L3 = []
    L4 = []

    T2 = []
    T3 = []
    T4 = []

    global K
    K = []

    logMonthly_mean = []
    logMonthyly_stddev = []

    global number_of_years
    number_of_years = int(entryNumber.get())
    
    loc = filedialog.askopenfilename(filetypes = (("Excel", "*.xlsx"), ("All files", "*")))
    wb= xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for i in range(12): 
        Q_data.append([])
        Sorted_Q_data.append([])
        for j in range(number_of_years):
            #Q_data[i].append(sheet.cell_value(i+1,j + 1))
            #Sorted_Q_data[i].append(sheet.cell_value(i+1,j + 1))
            Q_data[i].append(sheet.cell_value(i+2,j + 1))
            Sorted_Q_data[i].append(sheet.cell_value(i+2,j + 1))
    
    # calcuating mean of each month
    for i in range(12):
        sum = 0
        sum1 = 0
        for j in range(number_of_years):
            sum = sum + Q_data[i][j]
            sum1 = sum1 + math.log(Q_data[i][j])
        Monthly_mean.append(sum/number_of_years)
        logMonthly_mean.append(sum1/number_of_years)

    # calcuating standard deviation of each month
    for i in range(12):
        sum = 0
        sum1 = 0
        for j in range(number_of_years):
            sum = sum + (Q_data[i][j] - Monthly_mean[i])**2
            sum1 = sum1 + (math.log(Q_data[i][j]) -logMonthly_mean[i])**2
        stddev = math.sqrt(sum/(number_of_years-1))
        Monthyly_stddev.append(stddev)
        stddev = math.sqrt(sum1/(number_of_years-1))
        logMonthyly_stddev.append(stddev)

    # sorting discharge data
    srt = pmath.Sorting()
    for i in range(12):
        #Sorted_Q_data[i].sort()
        srt.SortInAscendingOrder(Sorted_Q_data[i])
    
    # Probability Weighted Moments Equations (PWMs)
    for i in range(12):
        sum100 = 0
        sum110 = 0
        sum120 = 0
        sum130 = 0
        for j in range(number_of_years):
            sum100 = sum100 + Sorted_Q_data[i][j]
            sum110 = sum110 + (j + 1 - 1)/(number_of_years - 1) * Sorted_Q_data[i][j]
            sum120 = sum120 + (j + 1 - 1) * (j + 1 - 2)/((number_of_years-1)*(number_of_years-2)) * Sorted_Q_data[i][j]
            sum130 = sum130 + (j + 1 - 1) * (j + 1 - 2) *(j + 1 - 3)/((number_of_years-1)*(number_of_years-2)*(number_of_years-3)) * Sorted_Q_data[i][j]
        Monthly_M100.append(sum100)
        Monthly_M110.append(sum110)
        Monthly_M120.append(sum120)
        Monthly_M130.append(sum130)

    # Calcuating L-Moments, useful ratios and shape parameter of GEV
    for i in range(12):
        L1.append(Monthly_M100[i])
        L2.append(2*Monthly_M110[i] - Monthly_M100[i])
        L3.append(6*Monthly_M120[i] - 6*Monthly_M110[i] + Monthly_M100[i])
        L4.append(20*Monthly_M130[i]-30*Monthly_M120[i]+12*Monthly_M110[i]-Monthly_M100[i])

        T2.append(L2[i] / L1[i]) # L-CV; Like Coeffiecient of variance
        T3.append(L3[i] / L2[i]) # L-Skewness; Shape parameter
        T4.append(L4[i] / L2[i]) # L-Kurtosis; measure of peakedness

        c = 2 / (3 + T3[i]) - math.log(2)/math.log(3)
        K.append(7.8590 * c + 2.9884 * c * c) # shape parameter of each month of GEV

    # ROW = 0
    Lbltitle = tk.Label(root, text = "CURVE FITTING FOR " + str(number_of_years) + " YEARS DATA", fg="white", bg="teal")
    Lbltitle.grid(row=0,column=0,sticky="e,w",columnspan=3)

    # ROW = 2
    LblImport = tk.Label(root, text = loc, fg="blue")
    LblImport.grid(row=2,column=1,sticky="e,w",columnspan=2)

# end ImportFromExcel------------------------------------------------
# Consolas
# ROW = 0
Lbltitle = tk.Label(root, text = "CURVE FITTING FOR _ YEARS DATA", fg="white",bg="teal")
Lbltitle.grid(row=0,column=0,sticky="e,w",columnspan=3)

# ROW = 1
LblEntry = tk.Label(root, text = "Enter number of Year's to be used", fg="darkviolet")
LblEntry.grid(row=1,column=0,sticky="e,w",columnspan=2)

entryNumber = tk.Entry(root, fg = "darkviolet")
entryNumber.grid(row=1,column=2,sticky="e,w")

# ROW = 2
BtnImport = tk.Button(root, text = "Import", fg="green",command=ImportFromExcel)
BtnImport.grid(row=2,column=0,sticky="e,w")

Month = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# plotting Normal distribution curve for all months
def PlotNormalPDFCurve():
    plt.clf()
    x_axis = []
    y_axis = []
    for i in range(0,12): 
        x_axis.append([])
        y_axis.append([])
        for j in range(number_of_years):
            x_axis[i].append(Sorted_Q_data[i][j])

            Zval = pd.NormalDistr(x_axis[i][j],Monthly_mean[i],Monthyly_stddev[i])
            y_axis[i].append(Zval.PDF())

        plt.title("Normal Probability Density Function")
        plt.plot(x_axis[i], y_axis[i], label = str(Month[i])) 

    plt.legend()
    plt.show()


def PlotNormalCDFCurve():
    plt.clf()
    x_axis_cdf = []
    y_axis_cdf = []
    for i in range(0,12): 
        x_axis_cdf.append([])
        y_axis_cdf.append([])
        for j in range(number_of_years):
            x_axis_cdf[i].append(Sorted_Q_data[i][j])

            Zval = pd.NormalDistr(x_axis_cdf[i][j],Monthly_mean[i],Monthyly_stddev[i])
            y_axis_cdf[i].append(Zval.CDF()) 

        plt.title("Normal Cummulative Distribution Function")
        plt.plot(x_axis_cdf[i], y_axis_cdf[i], label = str(Month[i]))
    plt.legend()
    plt.show()

# ROW = 3
LblNormal = tk.Label(root, text = "Normal Distribution", fg="blue")
LblNormal.grid(row=3,column=0,sticky="e,w")

BtnNormPDF = tk.Button(root, text = "Plot Normal PDF", command = PlotNormalPDFCurve,fg="red")
BtnNormPDF.grid(row=3,column=1,sticky="e,w")

BtnNormCDF = tk.Button(root, text = "Plot Normal CDF", command = PlotNormalCDFCurve,fg="red")
BtnNormCDF.grid(row=3,column=2,sticky="e,w")

def ALLMONTHCDF(MonthIndex):
    plt.clf() # clear the previous graph (if any)
    x_axis_cdf1 = []
    y_axis_cdf1 = []

    y_original_cdf = []

    logy_axis_cdf = []
    Expy_axis_cdf = []
    GEVy_axis_cdf = []

    diff_Normal = []
    diff_lognorm = []
    diff_Exp = []
    diff_GEV = []

    KS_Normal = []
    KS_logNorm = []
    KS_Exp = []
    KS_GEV = []

    for i in range(0,MonthIndex+1): 
        x_axis_cdf1.append([])
        y_axis_cdf1.append([])

        y_original_cdf.append([])

        logy_axis_cdf.append([])
        Expy_axis_cdf.append([])
        GEVy_axis_cdf.append([])

        diff_Normal.append([])
        diff_lognorm.append([])
        diff_Exp.append([])
        diff_GEV.append([])
        
        if i != MonthIndex: # to ensure that calculation is done for only one required month
            continue

        for j in range(number_of_years):
            x_axis_cdf1[i].append(Sorted_Q_data[i][j])

            # Original CDF
            y_original_cdf[i].append((j+1)/number_of_years)

            # Normal
            Zval = pd.NormalDistr(x_axis_cdf1[i][j],Monthly_mean[i],Monthyly_stddev[i],x_axis_cdf1[i][0])
            y_axis_cdf1[i].append(Zval.CDF()) 

            # Log Normal
            lognorm = pd.LogNormalDistr(x_axis_cdf1[i][j],logMonthly_mean[i],logMonthyly_stddev[i],x_axis_cdf1[i][0])
            logy_axis_cdf[i].append(lognorm.CDF())

            # Exponential
            Expo = pd.Exponential(x_axis_cdf1[i][j],Monthly_mean[i],x_axis_cdf1[i][0])
            Expy_axis_cdf[i].append(Expo.CDF())

            # GEV 
            Gevplot = pd.GeneralizedExtremeValue(x_axis_cdf1[i][j],Monthly_mean[i],Monthyly_stddev[i],K[i],x_axis_cdf1[i][0])
            GEVy_axis_cdf[i].append(Gevplot.CDF())

            # Calculating Kolmogorov–Smirnov (K-S) distance 
            diff_Normal[i].append(abs(y_original_cdf[i][j] - y_axis_cdf1[i][j]))
            diff_lognorm[i].append(abs(y_original_cdf[i][j] - logy_axis_cdf[i][j]))
            diff_Exp[i].append(abs(y_original_cdf[i][j] - Expy_axis_cdf[i][j]))
            diff_GEV[i].append(abs(y_original_cdf[i][j] - GEVy_axis_cdf[i][j]))

        if i == MonthIndex: # to ensure that only one required month is plotted in graph
            KS_Normal.append(max(diff_Normal[i]))
            KS_logNorm.append(max(diff_lognorm[i]))
            KS_Exp.append(max(diff_Exp[i]))
            KS_GEV.append(max(diff_GEV[i]))
            D_cr_point05 = 1.36/math.sqrt(number_of_years)
            plt.title("Cummulative Distribution Function for " + Month[i])
            plt.plot(x_axis_cdf1[i], y_original_cdf[i], 'bo', label = "Original;D_cr0.05= " + str(round(D_cr_point05,4))) # 'bo' gives blue circle plot
            #plt.plot(x_axis_cdf1[i], y_original_cdf[i], label = "Original;D_cr0.05= " + str(round(D_cr_point05,4)))
            plt.plot(x_axis_cdf1[i], y_axis_cdf1[i], label = "Normal; K-s= " + str(round(KS_Normal[0],4)))
            plt.plot(x_axis_cdf1[i], logy_axis_cdf[i], label = "LogNormal;K-S= "+ str(round(KS_logNorm[0],4)))
            plt.plot(x_axis_cdf1[i], Expy_axis_cdf[i], label = "Exponential;K-S= " + str(round(KS_Exp[0],4)))
            # plt.plot(x_axis_cdf1[i], GEVy_axis_cdf[i], label = "GEV with k = " + str(round(K[i],4)))
            plt.plot(x_axis_cdf1[i], GEVy_axis_cdf[i], label = "GEV;K-S= " + str(round(KS_GEV[0],4)))
    plt.legend()
    plt.show()

def ExportToExcel():
    #write_wb = Workbook(write_only = True)
    #write_ws = write_wb.create_sheet() 

    wb = xlsxwriter.Workbook('AJAY_074MSWRE002_k_S_Distance.xlsx')
    ws = wb.add_worksheet('ProbabilityDistr')

    # Formating above critical cells
    CriticalcellFormat = wb.add_format()
    # CriticalcellFormat.set_bold()
    CriticalcellFormat.set_font_color('red')
    CriticalcellFormat.set_border()
    CriticalcellFormat.set_border_color('red')

    # Formating Non-critical cells
    Noncriticalformat = wb.add_format()
    Noncriticalformat.set_font_color('green')
    Noncriticalformat.set_border()
    Noncriticalformat.set_border_color('green')

    # Fomating headings
    HeadingFormat = wb.add_format()
    HeadingFormat.set_bold()
    HeadingFormat.set_border()
    HeadingFormat.set_border_color('gray')
    HeadingFormat.set_bg_color('silver')

    x_axis_cdf1 = []
    y_axis_cdf1 = []

    y_original_cdf = []

    logy_axis_cdf = []
    Expy_axis_cdf = []
    GEVy_axis_cdf = []

    diff_Normal = []
    diff_lognorm = []
    diff_Exp = []
    diff_GEV = []

    KS_Normal = []
    KS_logNorm = []
    KS_Exp = []
    KS_GEV = []

    startrow = 3
    
    #write_ws.append(["Month", "Critical_0.05" , "KS_Normal","KS_logNorm","KS_Exp","KS_GEV"])
    ws.write(startrow-1,0, "Month",HeadingFormat)
    ws.write(startrow-1,1, "Critical_0.05",HeadingFormat)
    ws.write(startrow-1,2, "KS_Normal",HeadingFormat)
    ws.write(startrow-1,3, "KS_logNorm",HeadingFormat)
    ws.write(startrow-1,4, "KS_Exp",HeadingFormat)
    ws.write(startrow-1,5, "KS_GEV",HeadingFormat)

    for i in range(0,12): 
        x_axis_cdf1.append([])
        y_axis_cdf1.append([])

        y_original_cdf.append([])

        logy_axis_cdf.append([])
        Expy_axis_cdf.append([])
        GEVy_axis_cdf.append([])

        diff_Normal.append([])
        diff_lognorm.append([])
        diff_Exp.append([])
        diff_GEV.append([])

        for j in range(number_of_years):
            x_axis_cdf1[i].append(Sorted_Q_data[i][j])
            # Original CDF
            y_original_cdf[i].append((j+1)/number_of_years)
            # Normal
            Zval = pd.NormalDistr(x_axis_cdf1[i][j],Monthly_mean[i],Monthyly_stddev[i],x_axis_cdf1[i][0])
            y_axis_cdf1[i].append(Zval.CDF()) 
            # Log Normal
            lognorm = pd.LogNormalDistr(x_axis_cdf1[i][j],logMonthly_mean[i],logMonthyly_stddev[i],x_axis_cdf1[i][0])
            logy_axis_cdf[i].append(lognorm.CDF())
            # Exponential
            Expo = pd.Exponential(x_axis_cdf1[i][j],Monthly_mean[i],x_axis_cdf1[i][0])
            Expy_axis_cdf[i].append(Expo.CDF())
            # GEV 
            Gevplot = pd.GeneralizedExtremeValue(x_axis_cdf1[i][j],Monthly_mean[i],Monthyly_stddev[i],K[i],x_axis_cdf1[i][0])
            GEVy_axis_cdf[i].append(Gevplot.CDF())
            # Differencing for Kolmogorov–Smirnov (K-S) distance 
            diff_Normal[i].append(abs(y_original_cdf[i][j] - y_axis_cdf1[i][j]))
            diff_lognorm[i].append(abs(y_original_cdf[i][j] - logy_axis_cdf[i][j]))
            diff_Exp[i].append(abs(y_original_cdf[i][j] - Expy_axis_cdf[i][j]))
            diff_GEV[i].append(abs(y_original_cdf[i][j] - GEVy_axis_cdf[i][j]))
        
        # Taking maximum value of difference for Kolmogorov–Smirnov (K-S) distance
        D_cr_point05 = 1.36/math.sqrt(number_of_years)
        KS_Normal.append(max(diff_Normal[i]))
        KS_logNorm.append(max(diff_lognorm[i]))
        KS_Exp.append(max(diff_Exp[i]))
        KS_GEV.append(max(diff_GEV[i]))

        # in xlwt: write(row number, col, label='', style=<xlwt.Style.XFStyle object>)
        ws.write(i+startrow,0, Month[i],HeadingFormat)
        ws.write(i+startrow,1, D_cr_point05,Noncriticalformat)

        if D_cr_point05 < KS_Normal[i]:
            ws.write(i+startrow,2, KS_Normal[i],CriticalcellFormat)
        else:
            ws.write(i+startrow,2, KS_Normal[i],Noncriticalformat)
        
        if D_cr_point05 < KS_logNorm[i]:
            ws.write(i+startrow,3, KS_logNorm[i],CriticalcellFormat)
        else:
            ws.write(i+startrow,3, KS_logNorm[i],Noncriticalformat)
        
        if D_cr_point05 < KS_Exp[i]:
            ws.write(i+startrow,4, KS_Exp[i],CriticalcellFormat)
        else:
            ws.write(i+startrow,4, KS_Exp[i],Noncriticalformat)
     
        if D_cr_point05 < KS_GEV[i]:
            ws.write(i+startrow,5, KS_GEV[i],CriticalcellFormat)
        else:
            ws.write(i+startrow,5, KS_GEV[i],Noncriticalformat)
     
        '''ws.write(i+1,0, Month[i])
        ws.write(i+1,1, D_cr_point05)
        ws.write(i+1,2, KS_Normal[i])
        ws.write(i+1,3,KS_logNorm[i])
        ws.write(i+1,4,KS_Exp[i])
        ws.write(i+1,5,KS_GEV[i])'''

    wb.close()
        

# ROW = 4
BtnJANCDF = tk.Button(root, text = "JAN CDF", command = lambda: ALLMONTHCDF(0),fg="green")
BtnJANCDF.grid(row=4,column=0,sticky="e,w")

BtnFEBCDF = tk.Button(root, text = "FEB CDF", command = lambda: ALLMONTHCDF(1),fg="green")
BtnFEBCDF.grid(row=4,column=1,sticky="e,w")

BtnMARCDF = tk.Button(root, text = "MAR CDF", command = lambda: ALLMONTHCDF(2),fg="green")
BtnMARCDF.grid(row=4,column=2,sticky="e,w")

# ROW = 5
BtnAPRCDF = tk.Button(root, text = "APR CDF", command = lambda: ALLMONTHCDF(3),fg="green")
BtnAPRCDF.grid(row=5,column=0,sticky="e,w")

BtnMAYCDF = tk.Button(root, text = "MAY CDF", command = lambda: ALLMONTHCDF(4),fg="green")
BtnMAYCDF.grid(row=5,column=1,sticky="e,w")

BtnJUNCDF = tk.Button(root, text = "JUN CDF", command = lambda: ALLMONTHCDF(5),fg="green")
BtnJUNCDF.grid(row=5,column=2,sticky="e,w")

# ROW = 6
BtnJULCDF = tk.Button(root, text = "JUL CDF", command = lambda: ALLMONTHCDF(6),fg="green")
BtnJULCDF.grid(row=6,column=0,sticky="e,w")

BtnAUGCDF = tk.Button(root, text = "AUG CDF", command = lambda: ALLMONTHCDF(7),fg="green")
BtnAUGCDF.grid(row=6,column=1,sticky="e,w")

BtnSEPCDF = tk.Button(root, text = "SEP CDF", command = lambda: ALLMONTHCDF(8),fg="green")
BtnSEPCDF.grid(row=6,column=2,sticky="e,w")

# ROW = 7
BtnOCTCDF = tk.Button(root, text = "OCT CDF", command = lambda: ALLMONTHCDF(9),fg="green")
BtnOCTCDF.grid(row=7,column=0,sticky="e,w")

BtnNOVCDF = tk.Button(root, text = "NOV CDF", command = lambda: ALLMONTHCDF(10),fg="green")
BtnNOVCDF.grid(row=7,column=1,sticky="e,w")

BtnDECCDF = tk.Button(root, text = "DEC CDF", command = lambda: ALLMONTHCDF(11),fg="green")
BtnDECCDF.grid(row=7,column=2,sticky="e,w")

def About():
    messagebox.showinfo("About", "Fitting Data in Probability Distribution curve\n\nPython version 3.7.2 \n\nAjay Yadav \n074 MSWRE 002")

def HowTo():
    messagebox.showinfo("How To", "Fitting Data in Probability Distribution curve \n\nStep 1: Input number of years data to be used \nStep 2: Click on Import \nStep 3: Navigate to the excel file \nStep 4: Click on any month CDF e.g. JAN CDF")

# ROW = 8
BtnExportxl = tk.Button(root, text = "Export K-S To Excel", command = ExportToExcel,fg="blue")
BtnExportxl.grid(row=8,column=0,sticky="e,w",columnspan = 3)

# ROW = 9
BtnAbout = tk.Button(root, text = "About", command = About,fg="dark red")
BtnAbout.grid(row=9,column=0,sticky="e,w")

BtnHowTo = tk.Button(root, text = "How To", command = HowTo,fg="dark red")
BtnHowTo.grid(row=9,column=1,sticky="e,w")

BtnExit = tk.Button(root, text = "Exit", command = quit,fg="dark red")
BtnExit.grid(row=9,column=2,sticky="e,w")

root.mainloop()
