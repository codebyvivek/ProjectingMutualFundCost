import openpyxl
from tkinter import *

import Excel_Calculation
import FocusToGrowth

wb = openpyxl.load_workbook("C:\\Users\\vivek\\Desktop\\MR. MAHESH BASUDKAR PROJECTION REPORT.xlsx")

sheets = wb.sheetnames

s1 = wb['Need Base Analysis']

needBaseAnalyis =False
def child_1_calculation():
    global needBaseAnalyis
    print(rate.get())
    label_child_1 = Label(Frame, text=Excel_Calculation.needBaseSheetValue(rate.get())).grid(row=4, column=0)
    needBaseAnalyis =True


def ProjectionProject(Input_Per_month_Investment,Input_Total_Investment,Input_Total_Corpse_Amount):
    global needBaseAnalyis
    print(rate1.get())
    Input = Input_Per_month_Investment.get()
    Input2 =Input_Total_Investment.get()
    Input3 =Input_Total_Corpse_Amount()
    if needBaseAnalyis==True:
       label_ProjectionReport = Label(Frame, text=FocusToGrowth.methodToFillProjectionReport(rate1.get(),Input,Input2,Input3)).grid(row=8, column=0)


root = Tk(screenName="Focus To Growth", baseName="Focus To Growth")
Lframe = LabelFrame(root, text="Need base Analysis for : " + s1['D5'].value, padx=5, pady=5)
Lframe.grid(row=1, column=1)


Frame = LabelFrame(Lframe, text="Click Button to Complete Need Base Analysis", padx=10, pady=10)
Frame.grid(row=2, column=0)
values_of_interest =[1,2,3,4,5,6,7,8,9,10]
rate =IntVar()
rate.set(values_of_interest[9])
label_child_1 = Label(Frame, text="Please Select the rate to calculate compound interest").grid(row=2, column=1)
Interest =OptionMenu(Frame, rate, *values_of_interest).grid(row=2, column=2)
button_child_1 = Button(Frame, text="Click Button to Complete Need Base Analysis", command=lambda: child_1_calculation)
button_child_1.grid(row=3, column=1)

Frame = LabelFrame(Lframe, text="Click Button to Complete Projection Report", padx=10, pady=10)
Frame.grid(row=5, column=0)
values_of_interest1 =[6,7,8,9,10,11,12]
rate1 =IntVar()
rate1.set(values_of_interest1[6])
label_ProjectionReport = Label(Frame,
                               text="Please Select the rate to calculate compound interest").grid(row=5, column=1)
Interest =OptionMenu(Frame, rate1, *values_of_interest1).grid(row=5, column=2)
inp1 = StringVar()
inp2 = StringVar()
inp3 =StringVar()
Frame7 = LabelFrame(Lframe, text="Project Report Input", padx=10, pady=10)
Frame7.grid(row=7, column=0)
label_Input_1 = Label(Frame7, text="Input Per month investment").grid(row=7, column=0)
Input_Per_month_Investment = Entry(Frame7, textvariable=inp1, width=15).grid(row=7, column=1)
label_Input_2 = Label(Frame7, text="Input Total Investment").grid(row=8, column=0)
Input_Total_Investment = Entry(Frame7, textvariable=inp2, width=15).grid(row=8, column=1)
label_Input_3 = Label(Frame7, text="Input Total Corpse Amount").grid(row=9, column=0)
Input_Total_Corpse_Amount = Entry(Frame7, textvariable=inp3, width=15).grid(row=9, column=1)
button_ProjectionReport = Button(Frame, text="Click Button to Complete Need Base Analysis")
button_ProjectionReport.grid(row=6, column=1)
root.mainloop()