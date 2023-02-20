import openpyxl
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import numpy as np

#data2022 = openpyxl.open("2022.xlsx", read_only=True)
#sheet = data2022.active
import pandas as pd
import matplotlib.pyplot as plt
file_name = "2017.xlsx"
df_2022 = pd.read_excel(file_name)

def create_graph_country():
    #cal ctrb by countries
    unique_country = pd.unique(df_2022['country']) #looking for unique values
    ctrb_countries_1 = [] #create an empty array, count the ctrb of unique values and write to the array
    for index in range(len(unique_country)):
        ctrb_countries = df_2022[(df_2022['country'] == unique_country[index])]['ctrb_day'].sum()
        ctrb_countries_1.append(ctrb_countries)

    # DataFrame creation
    df_country = {'Country': unique_country, 'Ctrb': ctrb_countries_1}
    dframe_country = pd.DataFrame(df_country)

    #create an empty sheet, write the dataframe into it
    FilePath = "contribution_data" + file_name
    writer = pd.ExcelWriter(FilePath, engine='xlsxwriter')
    dframe_country.to_excel(writer, sheet_name ='ctrb_country')
    writer.save()

    #graph  creating
    plt.barh(unique_country, width=ctrb_countries_1)
    plt.title("Performance: countries")
    plt.xlabel("Performance, %")
    plt.savefig('ctrb_countries.png')
    plt.show()

    #creating boxplot
    boxplot = dframe_country.boxplot(column=['Ctrb'])
    boxplot.plot()
    plt.savefig('ctrb_countries_boxplot.png')
    plt.show()

    #adding graphical objects to the file
    wb = openpyxl.load_workbook(FilePath)
    active = wb['ctrb_country']
    active.add_image(Image('ctrb_countries.png'), 'F1')
    active.add_image(Image('ctrb_countries_boxplot.png'), 'P1')
    wb.save(FilePath)

    print('Country info is written successfully to Excel File.')

    #remove unnecessary images
    os.remove("ctrb_countries.png")
    os.remove("ctrb_countries_boxplot.png")

def create_graph_currency():
    #cal ctrb by currency
    unique_currency = pd.unique(df_2022['price_ccy']) #looking for unique values
    ctrb_currency_1 = [] #create an empty array, count the ctrb of unique values and write to the array
    for index in range(len(unique_currency)):
        ctrb_currency = df_2022[df_2022['price_ccy'] == unique_currency[index]]['ctrb_day'].sum()
        ctrb_currency_1.append(ctrb_currency)

    # DataFrame creation
    df_currency = {'Currency': unique_currency, 'Ctrb': ctrb_currency_1}
    dframe_currency = pd.DataFrame(df_currency)

    #graph  creating
    plt.barh(unique_currency, width=ctrb_currency_1)
    plt.title("Performance: currency")
    plt.xlabel("Performance, %")
    plt.savefig('ctrb_currency.png')
    plt.show()

    #creating boxplot
    boxplot = dframe_currency.boxplot(column=['Ctrb'])
    boxplot.plot()
    plt.savefig('ctrb_currency_boxplot.png')
    plt.show()

    #create an empty sheet, write the dataframe into it
    FilePath = "contribution_data" + file_name
    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    dframe_currency.to_excel(writer, sheet_name='ctrb_currency')
    writer.save()

    #adding graphical objects to the file
    wb = openpyxl.load_workbook(FilePath)
    # wb.create_sheet('ctrb_currency')
    active1 = wb['ctrb_currency']
    active1.add_image(Image('ctrb_currency.png'), 'F1')
    active1.add_image(Image('ctrb_currency_boxplot.png'), 'P1')
    wb.save(FilePath)

    print('Currency info is written successfully to Excel File.')

    #remove unnecessary images
    os.remove("ctrb_currency.png")
    os.remove("ctrb_currency_boxplot.png")

def create_graph_inst_class():
    # cal ctrb by inst_class
    unique_inst_class = pd.unique(df_2022['inst_class'])  #looking for unique values
    ctrb_inst_class_1 = [] #create an empty array, count the ctrb of unique values and write to the array
    for index in range(len(unique_inst_class)):
        ctrb_inst_class = df_2022[(df_2022['inst_class'] == unique_inst_class[index])]['ctrb_day'].sum()
        ctrb_inst_class_1.append(ctrb_inst_class)
    # DataFrame creation
    df_inst_class = {'Instrument Class': unique_inst_class, 'Ctrb': ctrb_inst_class_1}
    dframe_inst_class = pd.DataFrame(df_inst_class)

    #graph  creating
    plt.bar(unique_inst_class, height=ctrb_inst_class_1)
    plt.title("Performance: Instrument Class")
    plt.ylabel("Performance, %")
    plt.savefig('ctrb_inst_class.png')
    plt.show()

    #creating boxplot
    boxplot = dframe_inst_class.boxplot(column=['Ctrb'])
    boxplot.plot()
    plt.savefig('ctrb_inst_class_boxplot.png')
    plt.show()

    #create an empty sheet, write the dataframe into it
    FilePath = "contribution_data" + file_name
    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    dframe_inst_class.to_excel(writer, sheet_name='ctrb_inst_class')
    writer.save()

    #adding graphical objects to the file
    wb = openpyxl.load_workbook(FilePath)
    # wb.create_sheet('ctrb_inst_class')
    active1 = wb['ctrb_inst_class']
    active1.add_image(Image('ctrb_inst_class.png'), 'F1')
    active1.add_image(Image('ctrb_inst_class_boxplot.png'), 'P1')
    wb.save(FilePath)

    print('Inst_class info is written successfully to Excel File.')

    #remove unnecessary images
    os.remove("ctrb_inst_class.png")
    os.remove("ctrb_inst_class_boxplot.png")

def create_graph_sector():
    # cal ctrb by inst_sector
    unique_sector = pd.unique(df_2022['sector'])  #looking for unique values
    ctrb_sector_1 = [] #create an empty array, count the ctrb of unique values and write to the array
    for index in range(len(unique_sector)):
        ctrb_sector = df_2022[(df_2022['sector'] == unique_sector[index])]['ctrb_day'].sum()
        ctrb_sector_1.append(ctrb_sector)
    # DataFrame creation
    df_inst_class = {'Instrument Class': unique_sector, 'Ctrb': ctrb_sector_1}
    dframe_sector = pd.DataFrame(df_inst_class)

    #graph  creating
    plt.barh(unique_sector, width=ctrb_sector_1)
    plt.title("Performance: sector")
    plt.xlabel("Performance, %")
    plt.savefig('ctrb_sector.png')
    plt.show()

    #creating boxplot
    boxplot = dframe_sector.boxplot(column=['Ctrb'])
    boxplot.plot()
    plt.savefig('ctrb_sector_boxplot.png')
    plt.show()

    #create an empty sheet, write the dataframe into it
    FilePath = "contribution_data" + file_name
    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    dframe_sector.to_excel(writer, sheet_name='ctrb_sector')
    writer.save()

    #adding graphical objects to the file
    wb = openpyxl.load_workbook(FilePath)
    # wb.create_sheet('ctrb_')
    active = wb['ctrb_sector']
    active.add_image(Image('ctrb_sector.png'), 'F1')
    active.add_image(Image('ctrb_sector_boxplot.png'), 'P1')
    wb.save(FilePath)

    print('Sector info is written successfully to Excel File.')

    #remove unnecessary images
    os.remove("ctrb_sector.png")
    os.remove("ctrb_sector_boxplot.png")

def create_graph_day_week():
    # cal ctrb by day_week  0-monday, 6 - sunday
    df_2022['day_week'] = pd.to_datetime(df_2022['dt']).dt.dayofweek
    unique_day_week = pd.unique(df_2022['day_week']) #looking for unique values
    unique_day_week.sort()
    ctrb_day_week_1 = [] #create an empty array, count the ctrb of unique values and write to the array
    for index in range(len(unique_day_week)):
        ctrb_day_week = df_2022[(df_2022['day_week'] == unique_day_week[index])]['ctrb_day'].sum()
        ctrb_day_week_1.append(ctrb_day_week)

    # сhange the numbers to the names of the days of the week
    unique_day_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    df_day_week = {'Day of the week': unique_day_week, 'Ctrb': ctrb_day_week_1}
    # DataFrame creation
    dframe_day_week = pd.DataFrame(df_day_week)

    #graph  creating
    plt.barh(unique_day_week, width=ctrb_day_week_1)
    plt.title("Performance: Day of the week")
    plt.xlabel("Performance, %")
    plt.savefig('ctrb_day_week.png')
    plt.show()

    #creating boxplot
    boxplot = dframe_day_week.boxplot(column=['Ctrb'])
    boxplot.plot()
    plt.savefig('ctrb_day_week_boxplot.png')
    plt.show()

    #create an empty sheet, write the dataframe into it
    FilePath = "contribution_data" + file_name
    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    dframe_day_week.to_excel(writer, sheet_name='ctrb_day_week')
    writer.save()

    #adding graphical objects to the file
    wb = openpyxl.load_workbook(FilePath)
    # wb.create_sheet('ctrb_day_week')
    active = wb['ctrb_day_week']
    active.add_image(Image('ctrb_day_week.png'), 'F1')
    active.add_image(Image('ctrb_day_week_boxplot.png'), 'P1')
    wb.save(FilePath)

    print('Day_week info is written successfully to Excel File.')

    #remove unnecessary images
    os.remove("ctrb_day_week.png")
    os.remove("ctrb_day_week_boxplot.png")

def create_graph_week():
    # cal ctrb by week
    df_2022['week'] = pd.to_datetime(df_2022['dt']).dt.isocalendar().week
    unique_week = pd.unique(df_2022['week']) #looking for unique values
    unique_week = np.sort(unique_week)
    ctrb_week_1 = []  #create an empty array, count the ctrb of unique values and write to the array
    for index in range(len(unique_week)):
        ctrb_week = df_2022[(df_2022['week'] == unique_week[index])]['ctrb_day'].sum()
        ctrb_week_1.append(ctrb_week)

    # DataFrame creation
    df_week = {'Week': unique_week, 'Ctrb': ctrb_week_1}
    # df_week = df_week.sort_values(by = 'Week')
    dframe_week = pd.DataFrame(df_week)

    #graph  creating
    plt.bar(unique_week, height=ctrb_week_1)
    plt.title("Performance: Week in a year")
    plt.ylabel("Performance, %")
    plt.xlabel("Week number")
    plt.savefig('ctrb_week.png')
    plt.show()

    #creating boxplot
    boxplot = dframe_week.boxplot(column=['Ctrb'])
    boxplot.plot()
    plt.savefig('ctrb_week_boxplot.png')
    plt.show()

    #create an empty sheet, write the dataframe into it
    FilePath = "contribution_data"+file_name
    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    dframe_week.to_excel(writer, sheet_name='ctrb_week')
    writer.save()

    #adding graphical objects to the file
    wb = openpyxl.load_workbook(FilePath)
    # wb.create_sheet('ctrb_week')
    active = wb['ctrb_week']
    active.add_image(Image('ctrb_week.png'), 'F1')
    active.add_image(Image('ctrb_week_boxplot.png'), 'P1')
    wb.save(FilePath)

    print('Week info is written successfully to Excel File.')

    #remove unnecessary images
    os.remove("ctrb_week.png")
    os.remove("ctrb_week_boxplot.png")

def create_graph_month():
    # cal ctrb by month
    df_2022['month'] = pd.to_datetime(df_2022['dt']).dt.month_name()
    unique_month = pd.unique(df_2022['month']) #looking for unique values
    ctrb_month_1 = []
    for index in range(len(unique_month)):
        ctrb_month = df_2022[(df_2022['month'] == unique_month[index])]['ctrb_day'].sum()
        ctrb_month_1.append(ctrb_month)

    # DataFrame creation
    df_month = {'Month': unique_month, 'Ctrb': ctrb_month_1}
    dframe_month = pd.DataFrame(df_month)

    #graph  creating
    plt.bar(unique_month, height=ctrb_month_1)
    plt.title("Performance: Month")
    plt.ylabel("Performance, %")
    plt.savefig('ctrb_month.png')
    plt.show()

    #creating boxplot
    boxplot = dframe_month.boxplot(column=['Ctrb'])
    boxplot.plot()
    plt.savefig('ctrb_month_boxplot.png')
    plt.show()

    #create an empty sheet, write the dataframe into it
    FilePath = "contribution_data" + file_name
    ExcelWorkbook = load_workbook(FilePath)
    writer = pd.ExcelWriter(FilePath, engine='openpyxl')
    writer.book = ExcelWorkbook
    dframe_month.to_excel(writer, sheet_name='ctrb_month')
    writer.save()

    #adding graphical objects to the file
    wb = openpyxl.load_workbook(FilePath)
    # wb.create_sheet('ctrb_month')
    active = wb['ctrb_month']
    active.add_image(Image('ctrb_month.png'), 'F1')
    active.add_image(Image('ctrb_month_boxplot.png'), 'P1')
    wb.save(FilePath)

    print('Month info is written successfully to Excel File.')

    #remove unnecessary images
    os.remove("ctrb_month.png")
    os.remove("ctrb_month_boxplot.png")


create_graph_country()
create_graph_currency()
create_graph_inst_class()
create_graph_sector()
create_graph_day_week()
create_graph_week()
create_graph_month()


