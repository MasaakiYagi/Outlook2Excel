# -*- coding: utf-8 -*-
"""
Created on Sun Oct  4 17:01:46 2020

@author: Q20712
"""
import win32com.client
import datetime
import calendar

def getStartDateAndEndDateInThisMonth():
    startDate = datetime.date.today()
    startDate = startDate.replace(day=1)
    
    endDate = datetime.date.today()
    last_day = calendar.monthrange(endDate.year, endDate.month)[1]
    endDate = endDate.replace(day=last_day)
    return startDate, endDate

def getCalendarItemsFromTo(startDate,endDate):
    #Outlookの予定取得（全部）
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    appointments = calendar.Items
    
    #任意の期間の予定のみフィルタリング
    restriction = "[Start] >= '" + startDate.strftime("%m/%d/%Y") + "' AND [End] <= '" +endDate.strftime("%m/%d/%Y") + "'"
    restrictedItems = appointments.Restrict(restriction)
    
    restrictedItems.Sort("[Start]")
    
    return restrictedItems

def extractStarItems(allItems,user_fullname):
    starItems = []
    for item in allItems:
        sub = item.Subject
        org = item.Organizer
        if len(sub)>0 and sub[0] == "★" and org == user_fullname:
            starItems.append(item)
    
    return starItems

def excelInputData(starItems):
    excel_input_data = []
    for item in starItems:
        #タイトル⇒入力文字の変換
        original_sub = item.Subject
        if original_sub == "★テレワーク":
            input_value = "〇"
        elif original_sub == "★休日":
            input_value = "●"
        else:
            input_value = original_sub[1:]
        
        #日付作成
        date = datetime.date(item.Start.year,item.Start.month,item.Start.day)
        #辞書配列の作成
        datam = {'date':date, 'value':input_value}
        excel_input_data.append(datam)
    
    return excel_input_data
    
    
