import win32com
import win32com.client as win32
import time
import pyautogui

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True
wb1 = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\BTM Template.xlsx')
ws1 = wb1.Worksheets(1)
wb1.RefreshAll()
time.sleep(5)
pyautogui.click(600, 380) # współrzędne pierwszego konta
# shell = win32.Dispatch("WScript.Shell")
# shell.SendKeys("{ENTER}")
# excel.CalculateUntilAsyncQueriesDone()