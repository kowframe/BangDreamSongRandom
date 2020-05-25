### NOT FOR SALE ! ###
### CREATED BY: Zyptical ###
### twitter.com/zyptical ###
### twitch.tv/zyptical   ###
from random import randrange
import os
import xlrd
import wx

main_dir = os.path.dirname(__file__)
app = wx.App()
frame = wx.Frame(None, -1, 'BanG Dream! Song Random (NOT FOR SALE !)', style= wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX | wx.MINIMIZE_BOX)
frame.SetDimensions(0,0,430,250)
panel = wx.Panel(frame, wx.ID_ANY)

txtTitleJP = wx.TextCtrl(panel, pos=(90, 10), size=(300, -1), style=wx.TE_READONLY)
txtTitleEN =  wx.TextCtrl(panel, pos=(90, 40), size=(300, -1), style=wx.TE_READONLY)
txtTitleBand =  wx.TextCtrl(panel, pos=(90, 70), size=(300, -1), style=wx.TE_READONLY)
txtType =  wx.TextCtrl(panel, pos=(90, 100), size=(300, -1), style=wx.TE_READONLY)

def openExcel():
    DATA_EXCEL_PATH = os.path.join(main_dir, 'src\song_list.xlsx')
    wb = xlrd.open_workbook(DATA_EXCEL_PATH)
    sheet = wb.sheet_by_index(0)
    return sheet

def onButton(event):
    objSong = randomFuction(sheet)

    #Clear Text
    txtTitleJP.Clear()
    txtTitleEN.Clear()
    txtTitleBand.Clear()
    txtType.Clear()

    #Set Text
    txtTitleJP.AppendText(objSong["titleJP"])
    txtTitleEN.AppendText(objSong["titleEN"])
    txtTitleBand.AppendText(objSong["band"])
    txtType.AppendText(objSong["songType"])

def appearTextBox(panel, textLabel, y, x):
    lblname = wx.TextCtrl(panel, value=textLabel, pos=(y, x), size=(300, -1), style=wx.TE_READONLY)

def clearTextCtrl(textLabel, y, x):
    lblname = wx.TextCtrl(panel, value=textLabel, pos=(y, x)).Clear()

def randomFuction(sheet):
    result = randrange(1, (sheet.nrows))
    value = sheet.row_values(result)
    objSong = {
        "titleJP": str(value[0]),
        "titleEN": str(value[1]),
        "band": str(value[2]),
        "songType": str(value[3])
    }
    return objSong

def windowsFormRun():
    button = wx.Button(panel, wx.ID_ANY, 'Random Now', (230, 135))
    button.Bind(wx.EVT_BUTTON, onButton)

    lblTitleJP = wx.StaticText(panel, label="Title (JP): ", pos=(10, 10))
    lblTitleEN = wx.StaticText(panel, label="Title (EN): ", pos=(10, 40))
    lblTitleBand = wx.StaticText(panel, label="Band: ", pos=(10, 70))
    lblType = wx.StaticText(panel, label="Type: ", pos=(10, 100))

    lblNotForSale = wx.StaticText(panel, label="NOT FOR SALE !", pos=(10, 140))
    lblCreatedBy = wx.StaticText(panel, label="CREATED_BY: Zyptical", pos=(110, 140))
    lblTwitter = wx.StaticText(panel, label="twitter.com/zyptical", pos=(10, 160))
    lblTwitch = wx.StaticText(panel, label="twitch.tv/zyptical", pos=(10, 180))

    frame.Show()
    frame.Centre()
    app.MainLoop()

#--------Run time---------------
sheet = openExcel()
windowsFormRun()

#result = randrange(1, SongCapacity +1 )