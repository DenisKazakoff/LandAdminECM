#!/usr/bin/env python
# -*- coding: utf-8 -*-python

import pathlib
import datetime
import calendar
import modules

import wx
import wx.xrc
import wx.adv
import wx.lib.scrolledpanel as scrlpanel
import wx.lib.agw.flatnotebook as flatnb

ID_NEW = 1001
ID_OPEN = 1002
ID_SAVE = 1003
ID_SAVE_AS = 1004
ID_CHK_LST = 1011
ID_CLOSE = 1098
ID_EXIT = 1099
ID_MANUAL = 5001
ID_ABOUT = 5002


###########################################################################
## Class AppUZR
###########################################################################
class AppUZR(wx.App):

    def __init__(self):
        super().__init__()
        # declaration
        self.formMain = modules.FrameMain(None).Show(True)
        return

###########################################################################
## Class TabCommon
###########################################################################
class TabCommon ( scrlpanel.ScrolledPanel ):

    def __init__( self, parent ):
        super().__init__( parent = parent )
        self.Hide()
        
        dataSource = modules.DataSource()
        listCityNames = dataSource.RetrieveData(
            'common','cityname', 'alias')
        listServNames = dataSource.RetrieveData(
            'common','servname', 'alias')
        listClientTypes = dataSource.RetrieveData(
            'common','clienttype', 'alias')
        listMonths = list(calendar.month_name)
            
        self.txtSkpNum = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 250,-1 ), 0 )
        sizerSkpNum = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер в СИНКОПе" )
        sizerSkpNum.Add( self.txtSkpNum )

        self.spnSkpDay = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcSkpMonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), choices = listMonths, style = 0 )
        self.chcSkpMonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )
        
        self.spnSkpYear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y")), 
            initial=int(datetime.datetime.today().strftime("%Y")) )
        
        sizerSkpDate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата регистрации в СИНКОПе" )
        sizerSkpDate.Add( self.spnSkpDay )
        sizerSkpDate.Add( self.chcSkpMonth )
        sizerSkpDate.Add( self.spnSkpYear )

        self.txtRegNum = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 250,-1 ), 0 )
        sizerRegNum = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер регистрации в журнале" )
        sizerRegNum.Add( self.txtRegNum )

        self.spnRegDay = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcRegMonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), choices = listMonths, style = 0 )
        self.chcRegMonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )
        
        self.spnRegYear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y")), 
            initial=int(datetime.datetime.today().strftime("%Y")) )
        
        sizerRegDate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата регистрации в журнале" )
        sizerRegDate.Add( self.spnRegDay )
        sizerRegDate.Add( self.chcRegMonth )
        sizerRegDate.Add( self.spnRegYear )

        self.chcCityName = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 250,-1 ), choices = listCityNames, style = 0 )
        self.chcCityName.SetSelection( 0 )
        sizerCityName = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Сельское поселение" )
        sizerCityName.Add( self.chcCityName )
            
        self.chcClientType = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 300,-1 ), choices = listClientTypes, style = 0 )
        self.chcClientType.SetSelection( 0 )
        sizerClientType = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Статус заявителя" )
        sizerClientType.Add( self.chcClientType )

        self.chcServName = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 250,-1 ), choices = listServNames, style = 0 )
        self.chcServName.SetSelection( 0 )
        sizerServName = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Наименование услуги" )
        sizerServName.Add( self.chcServName )

        self.cbxConstruction = wx.CheckBox(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = wx.DefaultSize, label = u"Имеется", style = 0 )
        sizerConstruction = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Возможность строительства" )
        sizerConstruction.Add( self.cbxConstruction )

        self.cbxNeedAgent = wx.CheckBox(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = wx.DefaultSize, label = u"Да", style = 0 )
        self.cbxNeedAgent.Disable( )
        sizerNeedAgent = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Заявитель обращается через представителя" )
        sizerNeedAgent.Add( self.cbxNeedAgent )

        grid = wx.GridBagSizer( 0, 0 )
        grid.SetFlexibleDirection( wx.BOTH )
        grid.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        grid.Add( 
            sizerSkpNum, wx.GBPosition( 0, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerSkpDate, wx.GBPosition( 0, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerRegNum, wx.GBPosition( 1, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerRegDate, wx.GBPosition( 1, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerCityName, wx.GBPosition( 2, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerClientType, wx.GBPosition( 2, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerServName, wx.GBPosition( 3, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerConstruction, wx.GBPosition( 4, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerNeedAgent, wx.GBPosition( 4, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )

        gsizer = wx.BoxSizer( wx.HORIZONTAL )
        gsizer.Add( 
            grid, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 0 )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add(
            gsizer, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 20 )
        self.SetSizer( vsizer )
        self.SetupScrolling( )
        #vsizer.SetSizeHints( self )
        vsizer.Layout( )
        self.Show( )

        # Events
        self.txtSkpNum.Bind(
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnSkpDay.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnSkpDay.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.chcSkpMonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.spnSkpYear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnSkpYear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        
        self.txtRegNum.Bind(
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnRegDay.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnRegDay.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.chcRegMonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.spnRegYear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnRegYear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )

        self.chcCityName.Bind(
            wx.EVT_CHOICE, self.OnChoice_CityName )
        self.chcServName.Bind(
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.chcClientType.Bind(
            wx.EVT_CHOICE, self.OnChoice_ClientType )
        self.cbxConstruction.Bind(
            wx.EVT_CHECKBOX, self.OnChnage_AnyCtrl )
        self.cbxNeedAgent.Bind(
            wx.EVT_CHECKBOX, self.OnCheckBox_Agent )
        self.spnSkpDay.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.chcSkpMonth.Bind(
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.spnSkpYear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        return

    def OnChnage_AnyCtrl (self, event):
        event.Skip()
        parent = self.GetParent().GetParent() # PanelDeclarant
        parent.OnChnage_AnyCtrl( event )
        return

    def OnChoice_CityName( self, event ):
        event.Skip()
        parent = self.GetParent().GetParent().GetParent() # MainForm
        parent.RefreshMainMenu( )
        self.OnChnage_AnyCtrl( event )
        return

    def OnChoice_ClientType( self, event ):
        event.Skip()
        parent = self.GetParent().GetParent() # PanelDeclarant
        parent.RefreshTabDeclarant( )
        self.OnChnage_AnyCtrl( event )
        return

    def OnCheckBox_Agent( self, event ):
        event.Skip()
        parent = self.GetParent().GetParent() # PanelDeclarant
        parent.RefreshTabAgent( )
        self.OnChnage_AnyCtrl( event )
        return

###########################################################################
## Class PanelPerson
###########################################################################
class PanelPerson ( scrlpanel.ScrolledPanel ):

    def __init__( self, parent ):
        super().__init__( parent = parent )
        
        dataSource = modules.DataSource()
        listIDtypes = dataSource.RetrieveData(
            'person','idtype', 'alias')
        listMonths = list(calendar.month_name)

        self.txtName2 = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 200,-1 ), 0 )
        sizerName2 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Фамилия" )
        sizerName2.Add( self.txtName2 )

        self.txtName1 = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 200,-1 ), 0 )
        sizerName1 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Имя" )
        sizerName1.Add( self.txtName1 )

        self.txtName3 = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 200,-1 ), 0 )
        sizerName3 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Отчество" )
        sizerName3.Add( self.txtName3 )

        self.spnBirthDay = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcBirthMonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), choices = listMonths, style = 0 )
        self.chcBirthMonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )
        
        self.spnBirthYear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y"))-18, 
            initial=int(datetime.datetime.today().strftime("%Y"))-18 )
        
        sizerBirthDate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата рождения" )
        sizerBirthDate.Add( self.spnBirthDay )
        sizerBirthDate.Add( self.chcBirthMonth )
        sizerBirthDate.Add( self.spnBirthYear )

        self.cmbxIDtype = wx.ComboBox(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 420,-1 ), style = 0, 
            choices = listIDtypes )
        self.cmbxIDtype.SetSelection(1)
        sizerIDtype = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Документ, удостоверяющий личность" )
        sizerIDtype.Add( self.cmbxIDtype )

        self.spnIDday = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcIDmonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), style = 0,
            choices = listMonths )
        self.chcIDmonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )

        self.spnIDyear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y")), 
            initial=int(datetime.datetime.today().strftime("%Y")) )

        sizerIDdate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата выдачи удостоверяющего документа" )
        sizerIDdate.Add( self.spnIDday )
        sizerIDdate.Add( self.chcIDmonth )
        sizerIDdate.Add( self.spnIDyear )

        self.txtIDwho = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 650,-1 ), 
            style = 0)        
        sizerIDwho = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Кем выдан" )
        sizerIDwho.Add( self.txtIDwho )

        self.txtIDnum1 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 200,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDnum1.SetMaxLength( 4 )
        sizerIDnum1 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Серия документа" )
        sizerIDnum1.Add( self.txtIDnum1 )

        self.txtIDnum2 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 200,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDnum2.SetMaxLength( 6 )
        sizerIDnum2 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер документа" )
        sizerIDnum2.Add( self.txtIDnum2 )

        self.txtIDcode1 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 90,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDcode1.SetMaxLength( 3 )
        self.txtIDcode2 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 90,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDcode2.SetMaxLength( 3 )
        sizerIDcode = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Код подразделения из паспорта" )
        sizerIDcode.Add( self.txtIDcode1 )
        sizerIDcode.Add( self.txtIDcode2 )

        self.txtIDaddr = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 650,-1 ), 
            style = 0)        
        sizerIDaddr = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Адрес регистрации из паспорта" )
        sizerIDaddr.Add( self.txtIDaddr )

        self.txtPhoneNum = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT, 
            value = '' )
        self.txtPhoneNum.SetMaxLength( 14 )
        sizerPhoneNum = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер телефона" )
        sizerPhoneNum.Add( self.txtPhoneNum )

        self.txtEmail = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = 0, 
            value = '' )
        sizerEmail = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Электронная почта" )
        sizerEmail.Add( self.txtEmail )

        self.txtTaxID = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT, 
            value = '' )
        self.txtTaxID.SetMaxLength( 12 )
        self.sizerTaxID = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"ИНН" )
        self.sizerTaxID.Add( self.txtTaxID )

        self.txtOGRNIP = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 200,-1 ), style = wx.TE_RIGHT,
            value = u"" )
        self.txtOGRNIP.SetMaxLength( 15 )
        self.sizerOGRNIP = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"ОГРНИП" )
        self.sizerOGRNIP.Add( self.txtOGRNIP )

        grid = wx.GridBagSizer( 0, 0 )
        grid.SetFlexibleDirection( wx.BOTH )
        grid.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        grid.Add( 
            sizerName2, wx.GBPosition( 0, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerName1, wx.GBPosition( 0, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerName3, wx.GBPosition( 0, 2 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerBirthDate, wx.GBPosition( 1, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerIDtype, wx.GBPosition( 2, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerIDdate, wx.GBPosition( 3, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerIDwho, wx.GBPosition( 4, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )
        
        grid.Add( 
            sizerIDnum1, wx.GBPosition( 5, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerIDnum2, wx.GBPosition( 5, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerIDcode, wx.GBPosition( 5, 2 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerIDaddr, wx.GBPosition( 6, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerPhoneNum, wx.GBPosition( 7, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerEmail, wx.GBPosition( 7, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            self.sizerTaxID, wx.GBPosition( 8, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            self.sizerOGRNIP, wx.GBPosition( 8, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
       
        gsizer = wx.BoxSizer( wx.HORIZONTAL )
        gsizer.Add(
            grid, 0, 
            wx.ALIGN_CENTER_VERTICAL|wx.ALL, 0 )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add(
            gsizer, 0, 
            wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 20 )
        self.SetSizer( vsizer )
        self.SetupScrolling( )
        #vsizer.SetSizeHints( self )
        vsizer.Layout()
        #self.Show()
        
        # Events
        self.txtName1.Bind(
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtName2.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtName3.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDnum1.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDnum2.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDwho.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDcode1.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDcode2.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDaddr.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtPhoneNum.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtEmail.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtTaxID.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtOGRNIP.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnBirthDay.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnBirthDay.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnBirthYear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnBirthYear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnIDday.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnIDday.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnIDyear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnIDyear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.chcBirthMonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.chcIDmonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.cmbxIDtype.Bind( 
            wx.EVT_COMBOBOX, self.OnChnage_cmbxIDtype )
        self.cmbxIDtype.Bind( 
            wx.EVT_TEXT, self.OnChnage_cmbxIDtype )
        return
   
    def OnChnage_AnyCtrl (self, event):
        event.Skip()
        parent = self.GetParent().GetParent().GetParent() # PanelDeclarant
        parent.OnChnage_AnyCtrl( event )
        return
        
    def OnChnage_cmbxIDtype (self, event):
        event.Skip()
        if ( self.cmbxIDtype.GetValue() == \
            "Паспорт гражданина Российской Федерации" ):
            self.txtIDnum1.SetMaxLength( 4 )
            self.txtIDnum2.SetMaxLength( 6 )
        else:
            self.txtIDnum1.SetMaxLength( 0 )
            self.txtIDnum2.SetMaxLength( 0 )
        self.txtIDnum1.SetValue( "" )
        self.txtIDnum2.SetValue( "" )
        self.OnChnage_AnyCtrl( event )
        return

###########################################################################
## Class PanelFirm
###########################################################################
class PanelFirm ( scrlpanel.ScrolledPanel ):

    def __init__( self, parent ):
        super().__init__( parent = parent )
        
        dataSource = modules.DataSource()
        listFirmTypes = dataSource.RetrieveData(
            'firm','firmtype')
        listShortTypes = dataSource.RetrieveData(
            'firm','firmtype','alias')
        listMonths = list(calendar.month_name)
      
        self.txtFirmTitle = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 650,-1 ), style = 0,
            value = u"" )
        sizerFirmTitle = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Наименование" )
        sizerFirmTitle.Add( self.txtFirmTitle )

        self.cmbxFirmType = wx.ComboBox(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 420,-1 ), style = 0,
            choices = listFirmTypes )
        sizerFirmType = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Организационно-правовая форма" )
        sizerFirmType.Add( self.cmbxFirmType )

        self.txtShortType = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = 0,
            value = u"" )
        sizerShortType = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Сокращенная форма" )
        sizerShortType.Add( self.txtShortType )

        self.spnRegDay = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcRegMonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), style = 0,
            choices = listMonths )
        self.chcRegMonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )

        self.spnRegYear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y")), 
            initial=int(datetime.datetime.today().strftime("%Y")) )

        sizerRegDate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата регистрации" )
        sizerRegDate.Add( self.spnRegDay )
        sizerRegDate.Add( self.chcRegMonth )
        sizerRegDate.Add( self.spnRegYear )

        self.txtRegOrg = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 650,-1 ), style = 0,
            value = u"" )
        sizerRegOrg = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Регистрирующий орган" )
        sizerRegOrg.Add( self.txtRegOrg )

        self.txtRegAddr = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 650,-1 ), style = 0,
            value = u"" )
        sizerRegAddr = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Юридический адрес" )
        sizerRegAddr.Add( self.txtRegAddr )

        self.txtOGRN = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT,
            value = u"" )
        self.txtOGRN.SetMaxLength( 13 )
        sizerOGRN = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"ОГРН" )
        sizerOGRN.Add( self.txtOGRN )

        self.txtTaxID = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT,
            value = u"" )
        self.txtTaxID.SetMaxLength( 10 )
        sizerTaxID = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"ИНН" )
        sizerTaxID.Add( self.txtTaxID )
        
        self.txtKPP = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT,
            value = u"" )
        self.txtKPP.SetMaxLength( 9 )
        sizerKPP = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"КПП" )
        sizerKPP.Add( self.txtKPP )

        self.txtPhoneNum = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT,
            value = u"" )
        self.txtPhoneNum.SetMaxLength( 14 )
        sizerPhoneNum = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер телефона" )
        sizerPhoneNum.Add( self.txtPhoneNum )

        self.txtEmail = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT,
            value = u"" )
        sizerEmail = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Электронная почта" )
        sizerEmail.Add( self.txtEmail )

        grid = wx.GridBagSizer( 0, 0 )
        grid.SetFlexibleDirection( wx.BOTH )
        grid.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        grid.Add( 
            sizerFirmTitle, wx.GBPosition( 0, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerFirmType, wx.GBPosition( 1, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerShortType, wx.GBPosition( 1, 2 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerRegDate, wx.GBPosition( 2, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerRegOrg, wx.GBPosition( 3, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerRegAddr, wx.GBPosition( 4, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerOGRN, wx.GBPosition( 5, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerTaxID, wx.GBPosition( 5, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerKPP, wx.GBPosition( 5, 2 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerPhoneNum, wx.GBPosition( 6, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerEmail, wx.GBPosition( 6, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )

        gsizer = wx.BoxSizer( wx.HORIZONTAL )
        gsizer.Add( 
            grid, 0, 
            wx.ALIGN_CENTER_VERTICAL|wx.ALL, 0 )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add( 
            gsizer, 0, 
            wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 20 )
        self.SetSizer( vsizer )
        self.SetupScrolling( )
        #vsizer.SetSizeHints( self )
        vsizer.Layout()
        #self.Show()

        # Events
        self.cmbxFirmType.Bind( 
            wx.EVT_COMBOBOX, self.OnChnage_cmbxFirmType )

        self.txtFirmTitle.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtShortType.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtRegAddr.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtRegOrg.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtOGRN.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtTaxID.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKPP.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnRegDay.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnRegDay.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnRegYear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnRegYear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.chcRegMonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )

        self.cmbxFirmType.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        return

    def OnChnage_AnyCtrl (self, event):
        event.Skip()
        parent = self.GetParent().GetParent().GetParent() # PanelDeclarant
        parent.OnChnage_AnyCtrl( event )
        return

    def OnChnage_cmbxFirmType (self, event):
        event.Skip()
        dataSource = modules.DataSource()
        listShortTypes = dataSource.RetrieveData(
            'firm','firmtype','alias')
        selection = self.cmbxFirmType.GetSelection()
        self.txtShortType.SetValue( 
            listShortTypes[ selection ] )
        self.OnChnage_AnyCtrl( event )
        return

###########################################################################
## Class TabDeclarant
###########################################################################
class TabDeclarant ( wx.Panel ):

    def __init__( self, parent ):
        super().__init__( 
            parent = parent )
        self.Hide()
        self.sizer = wx.BoxSizer( wx.VERTICAL )
        self.SetSizer( self.sizer )
        self.sizer.SetSizeHints( self )
        self.sizer.Layout( )
        self.Show( )
        return

###########################################################################
## Class TabAgent
###########################################################################
class TabAgent ( scrlpanel.ScrolledPanel ):

    def __init__( self, parent ):
        super().__init__( parent = parent )
        self.Hide()

        dataSource = modules.DataSource()
        listIDtypes = dataSource.RetrieveData(
            'agent','idtype', 'alias' )
        listDocTypes = dataSource.RetrieveData(
            'agent','doctype', 'alias' )
        listMonths = list(calendar.month_name)

        self.txtName2 = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 200,-1 ), 0 )
        sizerName2 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Фамилия" )
        sizerName2.Add( self.txtName2 )

        self.txtName1 = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 200,-1 ), 0 )
        sizerName1 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Имя" )
        sizerName1.Add( self.txtName1 )

        self.txtName3 = wx.TextCtrl(
            self, wx.ID_ANY, u"",
            wx.DefaultPosition, ( 200,-1 ), 0 )
        sizerName3 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Отчество" )
        sizerName3.Add( self.txtName3 )

        self.spnBirthDay = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcBirthMonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), choices = listMonths, style = 0 )
        self.chcBirthMonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )
        
        self.spnBirthYear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT,
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y"))-18, 
            initial=int(datetime.datetime.today().strftime("%Y"))-18 )
        
        sizerBirthDate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата рождения" )
        sizerBirthDate.Add( self.spnBirthDay )
        sizerBirthDate.Add( self.chcBirthMonth )
        sizerBirthDate.Add( self.spnBirthYear )

        self.cmbxIDtype = wx.ComboBox(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 420,-1 ), style = 0, 
            choices = listIDtypes )
        self.cmbxIDtype.SetSelection(1)
        sizerIDtype = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Документ, удостоверяющий личность" )
        sizerIDtype.Add( self.cmbxIDtype )

        self.spnIDday = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcIDmonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), style = 0, 
            choices = listMonths )
        self.chcIDmonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )

        self.spnIDyear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y")), 
            initial=int(datetime.datetime.today().strftime("%Y")) )

        sizerIDdate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата выдачи удостоверяющего документа" )
        sizerIDdate.Add( self.spnIDday )
        sizerIDdate.Add( self.chcIDmonth )
        sizerIDdate.Add( self.spnIDyear )

        self.txtIDwho = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 650,-1 ), style = 0, value = u"") 
        sizerIDwho = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Кем выдан" )
        sizerIDwho.Add( self.txtIDwho )

        self.txtIDnum1 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 200,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDnum1.SetMaxLength( 4 )
        sizerIDnum1 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Серия документа" )
        sizerIDnum1.Add( self.txtIDnum1 )

        self.txtIDnum2 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 200,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDnum2.SetMaxLength( 6 )
        sizerIDnum2 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер документа" )
        sizerIDnum2.Add( self.txtIDnum2 )
        
        self.txtIDcode1 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 90,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDcode1.SetMaxLength( 3 )
        self.txtIDcode2 = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 90,-1 ), 
            style = wx.TE_RIGHT)
        self.txtIDcode2.SetMaxLength( 3 )
        sizerIDcode = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Код подразделения из паспорта" )
        sizerIDcode.Add( self.txtIDcode1 )
        sizerIDcode.Add( self.txtIDcode2 )

        self.txtIDaddr = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 650,-1 ), 
            style = 0)        
        sizerIDaddr = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Адрес регистрации из паспорта" )
        sizerIDaddr.Add( self.txtIDaddr )

        self.txtPhoneNum = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = wx.TE_RIGHT, 
            value = '' )
        self.txtPhoneNum.SetMaxLength( 14 )
        sizerPhoneNum = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер телефона" )
        sizerPhoneNum.Add( self.txtPhoneNum )

        self.txtEmail = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 200,-1 ), style = 0, 
            value = '' )
        sizerEmail = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Электронная почта" )
        sizerEmail.Add( self.txtEmail )

        self.cmbxDocType = wx.ComboBox(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 420,-1 ), style = 0, 
            value=  wx.EmptyString, 
            choices = listDocTypes )
        self.cmbxDocType.SetSelection( 1 )
        sizerDocType = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Документ, удостоверяющий полномочия" )
        sizerDocType.Add( self.cmbxDocType )

        self.txtDocNum = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 200,-1 ), style = wx.TE_RIGHT,
            value = u"" )
        sizerDocNum = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Номер документа" )
        sizerDocNum.Add( self.txtDocNum )

        self.spnDocDay = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "", 
            min=1, 
            max=31, 
            initial=int(datetime.datetime.today().strftime("%d")) )
        
        self.chcDocMonth = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 150,-1 ), style = 0, 
            choices = listMonths )
        self.chcDocMonth.SetSelection(
            int(datetime.datetime.today().strftime("%m")) )

        self.spnDocYear = wx.SpinCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 70,-1 ), style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            value = "",
            min=int(datetime.datetime.today().strftime("%Y"))-100, 
            max=int(datetime.datetime.today().strftime("%Y")), 
            initial=int(datetime.datetime.today().strftime("%Y")) )
        
        sizerDocDate = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Дата выдачи удостоверяющего документа" )
        sizerDocDate.Add( self.spnDocDay )
        sizerDocDate.Add( self.chcDocMonth )
        sizerDocDate.Add( self.spnDocYear )

        self.txtDocWho = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 650,-1 ), style = 0, 
            value = u"") 
        sizerDocWho = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, u"Кем выдан" )
        sizerDocWho.Add( self.txtDocWho )

        grid = wx.GridBagSizer( 0, 0 )
        grid.SetFlexibleDirection( wx.BOTH )
        grid.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        grid.Add( 
            sizerName2, wx.GBPosition( 0, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerName1, wx.GBPosition( 0, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerName3, wx.GBPosition( 0, 2 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerBirthDate, wx.GBPosition( 1, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerIDtype, wx.GBPosition( 2, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerIDdate, wx.GBPosition( 3, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerIDwho, wx.GBPosition( 4, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )
        
        grid.Add( 
            sizerIDnum1, wx.GBPosition( 5, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerIDnum2, wx.GBPosition( 5, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerIDcode, wx.GBPosition( 5, 2 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerIDaddr, wx.GBPosition( 6, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerPhoneNum, wx.GBPosition( 7, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerEmail, wx.GBPosition( 7, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )

        grid.Add( 
            sizerDocType, wx.GBPosition( 9, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerDocNum, wx.GBPosition( 9, 2 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerDocDate, wx.GBPosition( 10, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerDocWho, wx.GBPosition( 11, 0 ), 
            wx.GBSpan( 1, 3 ), wx.EXPAND|wx.ALL, 5 )

        gsizer = wx.BoxSizer( wx.HORIZONTAL )
        gsizer.Add( grid, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 0 )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add( gsizer, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 20 )
        self.SetSizer( vsizer )
        self.SetupScrolling( )
        #vsizer.SetSizeHints( self )
        vsizer.Layout( )
        self.Show( )

        # Events
        self.txtName1.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtName2.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtName3.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDnum1.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDnum2.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDwho.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDcode1.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDcode2.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtIDaddr.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtPhoneNum.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtEmail.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtDocWho.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtDocNum.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnBirthDay.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnBirthDay.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnBirthYear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnBirthYear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnIDday.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnIDday.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnIDyear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnIDyear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnDocDay.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnDocDay.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.spnDocYear.Bind( 
            wx.EVT_SPINCTRL, self.OnChnage_AnyCtrl )
        self.spnDocYear.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )        
        self.chcBirthMonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.chcIDmonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.chcDocMonth.Bind( 
            wx.EVT_CHOICE, self.OnChnage_AnyCtrl )        
        self.cmbxIDtype.Bind( 
            wx.EVT_COMBOBOX, self.OnChnage_cmbxIDtype )
        self.cmbxIDtype.Bind( 
            wx.EVT_TEXT, self.OnChnage_cmbxIDtype )
        self.cmbxDocType.Bind( 
            wx.EVT_COMBOBOX, self.OnChnage_AnyCtrl )
        self.cmbxDocType.Bind( 
            wx.EVT_TEXT, self.OnChnage_AnyCtrl )       
        return

    def OnChnage_AnyCtrl (self, event):
        event.Skip()
        parent = self.GetParent().GetParent()
        parent.OnChnage_AnyCtrl( event )
        return

    def OnChnage_cmbxIDtype (self, event):
        event.Skip()
        if ( self.cmbxIDtype.GetValue() == \
            "Паспорт гражданина Российской Федерации" ):
            self.txtIDnum1.SetMaxLength( 4 )
            self.txtIDnum2.SetMaxLength( 6 )
        else:
            self.txtIDnum1.SetMaxLength( 0 )
            self.txtIDnum2.SetMaxLength( 0 )
        self.txtIDnum1.SetValue( "" )
        self.txtIDnum2.SetValue( "" )
        self.OnChnage_AnyCtrl( event )
        return

###########################################################################
## Class TabArea
###########################################################################
class TabArea ( scrlpanel.ScrolledPanel ):

    def __init__( self, parent ):
        super().__init__( parent = parent )
        self.Hide()

        dataSource = modules.DataSource()
        listAreaRights = dataSource.RetrieveData(
            'area','rights', 'alias' )
        listAreaCategories = dataSource.RetrieveData(
            'area','category', 'alias' )
        listPlaces = dataSource.RetrieveData(
            'area','place', 'alias' )
        listAreaTargets = dataSource.RetrieveData(
            'area','target', 'alias' )
        listAreaTerZones = dataSource.RetrieveData(
            'area','terzone')
        listAreaSpcZones = dataSource.RetrieveData(
            'area','spczone')
       
        self.cmbxPlace = wx.ComboBox(           
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 770,-1 ), style = 0, 
            value = wx.EmptyString, 
            choices = listPlaces )
        sizerPlace = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Местоположение" )
        sizerPlace.Add( self.cmbxPlace )
      
        self.txtAddr = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 770,-1 ), style = wx.TE_LEFT, 
            value = wx.EmptyString )
        sizerAddr = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Почтовый адрес" )
        sizerAddr.Add( self.txtAddr )

        self.chcRights = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 375,-1 ), style = 0, 
            choices = listAreaRights )
        self.chcRights.SetSelection( 0 )
        sizerRights = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Вид права,указанный в заявлении" )
        sizerRights.Add( self.chcRights )

        self.txtSize = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 90,-1 ), style = wx.TE_RIGHT, 
            value = wx.EmptyString )
        sizerSize = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Площадь участка, м2" )
        sizerSize.Add( self.txtSize )

        self.txtKN11 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 35,-1 ), style = wx.TE_RIGHT, 
            value = u"23" )
        self.txtKN11.SetMaxLength( 2 )
        self.txtKN11.Enable( False )

        self.txtKN12 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 35,-1 ), style = wx.TE_RIGHT,
            value = u"15" )
        self.txtKN12.SetMaxLength( 2 )
        self.txtKN12.Enable( False )

        self.txtKN13 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 85,-1 ), style = wx.TE_RIGHT,
            value = wx.EmptyString )
        self.txtKN13.SetMaxLength( 7 )

        self.txtKN14 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 85,-1 ), style = wx.TE_RIGHT,
            value = wx.EmptyString )

        sizerKN1 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Кадастровый номер/квартал земельного участка" )
        sizerKN1.Add( self.txtKN11 )
        sizerKN1.Add( self.txtKN12 )
        sizerKN1.Add( self.txtKN13 )
        sizerKN1.Add( self.txtKN14 )

        self.txtKN21 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 35,-1 ), style = wx.TE_RIGHT, 
            value = u"23" )
        self.txtKN21.SetMaxLength( 2 )
        self.txtKN21.Enable( False )

        self.txtKN22 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 35,-1 ), style = wx.TE_RIGHT,
            value = u"15" )
        self.txtKN22.SetMaxLength( 2 )
        self.txtKN22.Enable( False )

        self.txtKN23 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 85,-1 ), style = wx.TE_RIGHT,
            value = wx.EmptyString )
        self.txtKN23.SetMaxLength( 7 )

        self.txtKN24 = wx.TextCtrl(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 85,-1 ), style = wx.TE_RIGHT,
            value = wx.EmptyString )

        sizerKN2 = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Кадастровый номер ближайшего земельного участка" )
        sizerKN2.Add( self.txtKN21 )
        sizerKN2.Add( self.txtKN22 )
        sizerKN2.Add( self.txtKN23 )
        sizerKN2.Add( self.txtKN24 )

        self.chcCategory = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 375,-1 ), style = 0,
            choices = listAreaCategories )
        self.chcCategory.SetSelection( 0 )
        sizerCategory = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Категория земель" )
        sizerCategory.Add( self.chcCategory )

        self.cmbxTarget = wx.ComboBox(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition,
            size = ( 375,-1 ), style = 0, 
            value = wx.EmptyString, 
            choices = listAreaTargets )
        self.cmbxTarget.SetSelection( 0 )
        sizerTarget = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Цели использования" )
        sizerTarget.Add( self.cmbxTarget )

        self.chcTerZones = wx.Choice(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 770,-1 ), style = 0,
            choices = listAreaTerZones )
        self.chcTerZones.SetSelection( 0 )
        sizerTerZones = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Территориальная зона" )
        sizerTerZones.Add( self.chcTerZones )

        self.lstbSpcZones = wx.ListBox(         
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 730, 100 ), style = wx.LB_ALWAYS_SB|wx.LB_MULTIPLE,
            choices = [] )
        self.btnSpcZones = wx.Button(
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( 30,-1 ), style = 0,
            label = "﹀" )
        sizerSpcZones = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            u"Зоны с особыми условиями использования территорий" )
        sizerSpcZones.Add( self.lstbSpcZones )
        sizerSpcZones.Add( self.btnSpcZones )

        grid = wx.GridBagSizer( 0, 0 )
        grid.SetFlexibleDirection( wx.BOTH )
        grid.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )

        grid.Add( 
            sizerPlace, wx.GBPosition( 0, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerAddr, wx.GBPosition( 1, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerRights, wx.GBPosition( 2, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerSize, wx.GBPosition( 2, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerKN1, wx.GBPosition( 3, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerKN2, wx.GBPosition( 3, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerCategory, wx.GBPosition( 4, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerTarget, wx.GBPosition( 4, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerTerZones, wx.GBPosition( 5, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerSpcZones, wx.GBPosition( 6, 0 ), 
            wx.GBSpan( 1, 2 ), wx.EXPAND|wx.ALL, 5 )

        gsizer = wx.BoxSizer( wx.HORIZONTAL )
        gsizer.Add( grid, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 0 )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add( gsizer, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 20 )
        self.SetSizer( vsizer )
        self.SetupScrolling( )
        #vsizer.SetSizeHints( self )
        vsizer.Layout( )        
        self.Show( )

        # Events
        self.txtAddr.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtSize.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN11.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN12.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN13.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN14.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN21.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN22.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN23.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtKN24.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )

        self.chcRights.Bind( wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.chcCategory.Bind( wx.EVT_CHOICE, self.OnChnage_AnyCtrl )
        self.chcTerZones.Bind( wx.EVT_CHOICE, self.OnChnage_AnyCtrl )

        self.cmbxPlace.Bind( wx.EVT_COMBOBOX, self.OnChnage_AnyCtrl )
        self.cmbxPlace.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.cmbxTarget.Bind( wx.EVT_COMBOBOX, self.OnChnage_AnyCtrl )
        self.cmbxTarget.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )

        self.btnSpcZones.Bind( wx.EVT_BUTTON, self.OnClickBtn_SpcZones )        
        return

    def OnChnage_AnyCtrl (self, event):
        event.Skip()
        parent = self.GetParent().GetParent()
        parent.OnChnage_AnyCtrl( event )
        return

    def RefrechSpcZones ( self, listZones ):
        dataSource = modules.DataSource()
        listAreaSpcZones = dataSource.RetrieveData(
            'area','spczone', 'alias', True )

        self.lstbSpcZones.Clear( )
        for zoneIdx in listZones:
            zoneTxt = listAreaSpcZones[ zoneIdx ]
            self.lstbSpcZones.Append( zoneTxt )        
        return
      
    def OnClickBtn_SpcZones( self, event ):

        with FrameSpcZones( self ) as dlgSpcZones:
            if ( dlgSpcZones.ShowModal() == wx.ID_OK ):
                selectedSpcZones = \
                    dlgSpcZones.lstbSpcZones.GetSelections( )
                self.RefrechSpcZones( selectedSpcZones )
        event.Skip()
        self.OnChnage_AnyCtrl( event )
        return

###########################################################################
## Class TabRules
###########################################################################
class TabRules ( scrlpanel.ScrolledPanel ):

    def __init__( self, parent ):
        super().__init__( parent = parent )
        self.Hide()
             
        self.txtMinArea = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.TE_RIGHT)
        self.txtMinArea.SetMaxLength( 7 )
        sizerMinArea = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Минимальная площадь участка, м2" )
        sizerMinArea.Add( self.txtMinArea )

        self.txtMaxArea = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.TE_RIGHT)
        self.txtMaxArea.SetMaxLength( 7 )
        sizerMaxArea = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Максимальная площадь участка, м2" )
        sizerMaxArea.Add( self.txtMaxArea )

        self.txtMinWidth = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.TE_RIGHT)
        self.txtMinWidth.SetMaxLength( 4 )
        sizerMinWidth = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Минимальная ширина участков вдоль фронта улицы (проезда), м" )
        sizerMinWidth.Add( self.txtMinWidth )

        self.txtMinOffset = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.TE_RIGHT)
        self.txtMinOffset.SetMaxLength( 2 )
        sizerMinOffset = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Минимальные отступы от границ участков, м" )
        sizerMinOffset.Add( self.txtMinOffset )

        self.txtMaxFloors = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.TE_RIGHT)
        self.txtMaxFloors.SetMaxLength( 2 )
        sizerMaxFloors = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Максимальное количество надземных этажей" )
        sizerMaxFloors.Add( self.txtMaxFloors )

        self.txtMaxHeight = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.TE_RIGHT)
        self.txtMaxHeight.SetMaxLength( 3 )
        sizerMaxHeight = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Максимальная высота строений, сооружений от уровня земли, м" )
        sizerMaxHeight.Add( self.txtMaxHeight )

        self.spnMaxDensity = wx.SpinCtrl(
            self, id = wx.ID_ANY, value = "",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            min=0, max=100, initial=100 )
        sizerMaxDensity = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Максимальный процент застройки в границах участка, %" )
        sizerMaxDensity.Add( self.spnMaxDensity )

        self.spnUgndDensity = wx.SpinCtrl(
            self, id = wx.ID_ANY, value = "",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.SP_ARROW_KEYS|wx.ALIGN_RIGHT, 
            min=0, max=100, initial=100 )
        sizerUgndDensity = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Процент застройки подземной части, %" )
        sizerUgndDensity.Add( self.spnUgndDensity )
        self.txtUtilization = wx.TextCtrl(
            self, id = wx.ID_ANY, value = u"",
            pos = wx.DefaultPosition, size = ( 100,-1 ), 
            style = wx.TE_RIGHT)
        self.txtUtilization.SetMaxLength( 5 )
        sizerUtilization = wx.StaticBoxSizer( 
            wx.HORIZONTAL, self, 
            "Коэффициент использования территории" )
        sizerUtilization.Add( self.txtUtilization )

        grid = wx.GridBagSizer( 0, 0 )
        grid.SetFlexibleDirection( wx.BOTH )
        grid.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )

        grid.Add( 
            sizerMinArea, wx.GBPosition( 0, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerMaxArea, wx.GBPosition( 0, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerMinWidth, wx.GBPosition( 1, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerMinOffset, wx.GBPosition( 1, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerMaxFloors, wx.GBPosition( 2, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerMaxHeight, wx.GBPosition( 2, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerMaxDensity, wx.GBPosition( 3, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerUgndDensity, wx.GBPosition( 3, 1 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )
        grid.Add( 
            sizerUtilization, wx.GBPosition( 4, 0 ), 
            wx.GBSpan( 1, 1 ), wx.EXPAND|wx.ALL, 5 )

        gsizer = wx.BoxSizer( wx.HORIZONTAL )
        gsizer.Add( 
            grid, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 0 )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add( 
            gsizer, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 20 )
        self.SetSizer( vsizer )
        self.SetupScrolling( )
        #vsizer.SetSizeHints( self )
        vsizer.Layout( )
        self.Show( )

        # Events
        self.txtMinArea.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtMaxArea.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtMinWidth.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtMinOffset.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtMaxFloors.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtMaxHeight.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.txtUtilization.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnUgndDensity.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        self.spnMaxDensity.Bind( wx.EVT_TEXT, self.OnChnage_AnyCtrl )
        return

    def OnChnage_AnyCtrl (self, event):
        event.Skip()
        parent = self.GetParent().GetParent()
        parent.OnChnage_AnyCtrl( event )
        return

###########################################################################
## Class panelDoc
###########################################################################
class PanelDoc ( wx.Panel ):

    def __init__( self, parent ):
        super().__init__( 
            parent = parent )
        self.Hide()

        nbstyle =   flatnb.FNB_FANCY_TABS|flatnb.FNB_NO_X_BUTTON|\
                    flatnb.FNB_NO_NAV_BUTTONS|flatnb.FNB_NODRAG|\
                    flatnb.FNB_NO_TAB_FOCUS
        self.notebook      = flatnb.FlatNotebook( self, agwStyle =  nbstyle )        
        self.tabCommon     = TabCommon(self.notebook)
        self.tabDeclarant  = TabDeclarant(self.notebook)
        self.tabAgent      = TabAgent(self.notebook)
        self.tabArea       = TabArea(self.notebook)
        self.tabRules      = TabRules(self.notebook)
        
        self.notebook.AddPage( 
            self.tabCommon, 
            u"Начальные сведения", True )
        self.notebook.AddPage( 
            self.tabDeclarant, 
            u"Сведения о заявителе", False )
        self.notebook.EnableTab( 1, enabled = False )      
        self.notebook.AddPage( 
            self.tabAgent, 
            u"Представитель заявителя", False )
        self.notebook.EnableTab( 2, enabled = False )
        self.notebook.AddPage( 
            self.tabArea, 
            u"Сведения об участке", False )
        self.notebook.AddPage( 
            self.tabRules, 
            u"Правила землепользования и застройки", False )
        vSizer_ntbk = wx.BoxSizer( wx.VERTICAL )
        vSizer_ntbk.Add(
            self.notebook, 1, 
            flag=wx.EXPAND|wx.ALL, border=5 )
        gSizer_ntbk = wx.BoxSizer( wx.HORIZONTAL )
        gSizer_ntbk.Add(
            vSizer_ntbk, 1, 
            flag=wx.EXPAND|wx.ALL, border=0 )

        self.btnPrev = wx.Button( 
            self, wx.ID_ANY, u"<<Назад")
        self.btnNext = wx.Button( 
            self, wx.ID_ANY, u"Вперед >>")
        gSizer_btns = wx.BoxSizer( wx.HORIZONTAL )
        gSizer_btns.Add( 
            self.btnPrev, 0, 
            flag=wx.ALL, border=5 )
        gSizer_btns.Add( 
            self.btnNext, 0, 
            flag=wx.ALL, border=5 )

        # Sizers
        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add( 
            gSizer_ntbk, 1, 
            flag=wx.EXPAND|wx.ALL, border=0 )
        vsizer.Add( 
            gSizer_btns, 0, 
            flag=wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, border=0 )
        self.SetSizer( vsizer )
        #vsizer.SetSizeHints( self )
        vsizer.Layout( )
        self.Show( )

        # Content
        self.fileName       = pathlib.Path( "Новый документ.csd" )
        self.caseDocument   = modules.DocumentCase( )

        # Events
        self.btnPrev.Bind( 
            wx.EVT_BUTTON, 
            self.OnClickbtnPrev )
        self.btnNext.Bind( 
            wx.EVT_BUTTON, 
            self.OnClickbtnNext )       
        return

    def RefreshTabAgent( self ):
        if self.tabCommon.cbxNeedAgent.IsChecked( ):
            self.notebook.EnableTab( 2, enabled = True )
        else:
            self.notebook.EnableTab( 2, enabled = False )
        return

    def RefreshTabDeclarant( self ):

        self.notebook.EnableTab( 1, enabled = False )        
        try: 
            self.tabDeclarant.pnlDeclarant.Hide( )
        except Exception as e: print (e)
        else: del(self.tabDeclarant.pnlDeclarant)

        if ( self.tabCommon.chcClientType.GetSelection( ) == 1 ):
            self.tabDeclarant.pnlDeclarant = PanelPerson( self.tabDeclarant )
            self.tabDeclarant.sizer.Add( 
                self.tabDeclarant.pnlDeclarant, 1, 
                flag=wx.EXPAND|wx.ALL, border=0 )
            parent = self.tabDeclarant.pnlDeclarant
            #parent.txtTaxID.Enable( False )
            parent.txtOGRNIP.Enable( False )
            parent.Show()

        if ( self.tabCommon.chcClientType.GetSelection( ) == 3 ):
            self.tabDeclarant.pnlDeclarant = PanelPerson( self.tabDeclarant )
            self.tabDeclarant.sizer.Add( 
                self.tabDeclarant.pnlDeclarant, 1, 
                flag=wx.EXPAND|wx.ALL, border=0 )
            parent = self.tabDeclarant.pnlDeclarant
            #parent.txtTaxID.Enable( True )
            parent.txtOGRNIP.Enable( True )
            parent.Show()

        if ( self.tabCommon.chcClientType.GetSelection( ) == 2 ):
            self.tabDeclarant.pnlDeclarant = PanelFirm( self.tabDeclarant )
            self.tabDeclarant.sizer.Add( 
                self.tabDeclarant.pnlDeclarant, 1, 
                flag=wx.EXPAND|wx.ALL, border=0 )          
            self.tabDeclarant.pnlDeclarant.Show()
            
        if ( self.tabCommon.chcClientType.GetSelection( ) != 0 ):
            self.notebook.EnableTab( 1, enabled = True )
            selectedTab = self.notebook.GetSelection( )
            if ( selectedTab == 1 ):
                self.notebook.SetSelection( 0 )
                self.notebook.SetSelection( 1 )
            self.tabCommon.cbxNeedAgent.SetValue( False )
            self.tabCommon.cbxNeedAgent.Enable( )
            self.RefreshTabAgent()           
        return 
     
    def SetAllTabs( self ):
        for idx in range(5):
            self.SetThisTab( idx )
        return

    def SetThisTab( self, tabIdx):
        caseDoc = self.caseDocument
        
        if ( tabIdx == 0 ):
            parent = self.tabCommon 
            sectionName = 'common'

            parent.txtSkpNum.SetValue( 
                caseDoc.GetTagValue( sectionName, 'sinkopa' ) )
          
            listDate = caseDoc.GetTagValue( 
                sectionName, 'skpdate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnSkpDay.SetValue( int( listDate[ 0 ] ) )
                parent.chcSkpMonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnSkpYear.SetValue( int( listDate[ 2 ] ) )

            parent.txtRegNum.SetValue( 
                caseDoc.GetTagValue( sectionName, 'regnum' ) )

            listDate = caseDoc.GetTagValue( 
                sectionName, 'regdate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnRegDay.SetValue( int( listDate[ 0 ] ) )
                parent.chcRegMonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnRegYear.SetValue( int( listDate[ 2 ] ) )

            try: parent.chcCityName.SetSelection( int( 
                    caseDoc.GetTagValue( sectionName, 'cityname' ) ) )
            except Exception as e: print (e)

            try: parent.chcServName.SetSelection( int( 
                    caseDoc.GetTagValue( sectionName, 'servname' ) ) )
            except Exception as e: print (e)
            
            try: parent.chcClientType.SetSelection( int( 
                    caseDoc.GetTagValue( sectionName, 'clienttype' ) ) )
            except Exception as e: print (e)
            else: self.RefreshTabDeclarant( )

            parent.cbxConstruction.SetValue(  
                caseDoc.GetTagValue( 
                    sectionName, 'construction' ) == 'True' )

            parent.cbxNeedAgent.SetValue(  
                caseDoc.GetTagValue( 
                    sectionName, 'needagent' ) == 'True' )
            self.RefreshTabAgent( )                    

        if ( tabIdx == 1
        and ( self.tabCommon.chcClientType.GetSelection( ) == 1
        or self.tabCommon.chcClientType.GetSelection( ) == 3 ) ):
            parent = self.tabDeclarant.pnlDeclarant
            sectionName = 'person'
            
            parent.txtName1.SetValue( 
                caseDoc.GetTagValue( sectionName, 'name1' ) )
            parent.txtName2.SetValue( 
                caseDoc.GetTagValue( sectionName, 'name2' ) )
            parent.txtName3.SetValue( 
                caseDoc.GetTagValue( sectionName, 'name3' ) )

            listDate = caseDoc.GetTagValue( 
                sectionName, 'birthdate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnBirthDay.SetValue( int( listDate[ 0 ] ) )
                parent.chcBirthMonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnBirthYear.SetValue( int( listDate[ 2 ] ) )

            parent.cmbxIDtype.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idtype' ) )
            parent.txtIDnum1.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idnum1' ) )
            parent.txtIDnum2.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idnum2' ) )
            parent.txtIDwho.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idwho' ) )
            parent.txtIDcode1.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idcode1' ) )
            parent.txtIDcode2.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idcode2' ) )

            listDate = caseDoc.GetTagValue( 
                sectionName, 'iddate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnIDday.SetValue( int( listDate[ 0 ] ) )
                parent.chcIDmonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnIDyear.SetValue( int( listDate[ 2 ] ) )

            parent.txtIDaddr.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idaddr' ) )
            parent.txtPhoneNum.SetValue( 
                caseDoc.GetTagValue( sectionName, 'phonenum' ) )
            parent.txtEmail.SetValue( 
                caseDoc.GetTagValue( sectionName, 'email' ) )
            parent.txtTaxID.SetValue( 
                caseDoc.GetTagValue( sectionName, 'taxid' ) )
            parent.txtOGRNIP.SetValue( 
                caseDoc.GetTagValue( sectionName, 'ogrnip' ) )

        if ( tabIdx == 1
        and self.tabCommon.chcClientType.GetSelection( ) == 2 ):
            parent = self.tabDeclarant.pnlDeclarant
            sectionName = 'firm'
            
            parent.txtFirmTitle.SetValue( 
                caseDoc.GetTagValue( sectionName, 'firmtitle' ) )
            parent.cmbxFirmType.SetValue( 
                caseDoc.GetTagValue( sectionName, 'firmtype' ) )
            parent.txtShortType.SetValue( 
                caseDoc.GetTagValue( sectionName, 'shorttype' ) )
            parent.txtRegAddr.SetValue( 
                caseDoc.GetTagValue( sectionName, 'regaddr' ) )
            parent.txtRegOrg.SetValue( 
                caseDoc.GetTagValue( sectionName, 'regorg' ) )
            
            listDate = caseDoc.GetTagValue( 
                sectionName, 'regdate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnRegDay.SetValue( int( listDate[ 0 ] ) )
                parent.chcRegMonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnRegYear.SetValue( int( listDate[ 2 ] ) )

            parent.txtOGRN.SetValue( 
                caseDoc.GetTagValue( sectionName, 'ogrn' ) )
            parent.txtTaxID.SetValue( 
                caseDoc.GetTagValue( sectionName, 'taxid' ) )
            parent.txtKPP.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kpp' ) )
            parent.txtPhoneNum.SetValue( 
                caseDoc.GetTagValue( sectionName, 'phonenum' ) )
            parent.txtEmail.SetValue( 
                caseDoc.GetTagValue( sectionName, 'email' ) )


        if ( tabIdx == 2
        and self.tabCommon.cbxNeedAgent.GetValue( ) ):
            parent =  self.tabAgent
            sectionName = 'agent'
            
            parent.txtName1.SetValue( 
                caseDoc.GetTagValue( sectionName, 'name1' ) )
            parent.txtName2.SetValue( 
                caseDoc.GetTagValue( sectionName, 'name2' ) )
            parent.txtName3.SetValue( 
                caseDoc.GetTagValue( sectionName, 'name3' ) )

            listDate = caseDoc.GetTagValue( 
                sectionName, 'birthdate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnBirthDay.SetValue( int( listDate[ 0 ] ) )
                parent.chcBirthMonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnBirthYear.SetValue( int( listDate[ 2 ] ) )

            parent.cmbxIDtype.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idtype' ) )
            parent.txtIDnum1.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idnum1' ) )
            parent.txtIDnum2.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idnum2' ) )
            parent.txtIDwho.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idwho' ) )
            parent.txtIDcode1.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idcode1' ) )
            parent.txtIDcode2.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idcode2' ) )

            listDate = caseDoc.GetTagValue( 
                sectionName, 'iddate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnIDday.SetValue( int( listDate[ 0 ] ) )
                parent.chcIDmonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnIDyear.SetValue( int( listDate[ 2 ] ) )

            parent.txtIDaddr.SetValue( 
                caseDoc.GetTagValue( sectionName, 'idaddr' ) )
            parent.txtPhoneNum.SetValue( 
                caseDoc.GetTagValue( sectionName, 'phonenum' ) )
            parent.txtEmail.SetValue( 
                caseDoc.GetTagValue( sectionName, 'email' ) )
            parent.cmbxDocType.SetValue(            
                caseDoc.GetTagValue( sectionName, 'doctype' ) )
            parent.txtDocWho.SetValue(            
                caseDoc.GetTagValue( sectionName, 'docwho' ) )

            listDate = caseDoc.GetTagValue( 
                sectionName, 'docdate' ).split( '-' )
            if ( len( listDate ) == 3 ):
                parent.spnDocDay.SetValue( int( listDate[ 0 ] ) )
                parent.chcDocMonth.SetSelection( int( listDate[ 1 ] ) )
                parent.spnDocYear.SetValue( int( listDate[ 2 ] ) )
                    
            parent.txtDocNum.SetValue( 
                caseDoc.GetTagValue( sectionName, 'docnum' ) )

        if ( tabIdx == 3 ):
            parent =  self.tabArea
            sectionName = 'area'

            parent.cmbxPlace.SetValue( 
                caseDoc.GetTagValue( sectionName, 'place' ) )
            parent.txtAddr.SetValue( 
                caseDoc.GetTagValue( sectionName, 'addr' ) )
            
            try: parent.chcRights.SetSelection( int( 
                    caseDoc.GetTagValue( sectionName, 'rights' ) ) )
            except Exception as e: print (e)
            
            parent.txtSize.SetValue( 
                caseDoc.GetTagValue( sectionName, 'size' ) )
            parent.txtKN11.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn11' ) )
            parent.txtKN12.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn12' ) )
            parent.txtKN13.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn13' ) )
            parent.txtKN14.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn14' ) )
            
            parent.txtKN21.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn21' ) )
            parent.txtKN22.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn22' ) )
            parent.txtKN23.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn23' ) )
            parent.txtKN24.SetValue( 
                caseDoc.GetTagValue( sectionName, 'kn24' ) )

            try: parent.chcCategory.SetSelection( int( 
                    caseDoc.GetTagValue( sectionName, 'category' ) ) )
            except Exception as e: print (e)
                
            parent.cmbxTarget.SetValue( 
                caseDoc.GetTagValue( sectionName, 'target' ) )
                        
            try: parent.chcTerZones.SetSelection( int( 
                    caseDoc.GetTagValue( sectionName, 'terzone' ) ) )
            except Exception as e: print (e)
                    
            selectedSpcZones = caseDoc.GetTagValue( 
                sectionName, 'spczone' ).split( ',' )
            if ( len( selectedSpcZones ) > 0 ):
                try:
                    for idx, zone in enumerate(selectedSpcZones):
                        selectedSpcZones[ idx ] = ( int(zone) - 1 )
                except Exception as e: print (e)
                else: parent.RefrechSpcZones( selectedSpcZones )

        if  ( tabIdx == 4 ):
            parent =  self.tabRules
            sectionName = 'rules'

            parent.txtMinArea.SetValue( 
                caseDoc.GetTagValue( sectionName, 'minarea' ) )
            parent.txtMaxArea.SetValue( 
                caseDoc.GetTagValue( sectionName, 'maxarea' ) )
            parent.txtMinWidth.SetValue( 
                caseDoc.GetTagValue( sectionName, 'minwidth' ) )
            parent.txtMinOffset.SetValue( 
                caseDoc.GetTagValue( sectionName, 'minoffset' ) )
            parent.txtMaxFloors.SetValue( 
                caseDoc.GetTagValue( sectionName, 'maxfloors' ) )
            parent.txtMaxHeight.SetValue( 
                caseDoc.GetTagValue( sectionName, 'maxheight' ) )
            parent.spnUgndDensity.SetValue( 
                caseDoc.GetTagValue( sectionName, 'ugnddensity' ) )
            parent.spnMaxDensity.SetValue( 
                caseDoc.GetTagValue( sectionName, 'maxdensity' ) )
            parent.txtUtilization.SetValue( 
                caseDoc.GetTagValue( sectionName, 'utilization' ) )
        return

    def GetAllTabs( self ):
        for idx in range(5):
            self.GetThisTab( idx )
        return

    def GetThisTab( self, tabIdx ):
        caseDoc = self.caseDocument
        
        if ( tabIdx == 0 ):
            parent = self.tabCommon
            sectionName = 'common'

            caseDoc.SetTagValue( sectionName, 'sinkopa',
                str( parent.txtSkpNum.GetValue( ) ) )

            listDate = [ 
                str( parent.spnSkpDay.GetValue( ) ).zfill( 2 ),
                str( parent.chcSkpMonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnSkpYear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'skpdate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'regnum',
                str( parent.txtRegNum.GetValue( ) ) )

            listDate = [ 
                str( parent.spnRegDay.GetValue( ) ).zfill( 2 ),
                str( parent.chcRegMonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnRegYear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'regdate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'cityname',
                str( parent.chcCityName.GetSelection( ) ) )
            caseDoc.SetTagValue( sectionName, 'servname',
                str( parent.chcServName.GetSelection( ) ) )
            caseDoc.SetTagValue( sectionName, 'clienttype',
                str( parent.chcClientType.GetSelection( ) ) )
            caseDoc.SetTagValue( sectionName, 'construction',
                str( parent.cbxConstruction.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'needagent',
                str( parent.cbxNeedAgent.GetValue( ) ) )
                      
        if ( tabIdx == 1
        and ( self.tabCommon.chcClientType.GetSelection( ) == 1
        or self.tabCommon.chcClientType.GetSelection( ) == 3 ) ):
            parent = self.tabDeclarant.pnlDeclarant
            sectionName = 'person'

            caseDoc.SetTagValue( sectionName, 'name1',
                str( parent.txtName1.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'name2',
                str( parent.txtName2.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'name3',
                str( parent.txtName3.GetValue( ) ) )

            listDate = [ 
                str( parent.spnBirthDay.GetValue( ) ).zfill( 2 ),
                str( parent.chcBirthMonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnBirthYear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'birthdate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'idtype',
                str( parent.cmbxIDtype.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idnum1',
                str( parent.txtIDnum1.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idnum2',
                str( parent.txtIDnum2.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idwho',
                str( parent.txtIDwho.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idcode1',
                str( parent.txtIDcode1.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idcode2',
                str( parent.txtIDcode2.GetValue( ) ) )

            listDate = [ 
                str( parent.spnIDday.GetValue( ) ).zfill( 2 ),
                str( parent.chcIDmonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnIDyear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'iddate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'idaddr',
                str( parent.txtIDaddr.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'phonenum',
                str( parent.txtPhoneNum.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'email',
                str( parent.txtEmail.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'taxid',
                str( parent.txtTaxID.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'ogrnip',
                str( parent.txtOGRNIP.GetValue( ) ) )

        if ( tabIdx == 1
        and self.tabCommon.chcClientType.GetSelection( ) == 2 ):
            parent = self.tabDeclarant.pnlDeclarant
            sectionName = 'firm'
           
            caseDoc.SetTagValue( sectionName, 'firmtitle',
                str( parent.txtFirmTitle.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'firmtype',
                str( parent.cmbxFirmType.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'shorttype',
                str( parent.txtShortType.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'regaddr',
                str( parent.txtRegAddr.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'regorg',
                str( parent.txtRegOrg.GetValue( ) ) )

            listDate = [ 
                str( parent.spnRegDay.GetValue( ) ).zfill( 2 ),
                str( parent.chcRegMonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnRegYear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'regdate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'ogrn',
                str( parent.txtOGRN.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'taxid',
                str( parent.txtTaxID.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'kpp',
                str( parent.txtKPP.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'phonenum',
                str( parent.txtPhoneNum.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'email',
                str( parent.txtEmail.GetValue( ) ) )

        if ( tabIdx == 2
        and self.tabCommon.cbxNeedAgent.GetValue( ) ):
            parent =  self.tabAgent
            sectionName = 'agent'

            caseDoc.SetTagValue( sectionName, 'name1',
                str( parent.txtName1.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'name2',
                str( parent.txtName2.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'name3',
                str( parent.txtName3.GetValue( ) ) )

            listDate = [ 
                str( parent.spnBirthDay.GetValue( ) ).zfill( 2 ),
                str( parent.chcBirthMonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnBirthYear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'birthdate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'idtype',
                str( parent.cmbxIDtype.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idnum1',
                str( parent.txtIDnum1.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idnum2',
                str( parent.txtIDnum2.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idwho',
                str( parent.txtIDwho.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idcode1',
                str( parent.txtIDcode1.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'idcode2',
                str( parent.txtIDcode2.GetValue( ) ) )

            listDate = [ 
                str( parent.spnIDday.GetValue( ) ).zfill( 2 ),
                str( parent.chcIDmonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnIDyear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'iddate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'idaddr',
                str( parent.txtIDaddr.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'phonenum',
                str( parent.txtPhoneNum.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'email',
                str( parent.txtEmail.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'doctype',
                str( parent.cmbxDocType.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'docwho',
                str( parent.txtDocWho.GetValue( ) ) )

            listDate = [ 
                str( parent.spnDocDay.GetValue( ) ).zfill( 2 ),
                str( parent.chcDocMonth.GetSelection( ) ).zfill( 2 ),
                str( parent.spnDocYear.GetValue( ) ) ]
            caseDoc.SetTagValue( sectionName, 
                'docdate', '-'.join( listDate ) )

            caseDoc.SetTagValue( sectionName, 'docnum',
                str( parent.txtDocNum.GetValue( ) ) )

        if ( tabIdx == 3 ):
            parent =  self.tabArea
            sectionName = 'area'
            
            caseDoc.SetTagValue( sectionName, 'place',
                str( parent.cmbxPlace.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'addr',
                str( parent.txtAddr.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'rights',
                str( parent.chcRights.GetSelection( ) ) )
            caseDoc.SetTagValue( sectionName, 'size',
                str( parent.txtSize.GetValue( ) ) )
            
            caseDoc.SetTagValue( sectionName, 'kn11',
                str( parent.txtKN11.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'kn12',
                str( parent.txtKN12.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'kn13',
                str( parent.txtKN13.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'kn14',
                str( parent.txtKN14.GetValue( ) ) )

            caseDoc.SetTagValue( sectionName, 'kn21',
                str( parent.txtKN21.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'kn22',
                str( parent.txtKN22.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'kn23',
                str( parent.txtKN23.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'kn24',
                str( parent.txtKN24.GetValue( ) ) )

            caseDoc.SetTagValue( sectionName, 'category',
                str( parent.chcCategory.GetSelection( ) ) )
            caseDoc.SetTagValue( sectionName, 'target',
                str( parent.cmbxTarget.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'terzone',
                str( parent.chcTerZones.GetSelection( ) ) )
            
            selectedSpcZones = parent.lstbSpcZones.GetStrings( )
            listZones = []
            for zoneItm in selectedSpcZones:
                zoneIdx = zoneItm.split( ') ' )[ 0 ]
                listZones.append( str( zoneIdx ) )
            caseDoc.SetTagValue( sectionName, 
                'spczone', ','.join( listZones ) )
              
        if  ( tabIdx == 4 ):
            parent =  self.tabRules
            sectionName = 'rules'

            caseDoc.SetTagValue( sectionName, 'minarea',
                str( parent.txtMinArea.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'maxarea',
                str( parent.txtMaxArea.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'minwidth',
                str( parent.txtMinWidth.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'minoffset',
                str( parent.txtMinOffset.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'maxfloors',
                str( parent.txtMaxFloors.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'maxheight',
                str( parent.txtMaxHeight.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'ugnddensity',
                str( parent.spnUgndDensity.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'maxdensity',
                str( parent.spnMaxDensity.GetValue( ) ) )
            caseDoc.SetTagValue( sectionName, 'utilization',
                str( parent.txtUtilization.GetValue( ) ) )
        return    

    def OnClickbtnPrev( self, event ):
        event.Skip()
        idx = self.notebook.GetSelection( )
        while idx > 0:
            idx = idx - 1
            if self.notebook.GetEnabled( idx ): 
                break
        self.notebook.SetSelection( idx )       
        return

    def OnClickbtnNext( self, event ):
        event.Skip()
        idx = self.notebook.GetSelection()
        while idx < 4:
            idx = idx + 1
            if self.notebook.GetEnabled( idx ): 
                break
        self.notebook.SetSelection( idx )
        return

    def OnChnage_AnyCtrl( self, event ):
        event.Skip()
        self.GetParent().RefreshCtrlStates( )
        return

###########################################################################
## Class MenuBarMain
###########################################################################
class MenuBarMain ( wx.MenuBar ):

    def __init__( self ):
        super().__init__( )

        # Menu File
        menuFile = wx.Menu()
        self.itemNew = wx.MenuItem( 
            None, ID_NEW, u"Создать Документ", 
            wx.EmptyString, wx.ITEM_NORMAL )
        
        self.itemOpen = wx.MenuItem( 
            None, ID_OPEN, u"Открыть Документ", 
            wx.EmptyString, wx.ITEM_NORMAL )
        
        self.itemSave = wx.MenuItem( 
            None, ID_SAVE, u"Сохранить Документ", 
            wx.EmptyString, wx.ITEM_NORMAL )

        self.itemCklist = wx.MenuItem( 
            None, ID_CHK_LST, u"Сохранить Чеклист", 
            wx.EmptyString, wx.ITEM_NORMAL )

        self.itemClose = wx.MenuItem( 
            None, ID_CLOSE, u"Закрыть Документ", 
            wx.EmptyString, wx.ITEM_NORMAL )
        
        self.itemExit = wx.MenuItem( 
            None, ID_EXIT, u"Завершить работу", 
            wx.EmptyString, wx.ITEM_NORMAL )

        menuFile.Append( self.itemNew )
        menuFile.Append( self.itemOpen )
        menuFile.AppendSeparator()
        menuFile.Append( self.itemSave )
        menuFile.Append( self.itemCklist )
        menuFile.AppendSeparator()
        menuFile.Append( self.itemClose )
        menuFile.Append( self.itemExit )
        self.Append( menuFile, u"Файл" )

        # Menu Actions
        menuQueries = wx.Menu()
        class SubMenuStructure( object ):
            def __init__( self, title = 'submenu' ):
                self.rootMenu = wx.Menu( ) 
                self.rootItem = \
                    wx.MenuItem(
                        None, wx.ID_ANY, title, 
                        wx.EmptyString, wx.ITEM_NORMAL )
                self.items = {}
                return

        self.menuScheme = {}
        queriesData = modules.MenuQueries( ).elementTree.getroot( )
        for submenu in queriesData:
            title = submenu.attrib['title']
            self.menuScheme.update(
                { title : SubMenuStructure( title ) } )
           
            for menuelem in submenu:
                try:
                    itemID = int( menuelem.attrib[ 'id' ] )
                except Exception as e: print (e)
                self.menuScheme[ title ].items.update( 
                    { itemID :\
                        wx.MenuItem( 
                            None, wx.ID_ANY, menuelem.text, 
                            wx.EmptyString, wx.ITEM_NORMAL ) } )
                
                self.menuScheme[ title ].rootMenu.Append( 
                     self.menuScheme[ title ].items[ itemID ])
            
            self.menuScheme[ title ].rootItem.SetSubMenu( 
                self.menuScheme[ title ].rootMenu )
            menuQueries.Append( 
                self.menuScheme[ title ].rootItem )

        self.Append( menuQueries, u"Запросы" )

        # Menu Help
        menuHelp = wx.Menu()
        self.itemMan = wx.MenuItem( 
            menuHelp, ID_MANUAL, u"Инструкция", 
            wx.EmptyString, wx.ITEM_NORMAL )
        self.itemAbout = wx.MenuItem( 
            menuHelp, ID_ABOUT, u"О программе", 
            wx.EmptyString, wx.ITEM_NORMAL )

        menuHelp.Append( self.itemMan )
        menuHelp.Append( self.itemAbout )
        self.Append( menuHelp, u"Помощь" )
        return

###########################################################################
## Class FrameMain
###########################################################################
class FrameMain ( wx.Frame ):

    def __init__( self, parent ):
        super().__init__(
            parent = parent,
            title = u'СЭД "Управление земельными ресурсами"',
            size = ( 1024, 640 ),
            style = wx.DEFAULT_FRAME_STYLE )

        DefaultFont = wx.SystemSettings.GetFont( wx.SYS_DEFAULT_GUI_FONT )
        DefaultFont.SetPointSize(9)
        self.SetFont( DefaultFont )

        # Widgets
        self.menubar = MenuBarMain( )
        self.SetMenuBar( self.menubar )
        self.statusbar = wx.StatusBar( self )
        self.statusbar.SetFieldsCount( number = 6 )
        self.statusbar.SetStatusText( "Документ создан:", i = 1 )
        self.statusbar.SetStatusText( "Документ изменён:", i = 3 )
        self.statusbar.SetStatusWidths( 
            [ -10,-4,-2,-4,-2,-1 ] )
        self.SetStatusBar( self.statusbar )
        self.toolbar = wx.ToolBar( self )
        self.SetToolBar( self.toolbar )

        self.vsizer = wx.BoxSizer( wx.VERTICAL )
        self.NewCaseDoc( )
        self.SetSizer( self.vsizer )
        self.Centre( wx.BOTH )
        
        # Events Menu
        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_New, 
            id = self.menubar.itemNew.GetId() )
        
        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_Open, 
            id = self.menubar.itemOpen.GetId() )
        
        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_Save, 
            id = self.menubar.itemSave.GetId() )

        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_ChkList, 
            id = self.menubar.itemCklist.GetId() )

        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_Close, 
            id = self.menubar.itemClose.GetId() )
        
        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_Exit, 
            id = self.menubar.itemExit.GetId() )
        
        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_Man, 
            id = self.menubar.itemMan.GetId() )
        
        self.menubar.Bind( 
            wx.EVT_MENU, self.OnSelect_MenuItem_About, 
            id = self.menubar.itemAbout.GetId() )       

        for submenu in self.menubar.menuScheme:
            for itemID in self.menubar.menuScheme[ submenu ].items:
                item = self.menubar.menuScheme[ submenu ].items[ itemID ]
                itemName = item.GetItemLabelText( )
                itemName = itemName.replace(' ','_')
                self.menubar.Bind( 
                    wx.EVT_MENU, 
                    lambda event, 
                    queryID = itemID, 
                    queryName = itemName :\
                    self.OnSelect_MenuItem_AnyQuery( 
                        event, queryID, queryName ), 
                    id = item.GetId() )
        return

    
    def RefreshMainMenu( self ):
        try: 
            if self.panelDoc: pass
        except Exception as e: 
            print ( e )
            self.menubar.itemClose.Enable( False )
            self.menubar.itemSave.Enable( False )
            self.menubar.itemCklist.Enable( False )           
            
            for submenu in self.menubar.menuScheme:
                for itemID in self.menubar.menuScheme[ submenu ].items:
                    item = self.menubar.menuScheme[ submenu ].items[ itemID ]
                    item.Enable( False )
        else:         
            self.menubar.itemClose.Enable( True )
            self.menubar.itemSave.Enable( True )
            self.menubar.itemCklist.Enable( True )
            
            for submenu in self.menubar.menuScheme:
                for itemID in self.menubar.menuScheme[ submenu ].items:
                    item = self.menubar.menuScheme[ submenu ].items[ itemID ]
                    item.Enable( True )
        return

    def RefreshStatusBar( self ):
        try: 
            if self.panelDoc: pass
        except Exception as e: 
            print ( e )
            self.statusbar.SetStatusText( "", i = 0 )
            self.statusbar.SetStatusText( "", i = 2 )
            self.statusbar.SetStatusText( "", i = 4 )
            self.statusbar.SetStatusText( "", i = 5 )
        else:
            self.statusbar.SetStatusText(                     
                pathlib.Path( self.panelDoc.fileName ).name, i = 0 )
            caseDoc = self.panelDoc.caseDocument
            status = caseDoc.GetTagValue( 'case','created')
            self.statusbar.SetStatusText( status, i = 2 )
            status = caseDoc.GetTagValue( 'case','changed')
            self.statusbar.SetStatusText( status, i = 4 )
            self.statusbar.SetStatusText(                     
                str( self.panelDoc.docChanges ), i = 5 )            
        return

    def RefreshCtrlStates( self, resetChanges = False ):
        try: 
            if self.panelDoc: pass
        except Exception as e: print (e)
        else:
            if resetChanges: self.panelDoc.docChanges = 0
            else: self.panelDoc.docChanges += 1
        self.RefreshStatusBar( )
        self.RefreshMainMenu( )
        return    

    def CloseCaseDoc( self ):
        try: 
            if self.panelDoc: pass
        except Exception as e: 
            print (e)
            return True
        else:
            caseDoc = self.panelDoc.caseDocument
            if ( self.panelDoc.docChanges > 0 ):
                if ( wx.MessageBox( 
                        "Изменения не были сохранены!\n\
                        Закрыть без сохранения?", "Требуется подтверждение!",
                        wx.ICON_QUESTION | wx.YES_NO, self) == wx.NO ):
                    return False
            try:
                self.vsizer.Hide( self.panelDoc )
                del(self.panelDoc)
                self.RefreshCtrlStates( resetChanges = True )
                self.vsizer.Layout( )
            except Exception as e: print (e)
        return True

    def NewCaseDoc( self ):
        if self.CloseCaseDoc( ):
            self.panelDoc = PanelDoc( self )
            self.vsizer.Add(
                self.panelDoc, 1, 
                flag=wx.ALL|wx.EXPAND, 
                border=0 )

            today = datetime.datetime.today().strftime('%d-%m-%Y')
            caseDoc = self.panelDoc.caseDocument
            caseDoc.SetTagValue( 'case', 'created', today )
            caseDoc.SetTagValue( 'case', 'changed', today )
            self.RefreshCtrlStates( resetChanges = True )
            self.vsizer.Layout( )
            return True
        return False

    def OpenCaseDoc( self, fileName ):
        if self.NewCaseDoc( ):
            if pathlib.Path( fileName ).exists( ):
                self.panelDoc.fileName = \
                    pathlib.Path( fileName )
                caseDoc = self.panelDoc.caseDocument
                caseDoc.Load( srcPath = fileName )
        return

    def SaveCaseDoc( self ):
        try:
            if self.panelDoc: pass
        except Exception as e: print (e)
        else:          
            self.panelDoc.caseDocument.Save( 
                self.panelDoc.fileName)
        return

    # Events
    def OnSelect_MenuItem_Close( self, event ):
        if self.CloseCaseDoc( ):
            pass
        event.Skip()
        return

    def OnSelect_MenuItem_New( self, event ):
        if self.NewCaseDoc( ):
            pass
        event.Skip()
        return

    def OnSelect_MenuItem_Open( self, event ):       
        event.Skip()
        try: 
            self.panelDoc.GetAllTabs( )
        except Exception as e: print (e)

        wildcard="Структурированные данные (*.csd)|*.csd"
        with wx.FileDialog(
            self, "Загрузить документ из файла", 
            wildcard=wildcard,
            style = wx.FD_OPEN|\
                    wx.FD_CHANGE_DIR|\
                    wx.FD_FILE_MUST_EXIST ) as fileDialog:

            fileDialog.SetSize((500,400))
            fileDialog.Centre( )
            
            try:
                if ( self.panelDoc.fileName.exists( ) ):
                    fileDialog.SetDirectory( 
                        str (self.panelDoc.fileName.parent ) )
            except Exception as e: print (e)
            else: fileDialog.SetDirectory( 
                    str( pathlib.Path.home( ) ) )
            
            if ( fileDialog.ShowModal( ) == wx.ID_CANCEL ):
                return
            fileName = fileDialog.GetPath( )
        if not fileName.endswith( '.csd' ):
            fileName = f'{fileName}.csd'
        
        self.OpenCaseDoc( fileName )
        self.panelDoc.SetAllTabs( )
        self.RefreshCtrlStates( 
            resetChanges = True )        
        return

    def OnSelect_MenuItem_Save( self, event ):
        
        event.Skip()        
        caseDoc = self.panelDoc.caseDocument
        self.panelDoc.GetAllTabs( )
        
        today = datetime.datetime.today().strftime('%d-%m-%Y')
        caseDoc.SetTagValue( 'case', 'changed', today )

        skpNum = caseDoc.GetTagValue( 'common', 'sinkopa' )
        skpNum = skpNum.replace( '/','-' )
        if ( skpNum == '' ):
            wx.MessageBox( 
                "Не указан номер в СИНКОПе!", 
                "Ошибка!", wx.OK, self )
            return False

        caseNum = caseDoc.GetTagValue( 'common', 'regnum' )
        if ( caseNum == '' ):
            wx.MessageBox( 
                "Не указан номер регистрации в журнале!", 
                "Ошибка!", wx.OK, self )
            return False
        
        if not self.panelDoc.fileName.exists( ):
            dirPath = self.SelectDir( )
            try: 
                if not dirPath.exists( ): 
                    return False
            except Exception as e: 
                    print (e)
                    return False                    
            self.panelDoc.fileName = \
                dirPath.joinpath( 
                    f'Дело{caseNum}_{skpNum}.csd' )

        self.SaveCaseDoc( )
        self.panelDoc.SetAllTabs( )
        self.RefreshCtrlStates( 
            resetChanges = True )
        return True

    def SelectDir( self ):
        # ask the user where new file to save
        with wx.DirDialog( 
            self, "Сохранить документ в каталог", 
            str( pathlib.Path.home( ) ),
            style = wx.DD_DEFAULT_STYLE |\
                    wx.DD_DIR_MUST_EXIST|\
                    wx.DD_CHANGE_DIR ) as dirDialog:
        
            dirDialog.SetSize((500,400))
            dirDialog.Centre()
            if ( dirDialog.ShowModal( ) == wx.ID_CANCEL ):
                return None
            dirName = dirDialog.GetPath( )
        return pathlib.Path( dirName )

    def OnSelect_MenuItem_Exit( self, event ):
        self.CloseCaseDoc()
        self.Close()
        self.Destroy()
        return

    def OnSelect_MenuItem_Man( self, event ):
        pass
        event.Skip()
        return

    def OnSelect_MenuItem_About( self, event ):
        FrameAbout( self ).Show(True)
        event.Skip()
        return

    def  OnSelect_MenuItem_AnyQuery( self, event, queryID, queryName ):
        
        # Ask about save changes
        if ( wx.MessageBox( 
                u"Требуется сохранить изменения!\n\
                Желаете продолжить?", u"Требуется подтверждение!",
                wx.ICON_QUESTION | wx.YES_NO, self) == wx.NO ):
            event.Skip()
            return
        
        if not ( self.OnSelect_MenuItem_Save( event ) ):
            event.Skip()
            return

        caseDoc = self.panelDoc.caseDocument
        caseNum = caseDoc.GetTagValue( 'common', 'regnum' )
        skpNum = caseDoc.GetTagValue( 'common', 'sinkopa' )
        skpNum = skpNum.replace( '/','-' )
        tgtName = f'Дело{caseNum}_{skpNum}_{queryName}.docx'
        tgtPath = self.panelDoc.fileName.parent
        tgtPath = tgtPath.joinpath( tgtName ).resolve( )     
        #print( f'queryID={queryID}; targetPath={tgtPath}' )
        self.queryDoc = modules.DOCXQuery( 
            queryID, tgtPath, caseDoc )
        event.Skip()
        return
      
    def OnSelect_MenuItem_ChkList( self, event ):
        event.Skip()
        # Ask about save changes     
        self.OnSelect_MenuItem_AnyQuery( 
                event = event,
                queryID = '1000', 
                queryName = 'Чеклист' )
        return

###########################################################################
## Class FrameSpcZones
###########################################################################
class FrameSpcZones ( wx.Dialog ):

    def __init__( self, parent ):
        super().__init__(
            parent= parent,
            title = u"Зоны с особыми условиями использования территорий",
            size = ( 800,520 ),
            style = wx.DEFAULT_DIALOG_STYLE|wx.STAY_ON_TOP )

        DefaultFont = wx.SystemSettings.GetFont( 
            wx.SYS_DEFAULT_GUI_FONT )
        DefaultFont.SetPointSize(9)
        self.SetFont( DefaultFont )

        dataSource = modules.DataSource()
        listAreaSpcZones = dataSource.RetrieveData(
            'area','spczone', 'alias', True )

        self.lstbSpcZones = wx.ListBox( 
            self, id = wx.ID_ANY, pos = wx.DefaultPosition, 
            size = ( -1,440 ), style = wx.LB_ALWAYS_SB|wx.LB_MULTIPLE,
            choices = listAreaSpcZones )
        
        gsizer = wx.BoxSizer( wx.HORIZONTAL )
        gsizer.Add(
            self.lstbSpcZones, proportion = 1, 
            border=5, flag=wx.ALL )

        self.btnOK = wx.Button( self, wx.ID_OK )
        bsizer = wx.BoxSizer( wx.HORIZONTAL )
        bsizer.Add( 
            self.btnOK, proportion = 0, 
            border=5, flag=wx.ALL )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add(
            gsizer, proportion = 1, border=0, 
            flag=wx.ALIGN_CENTER_HORIZONTAL|wx.ALL )
        vsizer.Add(
            bsizer, proportion = 0, border=0, 
            flag=wx.ALIGN_CENTER_HORIZONTAL|wx.ALL )
        self.SetSizer( vsizer )
        #vsizer.SetSizeHints( self )
        vsizer.Layout( )
        self.Centre( )
        return

###########################################################################
## Class FrameAbout
###########################################################################

class FrameAbout ( wx.Dialog ):

    def __init__( self, parent ):
        super().__init__(
            parent = parent,
            title = u"О программе",
            size = ( 400,350 ),
            style = wx.DEFAULT_DIALOG_STYLE|wx.STAY_ON_TOP )

        DefaultFont = wx.SystemSettings.GetFont( wx.SYS_DEFAULT_GUI_FONT )
        DefaultFont.SetPointSize(9)
        self.SetFont( DefaultFont )
       
        logoPath = pathlib.Path( __file__ )
        logoPath = logoPath.resolve( ).parent.parent
        logoPath = logoPath.joinpath( u'resources' )
        logoPath = logoPath.joinpath( u'logo.jpg' )
        
        bmpLogo = wx.StaticBitmap(
            self, wx.ID_ANY,
            wx.Bitmap( str(logoPath), wx.BITMAP_TYPE_ANY ),
            wx.DefaultPosition, ( 96,96 ), 0 )

        logsizer = wx.BoxSizer( wx.HORIZONTAL )
        logsizer.Add(
            bmpLogo, 1, 
            flag=wx.ALIGN_CENTER_VERTICAL|wx.ALL, border=5 )

        txtLcns = wx.TextCtrl(
            self, wx.ID_ANY,
            ('Система электронного документооборота\n'
           '"Управление Земельными Ресурсами"\n'
           'Версия 0.9-beta.1 (11.2021)\n\n'
            
            'Программа является свободным программным\n'
            'обеспечением: вы можете распространять ее\n' 
            'и/или изменять в соответствии с условиями\n'
            'Общей публичной лицензии GNU, опубликованной\n'
            'Фондом свободного программного обеспечения,\n' 
            'версии 3 Лицензии, либо (по вашему выбору)\n' 
            'любой более поздней версии.\n\n'
            
            'Программа поставляется по принципу\n'
            '"AS IS" ("как есть"). Никакие гарантии\n'
            'не прилагаются и не предусматриваются.\n\n'
            
            'Сайт проекта:\n' 
            'https://github.com/DenisKazakoff/land_admin_ecm\n'
            'Автор: Денис Казаков (denis.kazakoff@gmail.com)'),
            wx.DefaultPosition, ( 390,-1 ), wx.TE_READONLY|wx.TE_MULTILINE|wx.ALIGN_CENTRE_HORIZONTAL )
        
        licsizer = wx.BoxSizer( wx.HORIZONTAL )
        licsizer.Add(txtLcns, 1, flag=wx.EXPAND|wx.ALL, border=5 )

        btnOK = wx.Button( self, wx.ID_OK )
        btnsizer = wx.BoxSizer( wx.HORIZONTAL )
        btnsizer.Add( btnOK, 1, flag=wx.ALIGN_CENTER_VERTICAL|wx.ALL, border=5 )

        vsizer = wx.BoxSizer( wx.VERTICAL )
        vsizer.Add( logsizer, 0, flag=wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, border=0 )
        vsizer.Add( licsizer, 1, flag=wx.EXPAND|wx.ALL, border=0 )
        vsizer.Add( btnsizer, 0, flag=wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, border=0 )
        self.SetSizer( vsizer )
        #vsizer.SetSizeHints( self )
        vsizer.Layout( )        
        self.Centre( )
        return
