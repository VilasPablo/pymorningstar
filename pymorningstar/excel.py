import time
import os
import datetime
import xlwings as xw
import pandas as pd
import pyautogui as m

from datetime import date, timedelta
from dateutil.relativedelta import relativedelta


PATH = os.path.dirname(os.path.abspath(__file__))

class ExcelMorning():
    
    def __init__(self):
        
        # Store Info about the Attributes Download
        self.attr_info = pd.DataFrame(columns =['Index Len', 'Col Len', 'NaN', 'Size','Time','Formula'])
        self.attr_info.index.name ='serie_code'
        # Store Info about the Attributes Download
        self.hold_info= pd.DataFrame(columns =['SerieCode', 'StarDown', 'EndDown', 'StarDate','EndDate',
                                'IndexLen', 'ColLen', 'NaN', 'Size','Time','Formula']).set_index(  
                                ['SerieCode', 'StarDown', 'EndDown'])
        
        self.initialize_morning()
    
    def initialize_morning(self):
        self.wb = xw.Book()
        time.sleep(0.5)
        self.wb.app.api.WindowState = xw.constants.WindowState.xlMaximized
        self.wb.app.activate(steal_focus=True)
        time.sleep(5)
        self.sheet = self.wb.sheets[0]
         
             
    def get_holding (self, isin_fund, inception_date, obsolete_date, asset_id, holding_type='ALL',data_type = 'WEIGHT', 
                     frequency = 'A', show_holding_type = True, show_country=False, 
                     show_currency=False, show_maturity=False, show_coupon=False, wait_time=0, months_frac=12):
                        # asset_id = SecID / Ticker / ISIN /CUSIP
                        # holding_type = STOCKS / BONDS / FUNDS / ETFS / ALL
                        # data_type = WEIGHT / MV / SHARES 
                        # frequency =  A / D / M / Q / Y /             
        df_holding = pd.DataFrame() # df to store the holdings
        df_index=['Name', asset_id]
        inception_date = inception_date.date() 
        obsolete_date = obsolete_date.date()
        
        # Write part of the excel formula about holdings
        formula_1 = ',"CORR=C, ASCENDING=TRUE,' + 'HT='+ holding_type + ',' + data_type + '=TRUE,FREQ=' + frequency +  ',NAME =TRUE'
        for show_1, show_2, show_3 in zip([show_holding_type, show_country, show_currency, show_maturity, show_coupon],
                                  [',SHOWHT=TRUE', ',SHOWCOUNTRY=TRUE', ',SHOWCURRENCY=TRUE',
                                   ',SHOWMATURITYDATE=TRUE', ',SHOWCOUPON=TRUE'], 
                                  ['Detail Holding Type', 'Country', 'Currency', 'Maturity Date','Coupon %']):
            if show_1 == True:
                formula_1 +=  show_2
                df_index.append(show_3)           
        
        # Get the holdings in yearly basis 
        start_date = inception_date.replace(day=1)
        end_date = start_date
        while end_date < obsolete_date:
            end_date = start_date + relativedelta(months=months_frac) - timedelta(days=1)
            if end_date > obsolete_date:
                end_date = obsolete_date.replace(day=1) + relativedelta(months=1) - timedelta(days=1)            
            start_date = start_date.strftime("%d/%m/%Y")
            end_date = end_date.strftime("%d/%m/%Y")
            # Create the formula
            formula = ",".join(['"'+word+ '"' for word in [isin_fund, asset_id, start_date, end_date] ]) 
            formula = '=MSHOLDING(' + formula + formula_1 +'")' 
            
            # introduce formula and get the data
            df = self.get_data(formula, wait_time=wait_time)
            # Info about the download of holding
            start_date = datetime.datetime.strptime(start_date,"%d/%m/%Y").date()
            end_date = datetime.datetime.strptime(end_date,"%d/%m/%Y").date()  
        
            
            self.hold_info.loc[(isin_fund, start_date, end_date),:] = [inception_date, obsolete_date,
                                len(df.index),len(df.columns),df.isnull().sum().sum(), df.size, datetime.datetime.now(), "'"+formula.replace('""','"')]  
            # adjust date for the next loop
            start_date = end_date + timedelta(days=1) 
            
            if df.empty==True:
                continue
            # Clean holdings data
            df = df.reset_index()
            df = df.set_index(df_index).stack()
            df = df.to_frame()
            df.index.rename(names='Date',level=len(df.index.levels)-1, inplace=True)
            df.columns = [data_type]
            df['SerieCode'] = isin_fund
            df = df.reset_index().set_index(['SerieCode','Date'])
            df_holding= pd.concat([df_holding, df.sort_index()], axis= 0)        
            
        return df_holding
    
    # the first attribute can not be empty
    def get_attributes(self,serie_code, variables_name, star_date, end_date, frequency, days='C', wait_time=0):      
                            # frequency == monthly(M) -- quarterly(Q) -- semiannually (S) -- yearly (Y)
                            # days == trading/activity days(T) -- calendar days(C)!! -- weekdays (W)
                            # fill == Last available data (C) -- previous day's data(P) --  Zero (T)
        formula = '=MSTS("' + serie_code + '",' + ',"&'.join('"'+x for x in variables_name)+'","' +date(
        *star_date).strftime("%d/%m/%Y") + '","' + date(*end_date).strftime("%d/%m/%Y") + '",'+ (
        '"CORR=C, DATES=True, ASCENDING=TRUE, FILL=B, HEADERS=TRUE, FREQ=%s, DAYS=%s")' % (frequency, days))
        
        df = self.get_data(formula, variables_name, wait_time)# introduce formula and get the data
        
        self.attr_info.loc[serie_code,:] = [len(df.index), len(df.columns), df.isnull().sum().sum(), 
                                df.size, datetime.datetime.now(), "'"+formula.replace('""','"')]
        # Clean & prepare attributes data
        df['serie_code'] = serie_code
        df.index.name='Date'
        df = df.reset_index().set_index('serie_code')
        return df

    def get_data(self, formula, variables_name=None, wait_time=0):

        while True:
            self.wb.app.api.Cells.Clear()# remove the old values
            time.sleep(0.25)
            self.sheet.range('B1').value= variables_name
            self.sheet.range('A1').value = formula
            time.sleep(1+wait_time)
            self.wait_processing()# wait until data is download
            if self.is_limit() == False: # if limit is true we wait one day and we introduce formula again
                break
        if variables_name != None: #holding vs attributes
            self.sheet.range('B1').value= variables_name
        df = self.sheet.range(self.sheet.range('A1').expand(), self.sheet.range('B1').expand()).options(pd.DataFrame).value
        time.sleep(1.5)
        self.wb.app.api.Cells.Clear()# remove the old values
        time.sleep(1)
        return df
    
    def wait_processing(self):
        df = self.sheet.range(self.sheet.range('A1'), self.sheet.range('B1').expand()).options(pd.DataFrame).value
        while df.index.name=='Processing...':
            time.sleep(0.5)
            df = self.sheet.range(self.sheet.range('A1'), self.sheet.range('B1').expand()).options(pd.DataFrame).value 
        time.sleep(0.75)


    def is_limit(self):
        if m.locateOnScreen(os.path.join(PATH, 'morning_limit.png'), confidence=0.85)!=None:
            input('IS MORNING LIMIT')
            self.initialize_morning()
            return True 
        else:
            return False
