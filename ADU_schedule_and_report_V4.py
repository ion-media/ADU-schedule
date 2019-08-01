import pandas as pd
import numpy as np
import time
from dateutil.parser import parse
import math
import datetime as dt
from datetime import datetime
from collections import defaultdict
import xlsxwriter
from pyxlsb import open_workbook as open_xlsb
import zipfile
import os
import shutil
from openpyxl import load_workbook
import csv
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors


DIR_INPUT='C:/ION/Commercial/ADU_Report/V2/'
DIR_OUTPUT='C:/ION/Commercial/ADU_Report/V2/'
P = set(['Holiday Movies (Prime)', 'ION Originals (Prime)', 'Prime', 'Prime no CM'])
NP = set(['Daytime (M-F)', 'Early Morning (M-S)', 'Fringe (M-S)', 'Holiday Movies (Non Prime)', \
          'Late Night (M-S)', 'Morning (M-S)', 'Non-Prime ROS**', 'Non-Prime ROS*', 'Weekend Day (S-Sun)'])

# Compare dates
def date_comparison(date1, date2):
    date1 = parse(str(date1))
    date2 = parse(str(date2))
    return date1 < date2

# Number of weeks between two dates. 0619-0625 has 0 week in between, 0619-0626 has 1 week in between
def weeks_between(d1, d2):
    d1 = parse(str(d1))
    d2 = parse(str(d2))
    return math.trunc((d2 - d1).days / 7)


# transform different selling titles to P and NP
def dayparts(r):
    global P
    global NP
    if r['Selling Title'] in P:
        return 'P'
    if r['Selling Title'] in NP:
        return 'NP'
    return None


class GID:
    def __init__(self, r):
        self.row = r
        self.DealName = set()
        self.DealNum = set()
        self.AEName = set()
        self.Agency = set()
        self.Marketplace = set()
        dictionary = {'Booked $': 0, 'Deal Imp': 0, 'Delv Imp': 0, 'Imps Owed': 0, 'Units': 0, 'CPM': 0}
        self.Sold_P = dictionary.copy()
        self.Sold_NP = dictionary.copy()
        self.ADU_P = dictionary.copy()
        self.ADU_NP = dictionary.copy()
        self.Total = dictionary.copy()
        self.P = {'Guar': 0, 'Est': 0, 'Delv': 0, 'Q2 Imp': 0, 'ADUs': 0}
        self.NP = self.P.copy()

    def new_info(self, ratings):
        if len(str(self.row['Guarantee Name']))<4:
            self.GName = self.row['Deal Name']
        else:
            self.GName = self.row['Guarantee Name']
        self.DealNum.add(self.row['Deal Numbers in Guarantee'])
        self.Marketplace.add(self.row['Marketplace'])
        self.Advertiser = self.row['Advertiser']
        self.AEName.add(self.row['AE Name'])
        self.Agency.add(self.row['Agency Name (Billing)'])
        self.DealName.add(self.row['Deal Name'])
        self.SoldDemo = self.row['Primary Demo']
        self.StartDate = self.row['Week Start Date']
        self.EndDate = self.row['Week End Date']

        # forecast ratings
        try:
            self.P['Q2 Imp'] = ratings.loc[ratings['Demo'] == self.SoldDemo, 'Prime Imp'].iloc[0]
            self.NP['Q2 Imp'] = ratings.loc[ratings['Demo'] == self.SoldDemo, 'Non Prime Imp'].iloc[0]
        except:
            print(self.SoldDemo)

        l = [self.GName, self.DealNum, self.Marketplace, \
             self.Advertiser, 
             self.AEName, self.Agency, self.DealName, self.SoldDemo, self.StartDate, self.EndDate] \
            + self.type_by_daypart(self.row)
        return l

    def update_info(self, r):
        self.DealName.add(r['Deal Name'])
        self.Marketplace.add(r['Marketplace'])
        self.AEName.add(r['AE Name'])
        self.Agency.add(r['Agency Name (Billing)'])

        if not date_comparison(self.StartDate, r['Week Start Date']):
            self.StartDate = r['Week Start Date']
        if date_comparison(self.EndDate, r['Week End Date']):
            self.EndDate = r['Week End Date']

        l = [self.GName, self.DealNum, self.Marketplace,\
             self.Advertiser, 
             self.AEName, self.Agency, self.DealName, self.SoldDemo, self.StartDate, self.EndDate] \
            + self.type_by_daypart(r)
        return l

    def type_by_daypart(self, r):
        if r['ADU Ind'] == 'N':
            if dayparts(r) == 'P':
                self.Sold_P['Booked $'] += r['Booked Dollars']
                self.Sold_P['Deal Imp'] += r['Primary Demo Non-ADU Equiv Deal Imp'] / 1000
                self.Sold_P['Delv Imp'] += r['Primary Demo Equiv Post Imp'] / 1000
                self.Sold_P['Imps Owed'] = self.Sold_P['Deal Imp'] - self.Sold_P['Delv Imp']
                self.Sold_P['Units'] += r['Equiv Units']
                self.Sold_P['CPM'] = self.Sold_P['Booked $'] / self.Sold_P['Deal Imp'] if self.Sold_P['Deal Imp'] else 0
            elif dayparts(r) == 'NP':
                self.Sold_NP['Booked $'] += r['Booked Dollars']
                self.Sold_NP['Deal Imp'] += r['Primary Demo Non-ADU Equiv Deal Imp'] / 1000
                self.Sold_NP['Delv Imp'] += r['Primary Demo Equiv Post Imp'] / 1000
                self.Sold_NP['Imps Owed'] = self.Sold_NP['Deal Imp'] - self.Sold_NP['Delv Imp']
                self.Sold_NP['Units'] += r['Equiv Units']
                self.Sold_NP['CPM'] = self.Sold_NP['Booked $'] / self.Sold_NP['Deal Imp'] if self.Sold_NP['Deal Imp'] else 0

        else:
            if dayparts(r) == 'P':
                self.ADU_P['Delv Imp'] += r['Primary Demo Equiv Post Imp'] / 1000
                self.ADU_P['Imps Owed'] = self.ADU_P['Deal Imp'] - self.ADU_P['Delv Imp']
                self.ADU_P['Units'] += r['Equiv Units']

            elif dayparts(r) == 'NP':
                self.ADU_NP['Delv Imp'] += r['Primary Demo Equiv Post Imp'] / 1000
                self.ADU_NP['Imps Owed'] = self.ADU_NP['Deal Imp'] - self.ADU_NP['Delv Imp']
                self.ADU_NP['Units'] += r['Equiv Units']

        self.Total['Booked $'] = self.Sold_P['Booked $'] + self.Sold_NP['Booked $']
        self.Total['Deal Imp'] = self.Sold_P['Deal Imp'] + self.Sold_NP['Deal Imp']
        self.Total['Delv Imp'] = self.Sold_P['Delv Imp'] + self.Sold_NP['Delv Imp'] \
                                 + self.ADU_P['Delv Imp'] + self.ADU_NP['Delv Imp']
        self.Total['% Delv'] = self.Total['Delv Imp'] / self.Total['Deal Imp'] if self.Total['Deal Imp'] else 0
        self.Total['Imps Owed'] = self.Total['Deal Imp'] - self.Total['Delv Imp']
        self.Total['Units'] = self.Sold_P['Units'] + self.Sold_NP['Units'] + self.ADU_P['Units'] + self.ADU_NP['Units']
        self.Total['CPM'] = self.Total['Booked $'] / self.Total['Deal Imp'] if self.Total['Deal Imp'] else 0
        self.Total['Liability $'] = max(0, self.Total['Imps Owed'] * self.Total['CPM'])
        self.Total['P Mix %'] = self.Sold_P['Deal Imp'] / self.Total['Deal Imp'] if self.Total['Deal Imp'] else 0
        self.Total['NP Mix %'] = 1 - self.Total['P Mix %']

        self.P['Guar'] = self.Sold_P['Deal Imp'] / self.Sold_P['Units'] if self.Sold_P['Units'] else 0
        self.P['ADUs'] = round(self.Total['P Mix %'] * self.Total['Imps Owed'] / self.P['Q2 Imp'])
        self.P['Est'] = self.Sold_P['Delv Imp'] / self.Sold_P['Units'] if self.Sold_P['Units'] else 0
        self.P['Delv'] = self.P['Est'] / self.P['Guar'] if self.P['Guar'] else 0

        self.NP['Guar'] = self.Sold_NP['Deal Imp'] / self.Sold_NP['Units'] if self.Sold_NP['Units'] else 0
        self.NP['ADUs'] = round(self.Total['NP Mix %'] * self.Total['Imps Owed'] / self.NP['Q2 Imp'])
        self.NP['Est'] = self.Sold_NP['Delv Imp'] / self.Sold_NP['Units'] if self.Sold_NP['Units'] else 0
        self.NP['Delv'] = self.NP['Est'] / self.NP['Guar'] if self.NP['Guar'] else 0
        self.Total['ADUs'] = self.P['ADUs'] + self.NP['ADUs']

        return [self.Sold_P, self.Sold_NP, self.ADU_P, self.ADU_NP, self.Total, self.P, self.NP]


def get_dict(df, ratings, endq):
    dic = dict()
    for col, row in df.iterrows():
        if date_comparison(row['Week Start Date'], endq):
            # if the Guarantee ID has not shown before, get new info
            if row['Guarantee ID'] not in dic: 
                c = GID(row)
                dic[row['Guarantee ID']] = [c, c.new_info(ratings)]
            # else update the information
            else: 
                dic[row['Guarantee ID']][1] = dic[row['Guarantee ID']][0].update_info(row)
        else:
            continue
    return dic


# read in the result from get_dict, turn dictionary to dataframe
def form_df(result):
    column_names = ['Guarantee ID', 'Guarantee Name', 'Deal ID', 'Marketplace',\
                    'Advertiser', \
                    'AE Name', 'Agency', 'Deal Name', 'Primary Demo', 'Sold Start Date', 'Sold End Date', \
                    'Sold Prime Booked $', 'Sold Prime Deal Imp', 'Sold Prime Delv Imp', 'Sold Prime Imps Owed',
                    'Sold Prime Units', \
                    'Sold Prime CPM', 'Sold NP Booked $', 'Sold NP Deal Imp', 'Sold NP Delv Imp', 'Sold NP Imps Owed', \
                    'Sold NP Units', 'Sold NP CPM', 'ADU Prime Booked $', 'ADU Prime Deal Imp', 'ADU Prime Delv Imp', \
                    'ADU Prime Imps Owed', 'ADU Prime Units', 'ADU Prime CPM', 'ADU NP Booked $', 'ADU NP Deal Imp', \
                    'ADU NP Delv Imp', 'ADU NP Imps Owed', 'ADU NP Units', 'ADU NP CPM', 'Total Booked $', \
                    'Total Deal Imp', 'Total Delv Imp', 'Total Imps Owed', 'Total Units', 'Total CPM', \
                    'Total % Delv', 'Total Liability $', 'Total P Mix %', 'Total NP Mix %', 'Total ADUs', \
                    'P Guar', 'P Est', 'P Delv', 'P Q2 Imp', 'P ADUs', 'NP Guar', 'NP Est', 'NP Delv', 'NP Q2 Imp', \
                    'NP ADUs']
    rows = []
    for k, v in result.items():
        row = []
        row.append(k)  # G_ID

        for element in v[1]:
            if type(element) is set:
                element = list(element)
                row.append(','.join(str(e) for e in element))
            elif type(element) is dict:
                if len(element) == 6:
                    for key in ['Booked $', 'Deal Imp', 'Delv Imp', 'Imps Owed', 'Units', 'CPM']:
                        row.append(element[key])
                elif len(element) ==5:
                    for key in ['Guar', 'Est', 'Delv', 'Q2 Imp', 'ADUs']:
                        row.append(element[key])
                else:
                    for key in ['Booked $', 'Deal Imp', 'Delv Imp', 'Imps Owed', 'Units', 'CPM', \
                        '% Delv', 'Liability $', 'P Mix %', 'NP Mix %', 'ADUs']:
                        row.append(element[key])
            else:
                row.append(element)
        
        rows.append(row)
    output = pd.DataFrame(rows)
    output.columns = column_names
    output = output[['Guarantee ID', 'Guarantee Name', 'Marketplace', \
                     'Advertiser', \
                     'AE Name', 'Agency', 'Deal Name', 'Deal ID', 'Primary Demo', 'Sold Start Date', 'Sold End Date', \
                     'Sold Prime Booked $', 'Sold Prime Deal Imp', 'Sold Prime Delv Imp', 'Sold Prime Imps Owed',
                     'Sold Prime Units', \
                     'Sold Prime CPM', 'Sold NP Booked $', 'Sold NP Deal Imp', 'Sold NP Delv Imp', 'Sold NP Imps Owed', \
                     'Sold NP Units', 'Sold NP CPM', 'ADU Prime Booked $', 'ADU Prime Deal Imp', 'ADU Prime Delv Imp', \
                     'ADU Prime Imps Owed', 'ADU Prime Units', 'ADU Prime CPM', 'ADU NP Booked $', 'ADU NP Deal Imp', \
                     'ADU NP Delv Imp', 'ADU NP Imps Owed', 'ADU NP Units', 'ADU NP CPM', 'Total Booked $', \
                     'Total Deal Imp', 'Total Delv Imp', 'Total % Delv', 'Total Imps Owed', 'Total Units', 'Total CPM', \
                     'Total Liability $', 'Total P Mix %', 'Total NP Mix %', 'P Guar', \
                     'NP Guar', 'P Est', 'NP Est', 'P Delv', 'NP Delv', 'P Q2 Imp', 'NP Q2 Imp', 'P ADUs', 'NP ADUs', \
                     'Total ADUs']]
    return output


# Mondays between two dates, excluding the last Monday. Used for scheduling ADU
def week_range(startdate, enddate):
    weeks = pd.date_range(startdate, enddate, freq='W-MON').strftime('%m/%d/%Y').tolist()
    return weeks[:-1]


# Find the previous, current, next and the next after next quarter giving the input date
def find_quarters(quarters, startdate):
    mon = pd.date_range(startdate, periods=1, freq='W-MON').strftime('%m/%d/%Y').tolist()[0]
    current_q = quarters.loc[quarters['start_date'].astype(str) == mon, 'quarter'].iloc[0][1]
    current_year = startdate.year
    current = (current_q, current_year)

    prev = (str(int(current_q) - 1), current_year)
    one_after = (str(int(current_q) + 1), current_year)
    two_after = (str(int(current_q) + 2), current_year)

    if current_q == '1':
        prev = ('4', current_year - 1)
    elif current_q == '4':
        one_after = ('1', current_year + 1)
        two_after = ('2', current_year + 1)
    elif current_q == '3':
        two_after = ('1', current_year + 1)

    return prev, current, one_after, two_after


# Find the start date of each quarters
def quarter_startdate(quarters, four_q):
    l = []
    for q in four_q:
        a = quarters[quarters['end_date'].str.contains(str(q[1]))]
        a = a[a['quarter'].astype(str) == 'Q' + q[0]]
        l.append(a['start_date'].iloc[0])
    return l


# Get the prime, nonprime and total baselayers
def past(df, startdate, enddate):
    parts = df[['Guarantee ID', 'Week Start Date', 'ADU Ind', 'Equiv Units', 'Selling Title']]
    weeks = week_range(startdate, enddate)

    dic_s_p = dict()
    dic_s_np = dict()
    dic_adu_p = dict()
    dic_adu_np = dict()
    dic_total_p = dict()
    dic_total_np = dict()

    for ind, row in parts.iterrows():
        if row['Guarantee ID'] not in dic_s_p:
            dic_s_p[row['Guarantee ID']] = [0] * len(weeks)
            dic_s_np[row['Guarantee ID']] = [0] * len(weeks)
            dic_adu_p[row['Guarantee ID']] = [0] * len(weeks)
            dic_adu_np[row['Guarantee ID']] = [0] * len(weeks)
            dic_total_p[row['Guarantee ID']] = [0] * len(weeks)
            dic_total_np[row['Guarantee ID']] = [0] * len(weeks)      
            
        #reformat_w = str(datetime.strptime(str(row['Week Start Date']).split()[0], '%Y-%m-%d').strftime('%m/%d/%Y'))
        reformat_w = row['Week Start Date']
        if reformat_w in weeks:
            if dayparts(row) == 'P':
                if row['ADU Ind'] == 'N':  # Prime spots
                    dic_s_p[row['Guarantee ID']][weeks.index(reformat_w)] += row['Equiv Units']
                else: # Prime ADU
                    dic_adu_p[row['Guarantee ID']][weeks.index(reformat_w)] += row['Equiv Units']
                dic_total_p[row['Guarantee ID']][weeks.index(reformat_w)] += row['Equiv Units'] # Prime Total
            else:
                if row['ADU Ind'] == 'N': # Nonprime spots
                    dic_s_np[row['Guarantee ID']][weeks.index(reformat_w)] += row['Equiv Units']
                else: # Nonprime ADU
                    dic_adu_np[row['Guarantee ID']][weeks.index(reformat_w)] += row['Equiv Units']
                dic_total_np[row['Guarantee ID']][weeks.index(reformat_w)] += row['Equiv Units'] #Nonprime Total
    return dic_s_p, dic_s_np, dic_adu_p, dic_adu_np, dic_total_p, dic_total_np


# round units to a multiple of 1, and calculate the leftover value
def round_unit(num):
    result = round(num)
    left = num - result
    return (result, left)


def schedule_ADU(past_s_p, past_adu_p, past_s_np, past_adu_np, df1, startq, endq, startdate):
    weeks = week_range(startq, endq)

    dic_p = dict()
    dic_np = dict()
    total_weeks = week_range(startq, endq)

    l = []
    for ind, row in df1.iterrows():
        # Schedule start date
        if date_comparison(row['Sold Start Date'], startdate):
            s = pd.date_range(startdate, periods=1, freq='W-MON').strftime('%m/%d/%Y').tolist()[0]
        else:
            s = row['Sold Start Date']
            s = pd.date_range(s, periods=1, freq='W-MON').strftime('%m/%d/%Y').tolist()[0]
        
        # Schedule end date
        if date_comparison(row['Sold End Date'], weeks[-1]):
            e = row['Sold End Date']
            e = pd.date_range(e, periods=1, freq='W-MON').strftime('%m/%d/%Y').tolist()[0]
        else:
            e = weeks[-1]
        weeks_left = weeks_between(s, e) + 1

        if weeks_left <= 0: # If no available week, do not schedule anything
            dic_p[row['Guarantee ID']] = [0] * len(weeks)
            dic_np[row['Guarantee ID']] = [0] * len(weeks)

        else:
            if row['Guarantee ID'] not in dic_p:
                dic_p[row['Guarantee ID']] = [0] * len(weeks)
                dic_np[row['Guarantee ID']] = [0] * len(weeks)
            
            if row['P ADUs'] > 0: #If there is Prime ADU
                scheduled_spots = past_s_p[row['Guarantee ID']][total_weeks.index(s):] # scheduled spots = prime spots baselayer
                total = sum(scheduled_spots[:total_weeks.index(weeks[-1]) + 1]) # total number of prime spots 
                new = dic_p[row['Guarantee ID']]
                if total == 0: # if no prime spots
                    try: # Check whether there is ADU scheduled
                        scheduled_spots = past_adu_p[row['Guarantee ID']][total_weeks.index(s):] # scheduled spots = prime ADU baselayer
                        total = sum(scheduled_spots[:total_weeks.index(weeks[-1]) + 1]) # total number of prime ADU
                    except: # if no spots and ADUs scheduled, do not schedule new ADU
                        dic_p[row['Guarantee ID']] = new
                        
                left = 0 
                for i in range(weeks.index(s), weeks.index(e) + 1): # schedule new ADU proportional to the scheduled_spots
                    if scheduled_spots[i - weeks.index(s)] != 0:
                        new[i] = round_unit(row['P ADUs'] / total * scheduled_spots[i - weeks.index(s)] + left)[0]
                        left = round_unit(row['P ADUs'] / total * scheduled_spots[i - weeks.index(s)] + left)[1]
                dic_p[row['Guarantee ID']] = new

            if row['NP ADUs'] > 0: #If there is Nonprime ADU

                scheduled_spots = past_s_np[row['Guarantee ID']][total_weeks.index(s):]
                total = sum(scheduled_spots[:total_weeks.index(weeks[-1]) + 1])
                new = dic_np[row['Guarantee ID']]
                if total == 0:
                    try:
                        scheduled_spots = past_adu_np[row['Guarantee ID']][total_weeks.index(s):]
                        total = sum(scheduled_spots[:total_weeks.index(weeks[-1]) + 1])
                    except:
                        dic_np[row['Guarantee ID']] = new
                        
                left = 0
                for i in range(weeks.index(s), weeks.index(e) + 1):
                    if scheduled_spots[i - weeks.index(s)] != 0:
                        new[i] = round_unit(row['NP ADUs'] / total * scheduled_spots[i - weeks.index(s)] + left)[0]
                        left = round_unit(row['NP ADUs'] / total * scheduled_spots[i - weeks.index(s)] + left)[1]
                dic_np[row['Guarantee ID']] = new        
                
                
                
            if row['P ADUs'] < 0: #If there is Prime ADU need to take back
                new = dic_p[row['Guarantee ID']]
                try: # Check whether there is ADU scheduled
                    scheduled_spots = past_adu_p[row['Guarantee ID']][total_weeks.index(s):] # scheduled spots = prime ADU baselayer
                    total = sum(scheduled_spots[:total_weeks.index(weeks[-1]) + 1]) # total number of prime ADU
                except: # if ADUs scheduled, do not take back ADU
                    dic_p[row['Guarantee ID']] = new
                        
                left = 0 
                for i in range(weeks.index(s), weeks.index(e) + 1): # take back ADU proportional to the scheduled_spots
                    u = scheduled_spots[i - weeks.index(s)]
                    if u != 0:
                        new[i] = -min(u, -round(row['P ADUs'] / total * u + left))
                        left = (row['P ADUs'] / total * u + left) - new[i]
                dic_p[row['Guarantee ID']] = new

            if row['NP ADUs'] < 0: #If there is Prime ADU need to take back
                new = dic_np[row['Guarantee ID']]
                try: # Check whether there is ADU scheduled
                    scheduled_spots = past_adu_np[row['Guarantee ID']][total_weeks.index(s):] # scheduled spots = prime ADU baselayer
                    total = sum(scheduled_spots[:total_weeks.index(weeks[-1]) + 1]) # total number of prime ADU
                except: # if ADUs scheduled, do not take back ADU
                    dic_np[row['Guarantee ID']] = new
                        
                left = 0 
                for i in range(weeks.index(s), weeks.index(e) + 1): # take back ADU proportional to the scheduled_spots
                    u = scheduled_spots[i - weeks.index(s)]
                    if u != 0:
                        new[i] = -min(u, -round(row['NP ADUs'] / total * u + left))
                        left = (row['NP ADUs'] / total * u + left) - new[i]
                dic_np[row['Guarantee ID']] = new
      
                
    schedule_p_df = pd.DataFrame.from_dict(dic_p, orient='index', columns=weeks).reset_index()
    schedule_p_df = schedule_p_df.rename(columns={'index': 'Guarantee ID'})
    schedule_np_df = pd.DataFrame.from_dict(dic_np, orient='index', columns=weeks).reset_index()
    schedule_np_df = schedule_np_df.rename(columns={'index': 'Guarantee ID'})
    return schedule_p_df, schedule_np_df


#@df is the raw dealmake data
#@quater deleted ?
#data_String: the data to start adu schdul
def raw_result(df, quarters, date_string, startdate, ratings_file, four_q, startq, endq):
   
    weeks = week_range(startq, endq)

    # Read in Ratings
    internal_estimates = pd.read_excel(ratings_file)
    ratings = get_ratings(df, internal_estimates, int(four_q[1][0]))
    
    
    result = get_dict(df, ratings, endq)
    output = form_df(result)

    df1 = output[['Guarantee ID', 'Sold Start Date', 'Sold End Date', 'P ADUs', 'NP ADUs', 'Total ADUs']]

    past_s_p, past_s_np, past_adu_p, past_adu_np, bp, bnp = past(df, startq, endq)

    baselayer_p = pd.DataFrame.from_dict(bp, orient='index', columns=weeks).reset_index()
    baselayer_np = pd.DataFrame.from_dict(bnp, orient='index', columns=weeks).reset_index()
    
    baselayer_p = baselayer_p.rename(columns={'index': 'Guarantee ID'})
    baselayer_np = baselayer_np.rename(columns={'index': 'Guarantee ID'})

    P_ADU_schedule, NP_ADU_schedule = schedule_ADU(past_s_p, past_adu_p, past_s_np, past_adu_np, df1, startq, endq, startdate)

    basic_info = output.copy()

    return date_string, basic_info, baselayer_p, baselayer_np, P_ADU_schedule, NP_ADU_schedule #, changeDF


def format_df(raw, new, name):
    writer = pd.ExcelWriter(datetime.strptime(str(datetime.now().strftime("%m/%d/%Y")), '%m/%d/%Y').strftime('%Y-%m-%d')+' '+ name + '.xlsx', engine='xlsxwriter')#, datetime_format='%m/%d/%Y')
    workbook = writer.book

    count_row = raw[1].shape[0] + 1  # gives number of row count
    count_col = raw[1].shape[1] + 3  # gives number of col count
    raw[1].to_excel(writer, sheet_name=name, startrow=7, startcol=2, header=False, index = False)
    new.to_excel(datetime.strptime(str(datetime.now().strftime("%m/%d/%Y")), '%m/%d/%Y').strftime('%Y-%m-%d')+' ADU Report.xlsx', index = False)

    worksheet = writer.sheets[name]
    
    # Clean the headers
    for col_num, value in enumerate(raw[1].columns.values):
        if col_num <= 7:
            worksheet.write(5, col_num + 2, value)
        elif col_num <= 10:
            worksheet.write(5, col_num + 2, ' '.join(value.split()[1:]))
        elif col_num <= 34:
            worksheet.write(5, col_num + 2, ' '.join(value.split()[2:]))
        else:
            worksheet.write(5, col_num + 2, ' '.join(value.split()[1:]))

    s = [] # stores the start column of each dataframe
    e = [] # stores the end column of each dataframe
    for i in range(2, len(raw)):
        raw[i].iloc[:, 1:].to_excel(writer, sheet_name=name, startrow=7, startcol=count_col, index=False,
                                    header=False)
        for col_num, value in enumerate(raw[i].columns.values[1:]):
            worksheet.write(5, count_col + col_num, value)
        s.append(count_col)
        for r in range(count_row):
            for c in range(count_col, count_col):
                worksheet.write_blank(r, c, None)

        count_col += raw[i].shape[1]
        e.append(count_col - 2)
    s_letter = ['B'] #start column letter of each dataframe
    e_letter = ['L'] #end column letter of each dataframe
    for i in range(len(s)):
        s_letter.append(xlsxwriter.utility.xl_col_to_name(s[i]))
        e_letter.append(xlsxwriter.utility.xl_col_to_name(e[i]))

    # sum of scheduled spots and ADUs
    for i in range(len(e)):
        col = xlsxwriter.utility.xl_col_to_name(e[i] + 1)
        for r in range(8, count_row + 7):
            worksheet.write_formula(col + str(r),
                                    '{=SUM(' + s_letter[i + 1] + str(r) + ':' + e_letter[i + 1] + str(r) + ')}')

    # Deals not in flight
    col = xlsxwriter.utility.xl_col_to_name(e[-1] + 2)
    Total_P_ADU_col = xlsxwriter.utility.xl_col_to_name(e[2] + 1)
    Total_NP_ADU_col = xlsxwriter.utility.xl_col_to_name(e[3] + 1)
    Total_ADU_col = xlsxwriter.utility.xl_col_to_name(s[0] - 2)
    for r in range(8, count_row + 7):
        worksheet.write_formula(col + str(r), '{=' + Total_ADU_col + str(r) + '-' + Total_P_ADU_col + str(
            r) + '-' + Total_NP_ADU_col + str(r) + '}')

    # Header
    bold = workbook.add_format({'bold': True})
    worksheet.write(1, 1, 'ION Media', bold)
    worksheet.write(2, 1, 'ADU Trust 3.0', bold)
    bold_blue = workbook.add_format({'bold': True, 'font_color': 'blue'})
    worksheet.write(2, 2, raw[0], bold_blue)

    
    # Add Title & Merge
    format_b = workbook.add_format({
        'bold': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#99CCFF'})
    format_o = workbook.add_format({
        'bold': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#FFCC99'})
    format_y = workbook.add_format({
        'bold': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#FFFFCC'})
    format_g = workbook.add_format({
        'bold': 1,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#C0C0C0'})  # grey

    try:
        worksheet.merge_range('C4:I4', 'DEAL', format_g)
        worksheet.merge_range('C5:I5', ' ', format_g)
        worksheet.merge_range(s_letter[3] + '4:' + e_letter[3] + '4', 'Prime - ADU Suggested Flighting', format_b)
        worksheet.merge_range(s_letter[4] + '4:' + e_letter[4] + '4', 'Non Prime - ADU Suggested Flighting', format_o)
        worksheet.merge_range(s_letter[1] + '4:' + e_letter[1] + '4', 'Prime Fligting - Sold Units', format_b)
        worksheet.merge_range(s_letter[2] + '4:' + e_letter[2] + '4', 'Non Prime Fligting - Sold Units', format_o)

    except:
        print('nope')

    # Headers for dataframes
    for i in range(9, 58):
        if i <= 12:
            worksheet.write(3, i, 'SOLD', format_g)
            worksheet.write(4, i, ' ', format_g)
        elif i <= 18:
            worksheet.write(3, i, 'SOLD', format_b)
            worksheet.write(4, i, 'Prime', format_b)
        elif i <= 24:
            worksheet.write(3, i, 'SOLD', format_o)
            worksheet.write(4, i, 'NP', format_o)
        elif i <= 30:
            worksheet.write(3, i, 'ADU', format_b)
            worksheet.write(4, i, 'Prime', format_b)
        elif i <= 36:
            worksheet.write(3, i, 'ADU', format_o)
            worksheet.write(4, i, 'NP', format_o)
        elif i <= 46:
            worksheet.write(3, i, 'Total', format_g)
            worksheet.write(4, i, ' ', format_g)
        else:
            worksheet.write(3, i, ' ', format_g)
            if i != 57:
                if i % 2 == 1:
                    worksheet.write(4, i, 'P', format_g)
                else:
                    worksheet.write(4, i, 'NP', format_g)
            else:
                worksheet.write(4, i, 'Total', format_g)

    for i in range(s[0], e[0] + 2):
        worksheet.write(4, i, 'P', format_b)
        if i == e[0] + 1:
            worksheet.write(3, i, 'Total', format_b)
    for i in range(s[1], e[1] + 2):
        worksheet.write(4, i, 'NP', format_o)
        if i == e[1] + 1:
            worksheet.write(3, i, 'Total', format_o)
    for i in range(s[2], e[2] + 2):
        worksheet.write(4, i, 'P', format_b)
        if i == e[2] + 1:
            worksheet.write(3, i, 'Total', format_b)
    for i in range(s[3], e[3] + 2):
        worksheet.write(4, i, 'NP', format_o)
        if i == e[3] + 1:
            worksheet.write(3, i, 'Total', format_o)
    worksheet.write(3, e[3] + 2, 'Deals', format_g)
    worksheet.write(4, e[3] + 2, 'Not in', format_g)
    worksheet.write(5, e[3] + 2, 'Flight', format_g)

    # Group Columns
    worksheet.set_column('D:E', None, None, {'level': 1})
    worksheet.set_column('G:H', None, None, {'level': 1})
    worksheet.set_column('L:AK', None, None, {'level': 1})
    worksheet.set_column('AX:BA', None, None, {'level': 1})

    worksheet.set_column(s_letter[1] + ':' + xlsxwriter.utility.xl_col_to_name(e[1] + 1), None, None, {'level': 1})

    # Autofilter
    worksheet.autofilter('A7:' + xlsxwriter.utility.xl_col_to_name(e[-1] + 2) + str(count_row+100))

    # Get the Sum
    for col in range(s[0] - 4, e[3] + 3):
        col = xlsxwriter.utility.xl_col_to_name(col)
        worksheet.write_formula(col + str(count_row + 8),
                                '{=subtotal(9, ' + col + '8:' + col + str(count_row + 7) + ')}')

    # Conditional format for date
    # Add a format. Light red fill with dark red text.
    format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006'})
    # Add a format. Green fill with dark green text.
    format2 = workbook.add_format({'bg_color': '#C6EFCE',
                                   'font_color': '#006100'})
    format3 = workbook.add_format({'bg_color': 'white'})

    worksheet.conditional_format(s_letter[3] + '6:' + e_letter[3] + '6', {'type': 'cell',
                                                                          'criteria': '<',
                                                                          'value': '$C$3',
                                                                          'format': format1})
    worksheet.conditional_format(s_letter[3] + '6:' + e_letter[3] + '6', {'type': 'cell',
                                                                          'criteria': '>=',
                                                                          'value': '$C$3',
                                                                          'format': format2})
    worksheet.conditional_format(s_letter[4] + '6:' + e_letter[4] + '6', {'type': 'cell',
                                                                          'criteria': '<',
                                                                          'value': '$C$3',
                                                                          'format': format1})
    worksheet.conditional_format(s_letter[4] + '6:' + e_letter[4] + '6', {'type': 'cell',
                                                                          'criteria': '>=',
                                                                          'value': '$C$3',
                                                                          'format': format2})

    # Column Width
    worksheet.set_column(s_letter[0] + ':' + e_letter[0], 15)

    # freeze the top rows and left columns
    worksheet.freeze_panes(7, 11)

    writer.save()
    return s, s_letter, e_letter


def new_data(raw, quarters):
    general = dict()
    P_ADU_dict = raw[4].to_dict('index')
    NP_ADU_dict = raw[5].to_dict('index')

    for k, v in P_ADU_dict.items():
        for key, value in v.items():
            if key == 'Guarantee ID':
                gid = value
            # key is Monday of scheduling weeks
            else: 
                # if there is new schedule for this week
                if value > 0: 
                    if gid not in general:
                        general[gid] = {'Year': [], 'Quarter': [], 'Year + Quarter': [], 'Week Start Date': [],
                                        'Week End Date': [], 'Selling Title': [], \
                                        'Days And Times': [], 'ADU Ind': [], 'Booked Dollars': [],
                                        'Primary Demo Equiv Deal Imp': [], \
                                        #'Primary Demo Equiv Post Imp - IE 1': [],
                                        'Primary Demo Non-ADU Equiv Deal Imp': [], \
                                        #'Primary Demo Equiv Ratecard Imp': [], 
                                        'Primary Demo Deal CPM': [], \
                                        'Equiv Units': []}
                    # filling all the information
                    general[gid]['Week Start Date'].append(key)
                    general[gid]['Equiv Units'].append(value)

                    y = key.split('/')[2]
                    general[gid]['Year'].append(y)

                    mon = pd.date_range(key, periods=1, freq='W-MON').strftime('%m/%d/%Y').tolist()[0]
                    q = quarters.loc[quarters['start_date'].astype(str) == mon, 'quarter'].iloc[0][1]

                    general[gid]['Quarter'].append(q)
                    general[gid]['Year + Quarter'].append(y + ' ' + str(q) + 'Q')
                    general[gid]['Week End Date'].append(pd.date_range(key, periods=1, freq='W-SUN').strftime('%m/%d/%Y').tolist()[0])
                    general[gid]['Selling Title'].append('P')
                    general[gid]['Days And Times'].append('')
                    general[gid]['ADU Ind'].append('Y')
                    general[gid]['Booked Dollars'].append(0)
                    general[gid]['Primary Demo Equiv Deal Imp'].append(0)
                    #general[gid]['Primary Demo Equiv Post Imp - IE 1'].append(0)
                    general[gid]['Primary Demo Non-ADU Equiv Deal Imp'].append(0)
                    #general[gid]['Primary Demo Equiv Ratecard Imp'].append(0)
                    general[gid]['Primary Demo Deal CPM'].append(0)
                
               
    for k, v in NP_ADU_dict.items():
        for key, value in v.items():
            if key == 'Guarantee ID':
                gid = value
            else:
                if value > 0.1:
                    if gid not in general:
                        general[gid] = {'Year': [], 'Quarter': [], 'Year + Quarter': [], 'Week Start Date': [],
                                        'Week End Date': [], 'Selling Title': [], \
                                        'Days And Times': [], 'ADU Ind': [], 'Booked Dollars': [],
                                        'Primary Demo Equiv Deal Imp': [], \
                                        #'Primary Demo Equiv Post Imp - IE 1': [],
                                        'Primary Demo Non-ADU Equiv Deal Imp': [], \
                                        #'Primary Demo Equiv Ratecard Imp': [], \
                                        'Primary Demo Deal CPM': [], \
                                        'Equiv Units': []}
                    general[gid]['Week Start Date'].append(key)
                    general[gid]['Equiv Units'].append(value)

                    y = key.split('/')[2]
                    general[gid]['Year'].append(y)

                    mon = pd.date_range(key, periods=1, freq='W-MON').strftime('%m/%d/%Y').tolist()[0]
                    q = quarters.loc[quarters['start_date'].astype(str) == mon, 'quarter'].iloc[0][1]

                    general[gid]['Quarter'].append(q)
                    general[gid]['Year + Quarter'].append(y + ' ' + str(q) + 'Q')
                    general[gid]['Week End Date'].append(pd.date_range(key, periods=1, freq='W-SUN').strftime('%m/%d/%Y').tolist()[0])
                    general[gid]['Selling Title'].append('NP')
                    general[gid]['Days And Times'].append('')
                    general[gid]['ADU Ind'].append('Y')
                    general[gid]['Booked Dollars'].append(0)
                    general[gid]['Primary Demo Equiv Deal Imp'].append(0)
                    #general[gid]['Primary Demo Equiv Post Imp - IE 1'].append(0)
                    general[gid]['Primary Demo Non-ADU Equiv Deal Imp'].append(0)
                    #general[gid]['Primary Demo Equiv Ratecard Imp'].append(0)
                    general[gid]['Primary Demo Deal CPM'].append(0)
                
    return general


def newdata_to_df(df, general, output):
    df['In System'] = 'Y'
    basics = df[['Guarantee ID', 'Guarantee Name', 'Deal Numbers in Guarantee', 'Marketplace', \
                 'Advertiser', \
                 'AE Name', 'Agency Name (Billing)', 'Deal Name', 'Deal Number', 'Deal Flight Start Date', \
                 'Deal Flight End Date', 'Primary Demo']]
    
    # Create a dataframe from general dictionary
    column_names = ['Guarantee ID', 'Year', 'Quarter', 'Year + Quarter', 'Week Start Date', 'Week End Date',
                    'Selling Title', \
                    'Days And Times', 'ADU Ind', 'Booked Dollars', 'Primary Demo Equiv Deal Imp', \
                    #'Primary Demo Equiv Post Imp - IE 1', \
                    'Primary Demo Non-ADU Equiv Deal Imp', \
                    #'Primary Demo Equiv Ratecard Imp', 
                    'Primary Demo Deal CPM', \
                    'Equiv Units']
    rows = []
    for k, v in general.items():
        for i in range(len(v['Year'])):
            row = []
            row.append(k)
            for key in column_names[1:]:
                row.append(v[key][i])
            rows.append(row)

    newdata_df = pd.DataFrame(rows)


    newdata_df.columns = column_names
    newdata_df['In System'] = 'N'

    # To get the basic information for new data
    combined = pd.merge(newdata_df, basics, how='left', on='Guarantee ID').drop_duplicates(
        subset=['Guarantee ID', 'Week Start Date', 'Week End Date', 'Selling Title'])
    
    # To get impression for new data
    imp_df = output[['Guarantee ID', 'P Q2 Imp', 'NP Q2 Imp']]
    combined = pd.merge(combined, imp_df, on='Guarantee ID')

    imp = dict()
    for i, r in combined.iterrows():
        if r['Selling Title'] == 'P':
            imp[(r['Guarantee ID'], 'P', r['Equiv Units'])] = r['P Q2 Imp'] * r['Equiv Units'] * 1000
        else:
            imp[(r['Guarantee ID'], 'NP', r['Equiv Units'])] = r['NP Q2 Imp'] * r['Equiv Units'] * 1000

    ADU_E_D_I = pd.Series(imp).rename_axis(['Guarantee ID', 'Selling Title', 'Equiv Units']).reset_index(
        name='Primary Demo Equiv Post Imp')

    combined = pd.merge(combined, ADU_E_D_I, how='left', on=['Guarantee ID', 'Selling Title', 'Equiv Units'])
    #combined['Primary Demo Equiv Post Imp'] = combined['Primary Demo ADU Equiv Deal Imp']

    combined = combined[['Guarantee ID', 'Guarantee Name', 'Deal Numbers in Guarantee', 'Marketplace', \
                         'Advertiser', \
                         'AE Name', 'Agency Name (Billing)', 'Deal Name', 'Deal Number', 'Deal Flight Start Date', \
                         'Deal Flight End Date', 'Primary Demo', 'Year', 'Quarter', 'Year + Quarter', 'Week Start Date', \
                         'Week End Date', 'Selling Title', 'Days And Times', 'ADU Ind', 'Booked Dollars', \
                         'Primary Demo Equiv Deal Imp', \
                         #'Primary Demo Equiv Post Imp - IE 1', \
                         'Primary Demo Non-ADU Equiv Deal Imp', \
                         #'Primary Demo Equiv Ratecard Imp', \
                         'Primary Demo Equiv Post Imp', 'Primary Demo Deal CPM', 'Equiv Units', \
                         'In System']]

    total = pd.concat([df, combined], sort=False)

    return total


def liability(new):
    # Sort the dataframe
    new['Week Start Date'] =  pd.to_datetime(new['Week Start Date'])
    new['Week End Date'] =  pd.to_datetime(new['Week End Date'])
    df_sort = new.sort_values(['Guarantee ID', 'Week Start Date', 'Week End Date', 'Booked Dollars'],
                              ascending=[True, True, True, False])
    Acc_Deal_Imp = set()
    Acc_Deal_Imp_list = []
    Acc_Delv_Imp = []
    ACC_Effec_Delv_Imp = []
    Effec_Delv_Imp = []
    Owed_Imp = []

    Acc_Deal_value = []
    Effec_Delv_value = []
    Acc_Effec_Delv_value = []
    Owed_value = []
    
    # Compute impressions and values
    for i, r in df_sort.iterrows():

        if r['Guarantee ID'] not in Acc_Deal_Imp:
            Acc_Deal_Imp.add(r['Guarantee ID'])
            a = r['Primary Demo Non-ADU Equiv Deal Imp']
            b = r['Primary Demo Equiv Post Imp']
            c = min(a, b)
            d = c

            A = float(r['Booked Dollars'])

            pool = []            
            Guar = r['Primary Demo Non-ADU Equiv Deal Imp'] * float(r['Primary Demo Deal CPM'])
            pool.append([r['Primary Demo Non-ADU Equiv Deal Imp'], float(r['Primary Demo Deal CPM'])])
            B = d * float(r['Primary Demo Deal CPM'])

            pool[0][0] -= d
            C = B


        else:
            a = Acc_Deal_Imp_list[-1] + r['Primary Demo Non-ADU Equiv Deal Imp']
            b = Acc_Delv_Imp[-1] + r['Primary Demo Equiv Post Imp']
            c = min(a, b)
            d = c - ACC_Effec_Delv_Imp[-1]

            A = Acc_Deal_value[-1] + float(r['Booked Dollars'])

            Guar = r['Primary Demo Non-ADU Equiv Deal Imp'] * float(r['Primary Demo Deal CPM'])
            pool.append([r['Primary Demo Non-ADU Equiv Deal Imp'], float(r['Primary Demo Deal CPM'])])

            B = 0
            imp = d
            while imp > pool[0][0]:
                temp = pool.pop(0)
                B += temp[0] * temp[1]
                imp -= temp[0]
            if imp > 0:
                B += imp * pool[0][1]
                pool[0][0] -= imp

            C = Acc_Effec_Delv_value[-1] + B

        e = r['Primary Demo Non-ADU Equiv Deal Imp'] - d
        D = Guar - B

        Acc_Deal_Imp_list.append(a)
        Acc_Delv_Imp.append(b)
        ACC_Effec_Delv_Imp.append(c)
        Effec_Delv_Imp.append(d)
        Owed_Imp.append(e)

        Acc_Deal_value.append(A)
        Effec_Delv_value.append(B)
        Acc_Effec_Delv_value.append(C)
        Owed_value.append(D)

    df_sort['Acc_Deal_Imp'] = Acc_Deal_Imp_list
    df_sort['Acc_Delv_Imp'] = Acc_Delv_Imp
    df_sort['Acc_Effec_Delv_Imp'] = ACC_Effec_Delv_Imp
    df_sort['Effec_Delv_Imp'] = Effec_Delv_Imp
    df_sort['Owed_Imp'] = Owed_Imp

    df_sort['Acc_Deal_value'] = Acc_Deal_value
    df_sort['Effec_Delv_value'] = Effec_Delv_value
    df_sort['Acc_Effec_Delv_value'] = Acc_Effec_Delv_value
    df_sort['Owed_value'] = Owed_value

    df_sort['Primary Demo Equiv Deal Imp'] = df_sort['Primary Demo Equiv Deal Imp']/1000
    #df_sort['Primary Demo Equiv Post Imp - IE 1'] = df_sort['Primary Demo Equiv Post Imp - IE 1'] / 1000
    #df_sort['Primary Demo ADU Equiv Deal Imp'] = df_sort['Primary Demo ADU Equiv Deal Imp'] / 1000
    df_sort['Primary Demo Non-ADU Equiv Deal Imp'] = df_sort['Primary Demo Non-ADU Equiv Deal Imp'] / 1000
    #df_sort['Primary Demo Equiv Ratecard Imp'] = df_sort['Primary Demo Equiv Ratecard Imp'] / 1000
    df_sort['Primary Demo Equiv Post Imp'] = df_sort['Primary Demo Equiv Post Imp'] / 1000
    df_sort['Acc_Deal_Imp'] = df_sort['Acc_Deal_Imp'] / 1000
    df_sort['Acc_Delv_Imp'] = df_sort['Acc_Delv_Imp'] / 1000
    df_sort['Acc_Effec_Delv_Imp'] = df_sort['Acc_Effec_Delv_Imp'] / 1000
    df_sort['Effec_Delv_Imp'] = df_sort['Effec_Delv_Imp']/1000
    df_sort['Owed_Imp'] = df_sort['Owed_Imp']/1000
    
    df_sort['Effec_Delv_value'] = df_sort['Effec_Delv_value']/1000
    df_sort['Acc_Effec_Delv_value'] = df_sort['Acc_Effec_Delv_value']/1000
    df_sort['Owed_value'] = df_sort['Owed_value']/1000


    return df_sort

def combine_demo(df):
    demos = df['Primary Demo'].unique().tolist()
    startpoint = ['2','6','9','12','15','18','21','25','30','35','40','45', '50','55','65']
    demo_list = ['2-5','6-8','9-11','12-14', '15-17', '18-20', '21-24', '25-29', '30-34', '35-39', '40-44', '45-49', '50-54', '55-64', '65+']
    demo_dic = {'HH':['HHLD']}
    for d in demos:
        if d != 'HH':
            demo_dic[d] = []
            gen = d[0]
            if '+' not in d:
                s, e = d[1:].split('-')
                #print(s,e)
                s_ind = startpoint.index(s)
                e_ind = startpoint.index(str(int(e)+1))
                if gen != 'P':
                    for i in range(s_ind, e_ind):
                        demo_dic[d].append(gen + demo_list[i])
                else:
                    for i in range(s_ind, e_ind):
                        demo_dic[d].append('F' + demo_list[i])
                        demo_dic[d].append('M' + demo_list[i])

            else:
                s = d[1:-1]
                s_ind = startpoint.index(s)
                if gen != 'P':
                    for i in range(s_ind, len(demo_list)):
                        demo_dic[d].append(gen + demo_list[i])
                else:
                    for i in range(s_ind, len(demo_list)):
                        demo_dic[d].append('F' + demo_list[i])
                        demo_dic[d].append('M' + demo_list[i])
    return demo_dic

def get_ratings(df, internal_estimates, cur_q):
    estimates_NP = internal_estimates.loc[(internal_estimates['Selling Title'] == 'MSU7A7P1A3A') & (internal_estimates['Quarter'] ==cur_q )]
    estimates_P = internal_estimates.loc[(internal_estimates['Selling Title'] == 'MSU7p1a')& (internal_estimates['Quarter'] ==cur_q)]
    demo_dic = combine_demo(df)
    
    demo_ratings_P = dict()
    demo_ratings_NP = dict()

    rating_dic_P = list(estimates_P.to_dict(orient = 'index').values())[0]
    rating_dic_NP = list(estimates_NP.to_dict(orient = 'index').values())[0]

    for k,v in demo_dic.items():
        demo_ratings_P[k] = 0
        demo_ratings_NP[k] = 0

        for d in v:
            demo_ratings_P[k] += rating_dic_P[d]
            demo_ratings_NP[k] += rating_dic_NP[d]

    demo_ratings_P = pd.DataFrame.from_dict(demo_ratings_P, orient='index').reset_index()
    demo_ratings_NP = pd.DataFrame.from_dict(demo_ratings_NP, orient='index').reset_index()
    ratings = pd.merge(demo_ratings_P, demo_ratings_NP, on='index')
    ratings.columns = ['Demo', 'Prime Imp', 'Non Prime Imp']
    return ratings

def seperate(raw):
    df = raw[1]
    pos = df[df['Total Imps Owed']>=0]
    gid = pos['Guarantee ID'].tolist()

    sch = raw[0], pos, raw[2][raw[2]['Guarantee ID'].isin(gid)], raw[3][raw[3]['Guarantee ID'].isin(gid)], raw[4][raw[4]['Guarantee ID'].isin(gid)], raw[5][raw[5]['Guarantee ID'].isin(gid)]
    takeback = raw[0], df[df['Total Imps Owed']<0], raw[2][~raw[2]['Guarantee ID'].isin(gid)], raw[3][~raw[3]['Guarantee ID'].isin(gid)], raw[4][~raw[4]['Guarantee ID'].isin(gid)], raw[5][~raw[5]['Guarantee ID'].isin(gid)]
    return sch, takeback


def get_report_values(quarters, startdate, liab):
    last_q = quarter_startdate(quarters, find_quarters(quarters, startdate))[0]
    report_q = find_quarters(quarters, datetime.strptime(last_q, '%m/%d/%Y'))
    quar = []
    for i in range(4):
        quar.append(str(report_q[i][1]) + ' ' + report_q[i][0] + 'Q')
    table1 = pd.pivot_table(liab[liab['In System']=='Y'], index = 'Year + Quarter', columns = 'ADU Ind', values=['Owed_value', 'Owed_Imp', 'Equiv Units'], aggfunc=np.sum, fill_value=0, margins = True)
    table2 = pd.pivot_table(liab, index = 'Year + Quarter', columns = ['ADU Ind'], values=['Owed_value', 'Owed_Imp', 'Equiv Units'], aggfunc=np.sum, fill_value=0, margins = True)

    rating = 310
    begin_liab = []
    begin_imp_owed = []
    begin_adu_req = []

    cur_q_liab = []
    cur_q_imp_owed = []
    cur_q_adu_req = []

    cur_q_liab_paid = []
    cur_q_imp_paid = []
    cur_q_adu_given = []
    cur_q_liab_paid_new = []
    cur_q_imp_paid_new = []
    cur_q_adu_given_new = []

    ending_liab = []
    ending_imp_owed = []
    ending_adu_req =[]
    ending_liab_new = []
    ending_imp_owed_new = []
    ending_adu_req_new = []


    table1.reset_index(inplace=True)
    #print(table1)

    order = table1['Year + Quarter'].tolist()

    owed_v_spots = table1['Owed_value']['N'].tolist()
    owed_v_adu = table1['Owed_value']['Y'].tolist()
    owed_v_total = table1['Owed_value']['All'].tolist()

    owed_imp_spots = table1['Owed_Imp']['N'].tolist()
    owed_imp_adu = table1['Owed_Imp']['Y'].tolist()
    owed_imp_total = table1['Owed_Imp']['All'].tolist()

    adu_units = table1['Equiv Units']['Y'].tolist()

    owed_v_adu_new = table2['Owed_value']['Y'].tolist()
    owed_imp_adu_new = table2['Owed_Imp']['Y'].tolist()
    adu_units_new = table2['Equiv Units']['Y'].tolist()


    owed_v_total_new = table2['Owed_value']['All'].tolist()
    owed_imp_total_new = table2['Owed_Imp']['All'].tolist()

    i = order.index(quar[0])
    for j in range(i, i+4):
        begin_liab.append(sum(owed_v_total[:j]))
        begin_imp_owed.append(sum(owed_imp_total[:j]))
        begin_adu_req.append(sum(owed_imp_total[:j])/rating)

        ending_liab.append(sum(owed_v_total[:j+1]))
        ending_imp_owed.append(sum(owed_imp_total[:j+1]))
        ending_adu_req.append(sum(owed_imp_total[:j+1])/rating)
        ending_liab_new.append(sum(owed_v_total_new[:j+1]))
        ending_imp_owed_new.append(sum(owed_imp_total_new[:j+1]))
        ending_adu_req_new.append(sum(owed_imp_total_new[:j+1])/rating)

    for q in quar:
        i = order.index(q)

        cur_q_liab.append(owed_v_spots[i])
        cur_q_imp_owed.append(owed_imp_spots[i])
        cur_q_adu_req.append(owed_imp_spots[i]/rating)


        cur_q_liab_paid.append(owed_v_adu[i])
        cur_q_imp_paid.append(owed_imp_adu[i])
        cur_q_adu_given.append(adu_units[i])
        cur_q_liab_paid_new.append(owed_v_adu_new[i])
        cur_q_imp_paid_new.append(owed_imp_adu_new[i])
        cur_q_adu_given_new.append(adu_units_new[i])

    return (quar, (begin_liab, begin_imp_owed, begin_adu_req, cur_q_liab, cur_q_imp_owed , cur_q_adu_req, cur_q_liab_paid,\
cur_q_imp_paid, cur_q_adu_given, cur_q_liab_paid_new, cur_q_imp_paid_new, cur_q_adu_given_new,\
ending_liab, ending_imp_owed, ending_adu_req, ending_liab_new, ending_imp_owed_new, ending_adu_req_new))


def get_summary(report_values, date_string, quar):

    wb = load_workbook(filename = DIR_OUTPUT + 'Summary.xlsx')
    ws = wb["Sheet1"]
   
    row_start = ws.max_row + 3
    
    # write date and year+quarter  
    ws.cell(row_start, 2).value = date_string
    ws.cell(row_start, 2).font = Font(bold=True, color=colors.RED)
    ws.cell(row_start, 2).fill = PatternFill("solid", fgColor=colors.YELLOW)
    
    for i in range(2, 6):
        ws.cell(row_start+i, 2).value = quar[i-2]
    
    # headers
    h1 = ['Qtr Begin', 'Qtr Begin', 'Qtr Begin', 'Current Qtr', 'Current Qtr', 'Current Qtr', \
          'Current Qtr', 'Current Qtr', 'Current Qtr', 'Current Qtr', 'Current Qtr', 'Current Qtr', \
          'Qtr End','Qtr End', 'Qtr End', 'Qtr End', 'Qtr End', 'Qtr End']
    h2 = ['Liability', 'Impression Owed', 'ADU required', 'Liability', 'Impression Owed', 'ADU required',\
         'Liaility Paid', 'Impression Paid', 'ADUs Given', 'Liaility Paid(new)','Impression Paid (new sch)', 'ADUs Given(new sch)', \
         'Liability Bal', 'Impression owed', 'ADUs Required','Liaility Bal(new)', 'Impression owed(new sch)', 'ADUs Required(new Sch)']
    
    val_i = 0
    for col_i in range(4, 25):
        if col_i in {7,11,18}:
            pass
        else:
            ws.cell(row_start, col_i).value = h1[val_i]
            ws.cell(row_start, col_i).font = Font(bold=True, underline="single")
            ws.cell(row_start+1, col_i).value = h2[val_i]
            ws.cell(row_start+1, col_i).font = Font(bold=True, underline="single")
        
            for r in range(row_start+2, row_start+6):
                ws.cell(r, col_i).value = report_values[val_i][r-row_start-2]
            val_i += 1
    
    ws.column_dimensions.group(start='O', end='Q', hidden=True)
    ws.column_dimensions.group(start='V', end='X', hidden=True)

    
    wb.save(DIR_OUTPUT+'Summary.xlsx')
    return


def main(Q_num = 2):
    print("Reading Data")
    
    t1 = time.time()
    
    zf = zipfile.ZipFile(DIR_INPUT+'Dealmaker BI weekly reports.zip') 
    df = pd.read_csv(zf.open('Report 1.csv'))
    date = datetime.now()+ dt.timedelta(days=7)
    date_string = str(date.strftime("%m/%d/%Y"))
    startdate = datetime.strptime(date_string, '%m/%d/%Y')
    ratings_file = DIR_INPUT+'2019-07-17 ION Q3 & Q4 Internal Estimates.xlsx'    
    
    quarters = pd.read_csv(DIR_INPUT+'timeList.csv')
    quarters['start_date'] = pd.to_datetime(quarters['start_date']).dt.strftime('%m/%d/%Y')
    four_q = find_quarters(quarters, startdate)
    # Find the start date of each quarter
    quarter_sd = quarter_startdate(quarters, four_q)
    startq = quarter_sd[1] #schedule start date
    endq = quarter_sd[1 + Q_num] # schedule end date
    
    #df['Week Start Date']= pd.to_datetime(df['Week Start Date']) 
    #df = df[df['Week Start Date'] < parse(endq)]
  
    t2 = time.time()
    print('Time for reading files: ', t2 - t1)
    print('Scheduling ADU and generating new data')
    raw = raw_result(df, quarters, date_string, startdate, ratings_file, four_q, startq, endq)

    general = new_data(raw, quarters)

    #changeDF = raw[-1]
    #df['Guarantee Name'] = changeDF[0]
    #df['Selling Title'] = changeDF[1]
    
    new = newdata_to_df(df, general, raw[1])
    t3 = time.time()
    print('Time for scheduling ADU and generating new data: ', t3 - t2)

    print('Calculating liability')
    liab = liability(new)
    t4 = time.time()
    print('Time for computing liability: ', t4 - t3)

    print('Exporting ADU schedule file')
    sep = seperate(raw)
    format_df(sep[0], liab, 'ADU Schedule')
    format_df(sep[1], liab, 'ADU Take Back')
    t5 = time.time()
    print('Time for exporting ADU schedule: ', t5 - t4)
    
    print('Generating suammary')
    quar, report_values = get_report_values(quarters, startdate, liab)
    get_summary(report_values, date_string, quar)
    print('Time for generating summary')
    print('Done')

    print('Total Time: ', t5 - t1)
    return

main()