# -*- coding: utf-8 -*-
"""
Created on Mon Mar  1 07:59:55 2021

@author: BPS
"""


# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 10:13:41 2020

@author: HP
"""

import numpy as np
import pandas as pd
from pandas import ExcelWriter
from dateutil import parser
from datetime import datetime, timedelta
from os import listdir
from os.path import isfile, join

TANGGAL_LIBUR = []
HARI_LIBUR = ['Sabtu', 'Minggu']

SATPAM = ['M SYUKUR', 'BAGYO', 'SUYITNO', 'MOKHAMAD SYUKUR', 'Yiet']

def tgl_ke_hari(tanggal):
    date = parser.parse(tanggal).strftime("%a")
    if date.lower()=="sun":
        return "Minggu"
    if date.lower()=="mon":
        return "Senin"
    if date.lower()=="tue":
        return "Selasa"
    if date.lower()=="wed":
        return "Rabu"
    if date.lower()=="thu":
        return "Kamis"
    if date.lower()=="fri":
        return "Jumat"
    if date.lower()=="sat":
        return "Sabtu"
    
  
def cek_telat(nama, absen_masuk):
    if nama != "M SYUKUR" :
        if nama != "BAGYO" :
            if nama != "SUYITNO":      
                FMT = '%H:%M:%S'
                cek = datetime.strptime(absen_masuk, FMT) - datetime.strptime("07:30:00", FMT)
                if cek.days>=0:
                    hours, remainder = divmod(cek.seconds, 3600)
                    minutes, seconds = divmod(remainder, 60)  
                    return str(hours) +' jam, '+ str(minutes)+' menit, '+str(seconds)+' detik'
                else:
                    return "-"
            else:
                return "-"
        else:
            return "-"
    else:
        return "-"
    
path_kaizala = 'D://me - 2021//ABSENSI KAIZALA//raw//2021.04'
dfbulan = pd.DataFrame() 
folders = listdir(path_kaizala)
rarfile = []
for i in range(len(folders)):
    if len(listdir(path_kaizala+'//'+folders[i]))>0:
        rarfile.append(path_kaizala+'//'+folders[i]+'//' +listdir(path_kaizala+'//'+folders[i])[0])
i=7

df_month = pd.DataFrame()
for i in range(len(rarfile)):
    zip_path = rarfile[i]
    df = pd.read_csv(rarfile[i], compression='zip', header=0, sep=',', quotechar='"')
    cols_to_check = list(df.columns)
    df[cols_to_check] = df[cols_to_check].replace({'"':''}, regex=True).replace({'=':''}, regex=True)
    df_month = df_month.append(df)

df_month = df_month.reset_index()
df_month['date'] = ''
df_month['cektanggal'] = ''
for i in range(len(df_month)):
    df_month['ResponseTime (UTC)'][i] = datetime.strptime(df_month['ResponseTime (UTC)'][i], '%Y-%m-%d %H:%M:%S %p') + timedelta(hours=7)
    df_month['date'][i] = str(df_month['ResponseTime (UTC)'][i]).split(' ')[0]

df_month['Responder Name'] = df_month['Responder Name'].str.upper()

users = list(set(df_month['Responder Name']))
users.sort()
date = list(set(df_month['date']))
date.sort()
i=27
rekap = pd.DataFrame(columns=['Tanggal','Hari','Nama','Absen Masuk','Lokasi Masuk','Absen Pulang','Lokasi Pulang','Waktu Telat', 'TL', 'SW'])

j=8
for i in range(len(users)):
    df2=df_month[(df_month['Responder Name']==users[i])]    
    df2.sort_values(by=['ResponseTime (UTC)'], inplace=True, ascending=True)
    # df3 = pd.DataFrame()
    for j in range(len(date)):
        df3 = df2[(df2['date']==date[j])].reset_index()
        df3.sort_values(by=['ResponseTime (UTC)'], inplace=True, ascending=True)
        nama = users[i]
        tanggal = date[j]
        hari = tgl_ke_hari(date[j])
        telat_masuk = '-'
        
        
        if(len(df3))==0:
            absen_masuk = '-'
            lok_masuk = '-'
            absen_pulang = '-'
            lok_pulang = '-'
            if tanggal not in TANGGAL_LIBUR and hari not in HARI_LIBUR and nama not in SATPAM:
                TL = '4'
                SW = '4'
            else:
                TL = '-'
                SW = '-'
        if(len(df3))==1:
            if(5<df3['ResponseTime (UTC)'].iloc[-1].hour<11):
                absen_masuk = df3['ResponseTime (UTC)'].iloc[0].strftime("%H:%M:%S")
                lok_masuk = df3['Responder Location Location'].iloc[0]
                absen_pulang = '-'
                lok_pulang = '-'
                if tanggal not in TANGGAL_LIBUR and hari not in HARI_LIBUR and nama not in SATPAM:
                    telat_masuk = cek_telat(users[i], absen_masuk)
                else:
                    telat_masuk = '-'
                TL = '-'
                SW = '-'
                if tanggal not in TANGGAL_LIBUR and telat_masuk != '-' and hari not in HARI_LIBUR and nama not in SATPAM:
                    jm = int(telat_masuk.split(',')[0].split(' ')[0])*60
                    mn = int(telat_masuk.split(',')[1].split(' ')[1])
                    if 0<(jm+mn)<=30:
                        TL = '1'
                    if 30<(jm+mn)<=60:
                        TL = '2'
                    if 60<(jm+mn)<=90:
                        TL = '3'
                    if 90<(jm+mn)<=9999:
                        TL = '4'
                    
                if absen_pulang == '-' and nama not in SATPAM and tanggal not in TANGGAL_LIBUR and hari not in HARI_LIBUR:
                    SW = '4'
                
            if(14<df3['ResponseTime (UTC)'].iloc[-1].hour<23):
                absen_masuk = '-'
                lok_masuk = '-'
                absen_pulang = df3['ResponseTime (UTC)'].iloc[0].strftime("%H:%M:%S")
                lok_pulang = df3['Responder Location Location'].iloc[0]
                if tanggal not in TANGGAL_LIBUR and hari not in HARI_LIBUR:
                    if absen_masuk == '-' and nama not in SATPAM:
                        TL = '4'
                else:
                    TL = '-'
            
                SW = '-'
        if(len(df3))>=2:
            absen_masuk = df3['ResponseTime (UTC)'].iloc[0].strftime("%H:%M:%S")
            lok_masuk = df3['Responder Location Location'].iloc[0]
            absen_pulang = df3['ResponseTime (UTC)'].iloc[-1].strftime("%H:%M:%S")
            lok_pulang = df3['Responder Location Location'].iloc[-1]
            if tanggal not in TANGGAL_LIBUR and hari not in HARI_LIBUR and nama not in SATPAM:
                telat_masuk = cek_telat(users[i], absen_masuk)
            else:
                telat_masuk = '-'
            if tanggal not in TANGGAL_LIBUR and telat_masuk != '-' and hari not in HARI_LIBUR and nama not in SATPAM:
                jm = int(telat_masuk.split(',')[0].split(' ')[0])*60
                mn = int(telat_masuk.split(',')[1].split(' ')[1])
                if 0<(jm+mn)<=30:
                    TL = '1'
                if 30<(jm+mn)<=60:
                    TL = '2'
                if 60<(jm+mn)<=90:
                    TL = '3'
                if 90<(jm+mn)<=9999:
                    TL = '4'
            else:
                TL = '-'
            SW = '-'
        rekap = rekap.append({'Nama': nama,
                              'Hari': hari, 
                              'Tanggal': tanggal, 
                              'Absen Masuk': absen_masuk, 
                              'Lokasi Masuk': lok_masuk, 
                              'Absen Pulang': absen_pulang, 
                              'Lokasi Pulang': lok_pulang,
                              'Keterangan Telat': telat_masuk,
                              'TL' : TL,
                              'SW' : SW,
                           },ignore_index=True)
    rekap = rekap.append({'Nama': ' ','Hari': ' ','Tanggal': ' ','Absen Masuk': ' ','Lokasi Masuk': ' ','Absen Pulang': ' ','Lokasi Pulang': ' ',},ignore_index=True)    
    rekap = rekap.append({'Nama': ' ','Hari': ' ','Tanggal': ' ','Absen Masuk': ' ','Lokasi Masuk': ' ','Absen Pulang': ' ','Lokasi Pulang': ' ',},ignore_index=True)    
                                                                                             
writer = ExcelWriter('D://me - 2021//ABSENSI KAIZALA//hasil//April 2021.xlsx')
rekap.to_excel(writer, index=0)
writer.save()

