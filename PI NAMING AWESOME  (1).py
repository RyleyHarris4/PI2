#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import glob
import os
from datetime import date
import pyodbc


# In[2]:


LatestFile = max(glob.iglob('C:/Users/ryl38028/OneDrive - Hess Corporation/Docs/PI/Input/*'),key=os.path.getctime)


# In[3]:


print(LatestFile)


# In[4]:


dataframe1 = pd.read_excel(LatestFile)


# In[5]:


iconicsdisplayname = dataframe1.iloc[:,1]


# In[6]:


for col in dataframe1:
    print(col)


# In[7]:



#import getpass
#username = 'ihess\\'+os.environ.get('USERNAME')
#password = getpass.getpass('Password for '+username+': ')


# In[8]:


#server = 'hess-ods.database.windows.net'
#database = 'dw-ods'
#driver= 'ODBC Driver 13 for SQL Server'
#cnxn = pyodbc.connect('DRIVER='+driver+';SERVER=tcp:'+server+';PORT=1433;DATABASE='+database+';UID='+username+';PWD='+ password)
#cnxn = pyodbc.connect('Driver={SQL13};SERVER=hess-ods.database.windows.net;DATABASE=dw-ods;Authentication=ActiveDirectoryIntegrated;Encrypt=yes')


#cursor = cnxn.cursor()
#sql = """

#SELECT 


#					LEFT(ICS1.[Message],25)
#					,LEFT(ICS2.[Message],25)
#					--,LEFT(ICS3.[Message],25)
#					--,LEFT(ICS4.[Message],25)
#					,ODR.[SubMultiWellPad]

#				  FROM  [odr_nd].[OdrWellDaily] ODR
#				  RIGHT JOIN [xspoc].[ICO_AS1_AlarmData_Tran_AlarmLog] ICS1 ON LEFT(ODR.[SubMultiWellPad],25) = LEFT(ICS1.[Message],25)
#				  RIGHT JOIN [xspoc].[ICO_AS2_AlarmData_Tran_AlarmLog] ICS2 ON  LEFT(ICS2.[Message],25) = LEFT(ODR.[SubMultiWellPad],25) 
 #           """

#df = pd.read_sql(sql, cnxn)


# In[9]:


dataframe1['Process'] = np.where(dataframe1['DisplayName'].str.contains('AIT90003'), 'GENERAL UTILITIES', 
                        np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_CURRENT'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_VOLTAGE'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_ENERGY'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_POWER'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_CURRENT'), 'ELECTRICAL SYSTEMS',
                        np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_CURRENT'), 'ELECTRICAL SYSTEMS',
                        np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_VOLTAGE'), 'ELECTRICAL SYSTEMS',
                        np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_VOLTAGE'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_POWER_FACTOR'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_IN'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_OUT'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_POWER'), 'ELECTRICAL SYSTEMS',
                        np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_IN'), 'ELECTRICAL SYSTEMS',
                        np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_OUT'), 'ELECTRICAL SYSTEMS',
                        np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_POWER'), 'ELECTRICAL SYSTEMS', 
                        np.where(dataframe1['DisplayName'].str.contains('PIT90001'), 'GENERAL UTILITIES',
                        np.where(dataframe1['DisplayName'].str.contains('TIT90002'), 'STRUCTURES',
                        np.where(dataframe1['DisplayName'].str.contains('TIT00141'), 'STRUCTURES',
                        np.where(dataframe1['DisplayName'].str.contains('TIT90001'), 'STRUCTURES', 
                        np.where(dataframe1['DisplayName'].str.contains('TIT70800'), 'SALES AND PRD TRANSP',
                        np.where(dataframe1['DisplayName'].str.contains('TIT70801'), 'SALES AND PRD TRANSP',
                        np.where(dataframe1['DisplayName'].str.contains('HDPE_10_Yearly_Accum'), 'SALES AND PRD TRANSP',
                        np.where(dataframe1['DisplayName'].str.contains('HDPE_Yearly_Accum'), 'SALES AND PRD TRANSP', 
                        np.where(dataframe1['DisplayName'].str.contains('PIT00142'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('P00140_P_SPEED'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00140'), 'SALES AND PRD TRANSP',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_TOTAL'), 'SALES AND PRD TRANSP', 
                        np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_YDY'), 'SALES AND PRD TRANSP',
                        np.where(dataframe1['DisplayName'].str.contains('PIT00147'), 'SALES AND PRD TRANSP',
                        np.where(dataframe1['DisplayName'].str.contains('PIT00148'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00141'), 'PRODUCTION', 
                        np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TOTAL'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TDY'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_YDY'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('PIT00141'), 'OIL PROCESSING', 
                        np.where(dataframe1['DisplayName'].str.contains('PIT00140'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('TIT00140'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('PIT00144'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00142'), 'OIL PROCESSING', 
                        np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TOTAL'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TDY'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_YDY'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('PIT90101'), 'GENERAL UTILITIES', 
                        np.where(dataframe1['DisplayName'].str.contains('AIT90101'), 'SAFETY AND CONTROL',
                        np.where(dataframe1['DisplayName'].str.contains('AIT90103'), 'SAFETY AND CONTROL',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00100'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00100_VOL_TOTAL'), 'OIL PROCESSING', 
                        np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_ACTUAL_CMD_FREQ'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_CMD_FREQ'), 'OIL PROCESSING', 
                        np.where(dataframe1['DisplayName'].str.contains('PIT00107'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('PIT00108'), 'OIL PROCESSING', 
                        np.where(dataframe1['DisplayName'].str.contains('FUY00101'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TOTAL'), 'PRODUCTION', 
                        np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TDY'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_YDY'), 'PRODUCTION', 
                        np.where(dataframe1['DisplayName'].str.contains('PIT00101'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('TIT00101'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('PIT00100'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('TIT00100'), 'OIL PROCESSING', 
                        np.where(dataframe1['DisplayName'].str.contains('FUY00102'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TOTAL'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TDY'), 'OIL PROCESSING',
                        np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_YDY'), 'OIL PROCESSING', 
                        np.where(dataframe1['DisplayName'].str.contains('PIT0.101'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('PIT0.101X'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLC'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLO'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('PIT0.100'), 'PRODUCTION', 
                        np.where(dataframe1['DisplayName'].str.contains('PIT0.102'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('PIT0.102X'), 'PRODUCTION',
                        np.where(dataframe1['DisplayName'].str.contains('PIT00143'), 'OIL PROCESSING','NULL'))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))
                                 


# In[10]:


dataframe1['System'] = np.where(dataframe1['DisplayName'].str.contains('AIT90003'), 'INSTRUMENT AIR', 
                       np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_CURRENT'), 'MAIN POWER',
                       np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_VOLTAGE'), 'EMRG PWR DIST', 
                       np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_ENERGY'), 'MAIN POWER', 
                       np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_POWER'), 'MAIN POWER', 
                       np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_CURRENT'), 'MAIN POWER',
                       np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_CURRENT'), 'MAIN POWER', 
                       np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_VOLTAGE'), 'EMRG PWR DIST',
                       np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_VOLTAGE'), 'EMRG PWR DIST', 
                       np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_POWER_FACTOR'), 'MAIN POWER', 
                       np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_IN'), 'MAIN POWER', 
                       np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_OUT'), 'MAIN POWER', 
                       np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_POWER'), 'MAIN POWER',
                       np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_IN'), 'MAIN POWER',
                       np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_OUT'), 'MAIN POWER',
                       np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_POWER'), 'MAIN POWER', 
                       np.where(dataframe1['DisplayName'].str.contains('PIT90001'), 'INSTRUMENT AIR',
                       np.where(dataframe1['DisplayName'].str.contains('TIT90002'), 'BUILDING/ACCOM',
                       np.where(dataframe1['DisplayName'].str.contains('TIT00141'), 'BUILDING/ACCOM',
                       np.where(dataframe1['DisplayName'].str.contains('TIT90001'), 'BUILDING/ACCOM', 
                       np.where(dataframe1['DisplayName'].str.contains('TIT70800'), 'GAS EXPORT',
                       np.where(dataframe1['DisplayName'].str.contains('TIT70801'), 'GAS EXPORT',
                       np.where(dataframe1['DisplayName'].str.contains('HDPE_10_Yearly_Accum'), 'GAS EXPORT',
                       np.where(dataframe1['DisplayName'].str.contains('HDPE_Yearly_Accum'), 'GAS EXPORT', 
                       np.where(dataframe1['DisplayName'].str.contains('PIT00142'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('P00140_P_SPEED'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00140'), 'GAS METERING',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_TOTAL'), 'GAS METERING', 
                       np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_YDY'), 'GAS METERING',
                       np.where(dataframe1['DisplayName'].str.contains('PIT00147'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('PIT00148'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00141'), 'ALLOCATION METER', 
                       np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TOTAL'), 'ALLOCATION METER',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TDY'), 'ALLOCATION METER',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_YDY'), 'ALLOCATION METER',
                       np.where(dataframe1['DisplayName'].str.contains('PIT00141'), 'SEPARATION', 
                       np.where(dataframe1['DisplayName'].str.contains('PIT00140'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('TIT00140'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('PIT00144'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00142'), 'SEPARATION', 
                       np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TOTAL'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TDY'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_YDY'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('PIT90101'), 'INSTRUMENT AIR', 
                       np.where(dataframe1['DisplayName'].str.contains('AIT90101'), 'FIRE&GAS DETECT',
                       np.where(dataframe1['DisplayName'].str.contains('AIT90103'), 'FIRE&GAS DETECT',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00100'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00100_VOL_TOTAL'), 'SEPARATION', 
                       np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_ACTUAL_CMD_FREQ'), 'PRODUCED WATER',
                       np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_CMD_FREQ'), 'PRODUCED WATER',
                       np.where(dataframe1['DisplayName'].str.contains('PIT00107'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('PIT00108'), 'SEPARATION', 
                       np.where(dataframe1['DisplayName'].str.contains('FUY00101'), 'ALLOCATION METER',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TOTAL'), 'ALLOCATION METER', 
                       np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TDY'), 'ALLOCATION METER',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_YDY'), 'ALLOCATION METER', 
                       np.where(dataframe1['DisplayName'].str.contains('PIT00101'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('TIT00101'), 'SEPARATION', 
                       np.where(dataframe1['DisplayName'].str.contains('PIT00100'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('TIT00100'), 'SEPARATION', 
                       np.where(dataframe1['DisplayName'].str.contains('FUY00102'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TOTAL'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TDY'), 'SEPARATION',
                       np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_YDY'), 'SEPARATION', 
                       np.where(dataframe1['DisplayName'].str.contains('PIT0.101'), 'WELLS',
                       np.where(dataframe1['DisplayName'].str.contains('PIT0.101X'), 'WELLS',
                       np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLC'), 'WELLS',
                       np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLO'), 'WELLS',
                       np.where(dataframe1['DisplayName'].str.contains('PIT0.100'), 'WELLS', 
                       np.where(dataframe1['DisplayName'].str.contains('PIT0.102'), 'WELLS',
                       np.where(dataframe1['DisplayName'].str.contains('PIT0.102X'), 'WELLS',
                       np.where(dataframe1['DisplayName'].str.contains('PIT00143'), 'SEPARATION','NULL'))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))


# In[11]:


dataframe1['Equipment'] = np.where(dataframe1['DisplayName'].str.contains('AIT90003'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_CURRENT'), 'Electrical Building',
                          np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_VOLTAGE'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_ENERGY'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_POWER'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_CURRENT'), 'Electrical Building',
                          np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_CURRENT'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_VOLTAGE'), 'Electrical Building',
                          np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_VOLTAGE'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_POWER_FACTOR'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_IN'), 'Electrical Building',
                          np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_OUT'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_POWER'), 'Electrical Building',
                          np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_IN'), 'Electrical Building',
                          np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_OUT'), 'Electrical Building',
                          np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_POWER'), 'Electrical Building', 
                          np.where(dataframe1['DisplayName'].str.contains('PIT90001'), 'NULL',
                          np.where(dataframe1['DisplayName'].str.contains('TIT90002'), 'Compressor Building',
                          np.where(dataframe1['DisplayName'].str.contains('TIT00141'), 'Electric Building',
                          np.where(dataframe1['DisplayName'].str.contains('TIT90001'), 'Electric Building', 
                          np.where(dataframe1['DisplayName'].str.contains('TIT70800'), 'Gas Cooler Inlet',
                          np.where(dataframe1['DisplayName'].str.contains('TIT70801'), 'Gas Cooler Outlet',
                          np.where(dataframe1['DisplayName'].str.contains('HDPE_10_Yearly_Accum'), 'NULL',
                          np.where(dataframe1['DisplayName'].str.contains('HDPE_Yearly_Accum'), 'NULL', 
                          np.where(dataframe1['DisplayName'].str.contains('PIT00142'), 'Separator Emulsion Line',
                          np.where(dataframe1['DisplayName'].str.contains('P00140_P_SPEED'), 'Transfer Pump',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00140'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_TOTAL'), 'Separator Production', 
                          np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_YDY'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('PIT00147'), 'Separator Production 1 Emulsion Line',
                          np.where(dataframe1['DisplayName'].str.contains('PIT00148'), 'Separator Production 1 Emulsion Line',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00141'), 'Separator Production', 
                          np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TOTAL'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TDY'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_YDY'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('PIT00141'), 'Separator Production', 
                          np.where(dataframe1['DisplayName'].str.contains('PIT00140'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('TIT00140'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('PIT00144'), 'Separator Production Transfer Pump',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00142'), 'Separator Production', 
                          np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TOTAL'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TDY'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_YDY'), 'Separator Production',
                          np.where(dataframe1['DisplayName'].str.contains('PIT90101'), 'Separator Building', 
                          np.where(dataframe1['DisplayName'].str.contains('AIT90101'), 'Separator Test Building',
                          np.where(dataframe1['DisplayName'].str.contains('AIT90103'), 'Separator Test Building',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00100'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00100_VOL_TOTAL'), 'Separator Test', 
                          np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_ACTUAL_CMD_FREQ'), 'Pump',
                          np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_CMD_FREQ'), 'Pump', 
                          np.where(dataframe1['DisplayName'].str.contains('PIT00107'), 'Separator Test Line',
                          np.where(dataframe1['DisplayName'].str.contains('PIT00108'), 'Separator Test Line', 
                          np.where(dataframe1['DisplayName'].str.contains('FUY00101'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TOTAL'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TDY'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_YDY'), 'Separator Test', 
                          np.where(dataframe1['DisplayName'].str.contains('PIT00101'), 'Separator Test Line',
                          np.where(dataframe1['DisplayName'].str.contains('TIT00101'), 'Separator Test Line',
                          np.where(dataframe1['DisplayName'].str.contains('PIT00100'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('TIT00100'), 'Separator Test', 
                          np.where(dataframe1['DisplayName'].str.contains('FUY00102'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TOTAL'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TDY'), 'Separator Test',
                          np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_YDY'), 'Separator Test', 
                          np.where(dataframe1['DisplayName'].str.contains('PIT0.101'), 'Casing',
                          np.where(dataframe1['DisplayName'].str.contains('PIT0.101X'), 'Casing',
                          np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLC'), 'Emergency Shutdown Valve',
                          np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLO'), 'Emergency Shutdown Valve',
                          np.where(dataframe1['DisplayName'].str.contains('PIT0.100'), 'Flowline', 
                          np.where(dataframe1['DisplayName'].str.contains('PIT0.102'), 'Tubing',
                          np.where(dataframe1['DisplayName'].str.contains('PIT0.102X'), 'Tubing',
                          np.where(dataframe1['DisplayName'].str.contains('PIT00143'), 'Separator Production','NULL'))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))


# In[12]:


dataframe1['SubEquipment'] = np.where(dataframe1['DisplayName'].str.contains('AIT90003'), 'Air Compressor Dew Point', 
                             np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_CURRENT'), '3 Phase A Current', 
                             np.where(dataframe1['DisplayName'].str.contains('SEL_A_PHASE_VOLTAGE'), '3 Phase A Voltage', 
                             np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_ENERGY'), '3 Phase Apparent Energy',
                             np.where(dataframe1['DisplayName'].str.contains('3PHASE_APPARENT_POWER'), '3 Phase Apparent Power', 
                             np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_CURRENT'), '3 Phase B Current',
                             np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_CURRENT'), '3 Phase C Current',
                             np.where(dataframe1['DisplayName'].str.contains('SEL_B_PHASE_VOLTAGE'), '3 Phase B Voltage',
                             np.where(dataframe1['DisplayName'].str.contains('SEL_C_PHASE_VOLTAGE'), '3 Phase C Voltage', 
                             np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_POWER_FACTOR'), '3 Phase Power Factor', 
                             np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_IN'), '3 Phase Reactive Energy In', 
                             np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_ENERGY_OUT'), '3 Phase Reactive Energy Out', 
                             np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REACTIVE_POWER'), '3 Phase Reactive Power',
                             np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_IN'), '3 Phase Real Energy In',
                             np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_ENERGY_OUT'), '3 Phase Reactive Energy Out',
                             np.where(dataframe1['DisplayName'].str.contains('SEL_3PHASE_REAL_POWER'), '3 Phase Real Power', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT90001'), 'Air Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('TIT90002'), 'Temperature',
                             np.where(dataframe1['DisplayName'].str.contains('TIT00141'), 'Temperature',
                             np.where(dataframe1['DisplayName'].str.contains('TIT90001'), 'Temperature', 
                             np.where(dataframe1['DisplayName'].str.contains('TIT70800'), 'Temperature',
                             np.where(dataframe1['DisplayName'].str.contains('TIT70801'), 'Temperature',
                             np.where(dataframe1['DisplayName'].str.contains('HDPE_10_Yearly_Accum'), 'Sales Gas HDPE Press 10% Yearly Accumulated Time',
                             np.where(dataframe1['DisplayName'].str.contains('HDPE_Yearly_Accum'), 'Sales Gas HDPE Press Yearly Accumulated Time', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT00142'), 'Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('P00140_P_SPEED'), 'Speed Command',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00140'), 'Gas Flow Rate',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_TOTAL'), 'Gas Volume Total', 
                             np.where(dataframe1['DisplayName'].str.contains('FUY00140_VOL_YDY'), 'Gas Volume Yesterday',
                             np.where(dataframe1['DisplayName'].str.contains('PIT00147'), 'Oil Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('PIT00148'), 'Water Discharge Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00141'), 'Oil Flow Rate', 
                             np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TOTAL'), 'Oil Volume Total',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_TDY'), 'Oil Volume Today',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00141_VOL_YDY'), 'Oil Volume Yesterday',
                             np.where(dataframe1['DisplayName'].str.contains('PIT00141'), 'Sales Gas Pressure', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT00140'), 'Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('TIT00140'), 'Temperature',
                             np.where(dataframe1['DisplayName'].str.contains('PIT00144'), 'Discharge Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00142'), 'Water Flow Rate', 
                             np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TOTAL'), 'Water Volume Total',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_TDY'), 'Water Volume Today',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00142_VOL_YDY'), 'Water Volume Yesterday',
                             np.where(dataframe1['DisplayName'].str.contains('PIT90101'), 'Air Pressure', 
                             np.where(dataframe1['DisplayName'].str.contains('AIT90101'), 'LEL Level',
                             np.where(dataframe1['DisplayName'].str.contains('AIT90103'), 'LEL Level',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00100'), 'Gas Flow Rate',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00100_VOL_TOTAL'), 'Gas Volume Total', 
                             np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_ACTUAL_CMD_FREQ'), 'Actual Frequency',
                             np.where(dataframe1['DisplayName'].str.contains('TPWS_P51000_CMD_FREQ'), 'Frequency', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT00107'), 'Oil Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('PIT00108'), 'Water Discharge Pressure', 
                             np.where(dataframe1['DisplayName'].str.contains('FUY00101'), 'Oil Flow Rate',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TOTAL'), 'Oil Volume Total', 
                             np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_TDY'), 'Oil Volume Today',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00101_VOL_YDY'), 'Oil Volume Yesterday', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT00101'), 'Sales Gas Pressuree',
                             np.where(dataframe1['DisplayName'].str.contains('TIT00101'), 'Sales Gas Temperature', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT00100'), 'Pressure',
                             np.where(dataframe1['DisplayName'].str.contains('TIT00100'), 'Temperature', 
                             np.where(dataframe1['DisplayName'].str.contains('FUY00102'), 'Water Flow Rate',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TOTAL'), 'Water Volume Total',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_TDY'), 'Water Volume Today',
                             np.where(dataframe1['DisplayName'].str.contains('FUY00102_VOL_YDY'), 'Water Volume Yesterday', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT0.101'), 'Pressure From ICONICS',
                             np.where(dataframe1['DisplayName'].str.contains('PIT0.101X'), 'Pressure From XSPOC',
                             np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLC'), 'Closed Status',
                             np.where(dataframe1['DisplayName'].str.contains('ESDV0.100_ZLO'), 'Open Status',
                             np.where(dataframe1['DisplayName'].str.contains('PIT0.100'), 'Pressure', 
                             np.where(dataframe1['DisplayName'].str.contains('PIT0.102'), 'Pressure From ICONICS',
                             np.where(dataframe1['DisplayName'].str.contains('PIT0.102X'), 'Pressure From XSPOC',
                             np.where(dataframe1['DisplayName'].str.contains('PIT00143'), 'Discharge Pressure','NULL'))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))


# In[13]:


dataframe1


# In[14]:


current_datetime = date.today()


# In[15]:


str_current_datetime = str(current_datetime)
print(str_current_datetime)


# In[16]:


file_name = 'C:/Users/ryl38028/OneDrive - Hess Corporation/Docs/PI/Output/'+str_current_datetime+".xlsx"


# In[17]:


datatoexcel= pd.ExcelWriter(file_name, engine = 'xlsxwriter')


# In[18]:


dataframe1.to_excel(datatoexcel, index=False)


# In[19]:


datatoexcel.save()


# In[ ]:





# In[ ]:




