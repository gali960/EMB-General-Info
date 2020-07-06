import cx_Oracle
import xlsxwriter
from datetime import datetime, timedelta
import os
import pandas as pd
import win32com.client as win32

# Dependencies: Operator, Config Matrix Values, Liveries,
# Maintenix Data, Reconfigurations of seats, ovens, galleys, A-CKS, C-CKS intervals,
# Connection to intranet (e.g. via VPN)

# Defining valid aircraft list
aircraft_list = ['HP-1540CMP', 'HP-1556CMP', 'HP-1557CMP', 'HP-1558CMP', 'HP-1559CMP', 'HP-1560CMP', 'HP-1561CMP',
                 'HP-1562CMP', 'HP-1563CMP', 'HP-1564CMP', 'HP-1565CMP', 'HP-1567CMP', 'HP-1568CMP', 'HP-1569CMP']

# Requesting and validating correct aircraft input
while True:
    aircraft = input('Type a valid ERJ-190 aircraft registration in format HP-XXXXCMP: ')
    if aircraft_list.count(aircraft) == 1:
        print(f'Generating General Info Report for {aircraft}...')
        break

# Date of report generation
today = datetime.today().strftime('%d-%b-%y')

# Oracle SQL Queries Definitions
query_ac_id = f'''
SELECT INV_AC_REG.AC_REG_CD, INV_INV.MANUFACT_DT, INV_INV.SERIAL_NO_OEM,
EQP_PART_NO.PART_NO_OEM AS AC_MODEL

FROM INV_AC_REG

INNER JOIN INV_INV ON
INV_AC_REG.INV_NO_ID = INV_INV.INV_NO_ID

INNER JOIN EQP_PART_NO ON
INV_INV.PART_NO_ID = EQP_PART_NO.PART_NO_ID

WHERE AC_REG_CD = '{aircraft}'
'''

query_ac_times = f'''
SELECT 
INV_CURR_USAGE.TSN_QT, INV_CURR_USAGE.DATA_TYPE_ID /*1 = FH, 10 = FC*/

FROM INV_AC_REG 

INNER JOIN INV_CURR_USAGE  ON 
INV_AC_REG.INV_NO_ID = INV_CURR_USAGE.INV_NO_ID 

WHERE AC_REG_CD = '{aircraft}'
'''
query_main_assys = f'''
SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT,
II.CONFIG_POS_SDESC,
INV_CURR_USAGE.DATA_TYPE_ID,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    

INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND (EQP_BOM_PART.BOM_PART_CD = '71-00-00-00' /*ENGINES*/
OR EQP_BOM_PART.BOM_PART_CD = '49-10-00-00' /*APU*/

OR EQP_BOM_PART.BOM_PART_CD = '32-21-00-01' /*NLG LOCKING STAY*/
OR EQP_BOM_PART.BOM_PART_CD = '32-21-01-02' /*NLG SHOCK STRUT*/
OR EQP_BOM_PART.BOM_PART_CD = '32-21-03-01' /*NLG DRAG BRACE*/

OR EQP_BOM_PART.BOM_PART_CD = '32-11-00-01-1' /*MLG LH LOCKING STAY*/
OR EQP_BOM_PART.BOM_PART_CD = '32-11-01-01A' /*MLG LH SHOCK STRUT*/
OR EQP_BOM_PART.BOM_PART_CD = '32-11-05-01' /*MLG LH DRAG BRACE*/

OR EQP_BOM_PART.BOM_PART_CD = '32-11-00-01-5' /*MLG RH LOCKING STAY*/
OR EQP_BOM_PART.BOM_PART_CD = '32-11-01-01B' /*MLG RH SHOCK STRUT*/
OR EQP_BOM_PART.BOM_PART_CD = '32-11-05-05' /*MLG RH DRAG BRACE*/)

'''

# Connecting to Maintenix Oracle Database
dsn_tns = cx_Oracle.makedsn('maintenixdb-test.somoscopa.com', '1521', service_name='COPAT')
conn = cx_Oracle.connect(user='MX_TEST', password='MXT35T2016', dsn=dsn_tns)

# Executing queries and storing in pandas dataframes

df_ac_id = pd.read_sql(query_ac_id, con=conn)
df_ac_times = pd.read_sql(query_ac_times, con=conn)
df_main_assys = pd.read_sql(query_main_assys, con=conn)

# Getting aircraft ID information (AIRFRAME SECTION)
ac_model = 'ERJ 190-100 IGW'
ac_rg = aircraft
msn = df_ac_id['SERIAL_NO_OEM'][0]
man_date = (df_ac_id['MANUFACT_DT'][0]).strftime('%d-%b-%y')
filt_fh = df_ac_times['DATA_TYPE_ID'] == 1
filt_fc = df_ac_times['DATA_TYPE_ID'] == 10
ac_tsn_fh = int((df_ac_times[filt_fh])['TSN_QT'])
ac_tsn_fc = int((df_ac_times[filt_fc])['TSN_QT'])

# Main Assys info

# Engine L/H
filt_eng_lh = df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (LH)'
filt_eng_lh_tsn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (LH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_eng_lh_csn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (LH)') & (df_main_assys['DATA_TYPE_ID'] == 10)

eng_lh_sn = int(df_main_assys[filt_eng_lh]['SERIAL_NO_OEM'].values[0])
eng_lh_tsn = int(df_main_assys[filt_eng_lh_tsn]['TSN_QT'].values[0])
eng_lh_csn = int(df_main_assys[filt_eng_lh_csn]['TSN_QT'].values[0])

# Engine R/H
filt_eng_rh = df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (RH)'
filt_eng_rh_tsn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (RH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_eng_rh_csn = (df_main_assys['CONFIG_POS_SDESC'] == '71-00-00-00 (RH)') & (df_main_assys['DATA_TYPE_ID'] == 10)

eng_rh_sn = int(df_main_assys[filt_eng_rh]['SERIAL_NO_OEM'].values[0])
eng_rh_tsn = int(df_main_assys[filt_eng_rh_tsn]['TSN_QT'].values[0])
eng_rh_csn = int(df_main_assys[filt_eng_rh_csn]['TSN_QT'].values[0])

# APU
filt_apu_sn = (df_main_assys['CONFIG_POS_SDESC'] == '49-10-00-00')
filt_apu_aot = (df_main_assys['CONFIG_POS_SDESC'] == '49-10-00-00') & (df_main_assys['DATA_TYPE_ID'] == 101017)
filt_apu_acyc = (df_main_assys['CONFIG_POS_SDESC'] == '49-10-00-00') & (df_main_assys['DATA_TYPE_ID'] == 101018)

apu_sn = df_main_assys[filt_apu_sn]['SERIAL_NO_OEM'].values[0]
apu_aot = int(df_main_assys[filt_apu_aot]['TSN_QT'].values[0])
apu_acyc = int(df_main_assys[filt_apu_acyc]['TSN_QT'].values[0])


# NLG
filt_nlg_strut_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-01-02') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_nlg_brace_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-03-01') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_nlg_lock_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-00-01') & (df_main_assys['DATA_TYPE_ID'] == 1)

filt_nlg_strut_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-01-02') & (df_main_assys['DATA_TYPE_ID'] == 10)
filt_nlg_brace_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-03-01') & (df_main_assys['DATA_TYPE_ID'] == 10)
filt_nlg_lock_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-21-00-01') & (df_main_assys['DATA_TYPE_ID'] == 10)

nlg_strut_pn = df_main_assys[filt_nlg_strut_fh]['PART_NO_OEM'].values[0]
nlg_strut_sn = df_main_assys[filt_nlg_strut_fh]['SERIAL_NO_OEM'].values[0]
nlg_strut_fh = int(df_main_assys[filt_nlg_strut_fh]['TSN_QT'].values[0])
nlg_strut_fc = int(df_main_assys[filt_nlg_strut_fc]['TSN_QT'].values[0])

nlg_brace_pn = df_main_assys[filt_nlg_brace_fh]['PART_NO_OEM'].values[0]
nlg_brace_sn = df_main_assys[filt_nlg_brace_fh]['SERIAL_NO_OEM'].values[0]
nlg_brace_fh = int(df_main_assys[filt_nlg_brace_fh]['TSN_QT'].values[0])
nlg_brace_fc = int(df_main_assys[filt_nlg_brace_fc]['TSN_QT'].values[0])

nlg_lock_pn = df_main_assys[filt_nlg_lock_fh]['PART_NO_OEM'].values[0]
nlg_lock_sn = df_main_assys[filt_nlg_lock_fh]['SERIAL_NO_OEM'].values[0]
nlg_lock_fh = int(df_main_assys[filt_nlg_lock_fh]['TSN_QT'].values[0])
nlg_lock_fc = int(df_main_assys[filt_nlg_lock_fc]['TSN_QT'].values[0])

# MLG L/H
filt_mlg_lh_strut_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01A (LH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_mlg_lh_side_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-05-01') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_mlg_lh_lock_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-00-01-1') & (df_main_assys['DATA_TYPE_ID'] == 1)

filt_mlg_lh_strut_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01A (LH)') & (df_main_assys['DATA_TYPE_ID'] == 10)
filt_mlg_lh_side_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-05-01') & (df_main_assys['DATA_TYPE_ID'] == 10)
filt_mlg_lh_lock_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-00-01-1') & (df_main_assys['DATA_TYPE_ID'] == 10)

mlg_lh_strut_pn = df_main_assys[filt_mlg_lh_strut_fh]['PART_NO_OEM'].values[0]
mlg_lh_strut_sn = df_main_assys[filt_mlg_lh_strut_fh]['SERIAL_NO_OEM'].values[0]
mlg_lh_strut_fh = int(df_main_assys[filt_mlg_lh_strut_fh]['TSN_QT'].values[0])
mlg_lh_strut_fc = int(df_main_assys[filt_mlg_lh_strut_fc]['TSN_QT'].values[0])

mlg_lh_lock_pn = df_main_assys[filt_mlg_lh_lock_fh]['PART_NO_OEM'].values[0]
mlg_lh_lock_sn = df_main_assys[filt_mlg_lh_lock_fh]['SERIAL_NO_OEM'].values[0]
mlg_lh_lock_fh = int(df_main_assys[filt_mlg_lh_lock_fh]['TSN_QT'].values[0])
mlg_lh_lock_fc = int(df_main_assys[filt_mlg_lh_lock_fc]['TSN_QT'].values[0])

mlg_lh_side_pn = df_main_assys[filt_mlg_lh_side_fh]['PART_NO_OEM'].values[0]
mlg_lh_side_sn = df_main_assys[filt_mlg_lh_side_fh]['SERIAL_NO_OEM'].values[0]
mlg_lh_side_fh = int(df_main_assys[filt_mlg_lh_side_fh]['TSN_QT'].values[0])
mlg_lh_side_fc = int(df_main_assys[filt_mlg_lh_side_fc]['TSN_QT'].values[0])

# MLG R/H
filt_mlg_rh_strut_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01B (RH)') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_mlg_rh_side_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-05-05') & (df_main_assys['DATA_TYPE_ID'] == 1)
filt_mlg_rh_lock_fh = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-00-01-5 (2)') & (df_main_assys['DATA_TYPE_ID'] == 1)

filt_mlg_rh_strut_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-01-01B (RH)') & (df_main_assys['DATA_TYPE_ID'] == 10)
filt_mlg_rh_side_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-05-05') & (df_main_assys['DATA_TYPE_ID'] == 10)
filt_mlg_rh_lock_fc = (df_main_assys['CONFIG_POS_SDESC'] == '32-11-00-01-5 (2)') & (df_main_assys['DATA_TYPE_ID'] == 10)

mlg_rh_strut_pn = df_main_assys[filt_mlg_rh_strut_fh]['PART_NO_OEM'].values[0]
mlg_rh_strut_sn = df_main_assys[filt_mlg_rh_strut_fh]['SERIAL_NO_OEM'].values[0]
mlg_rh_strut_fh = int(df_main_assys[filt_mlg_rh_strut_fh]['TSN_QT'].values[0])
mlg_rh_strut_fc = int(df_main_assys[filt_mlg_rh_strut_fc]['TSN_QT'].values[0])

mlg_rh_lock_pn = df_main_assys[filt_mlg_rh_lock_fh]['PART_NO_OEM'].values[0]
mlg_rh_lock_sn = df_main_assys[filt_mlg_rh_lock_fh]['SERIAL_NO_OEM'].values[0]
mlg_rh_lock_fh = int(df_main_assys[filt_mlg_rh_lock_fh]['TSN_QT'].values[0])
mlg_rh_lock_fc = int(df_main_assys[filt_mlg_rh_lock_fc]['TSN_QT'].values[0])

mlg_rh_side_pn = df_main_assys[filt_mlg_rh_side_fh]['PART_NO_OEM'].values[0]
mlg_rh_side_sn = df_main_assys[filt_mlg_rh_side_fh]['SERIAL_NO_OEM'].values[0]
mlg_rh_side_fh = int(df_main_assys[filt_mlg_rh_side_fh]['TSN_QT'].values[0])
mlg_rh_side_fc = int(df_main_assys[filt_mlg_rh_side_fc]['TSN_QT'].values[0])

# Setting up output excel file
user_path = os.environ['USERPROFILE']
filename = f'GENERAL INFO {aircraft}.xlsx'
location = os.path.join(user_path, 'Documents', f'{filename}')

if os.path.isfile(location):  # Delete if filename already exists
    os.remove(location)
else:
    pass

workbook = xlsxwriter.Workbook(location)
worksheet = workbook.add_worksheet()
worksheet.set_margins(0.5, 0.5, 0.5, 0.5)
footer = 'Rodolfo Naranjo | Records and Technical Publications Manager | Copa Airlines'
worksheet.set_footer(footer)

# Setting cells formatting
cell_format1 = workbook.add_format({'bold': True, 'underline': True})  # Title format

# Setting Copa Logo
image_path = r'C:\Users\ggalina\General Info Report\Logo.png'
worksheet.insert_image('A1', image_path, {'object_position': 3, 'y_offset': 1})

cell_format20 = workbook.add_format({'bold': True})
cell_format20.set_font_size(13)

cell_format21 = workbook.add_format()
cell_format21.set_font_size(13)
cell_format21.set_num_format('#,##0')

worksheet.write('B6', today, cell_format20)


worksheet.set_column('H:H', 11.9)
worksheet.set_column('E:E', 12.7)
worksheet.set_column('F:F', 9)
worksheet.set_column('G:G', 10)

cell_format1.set_font_size(13)

worksheet.write('G7', 'Reg.:', cell_format1)
worksheet.write('G8', 'MSN:',cell_format1)
worksheet.write('G9', 'Model:', cell_format1)
worksheet.write('G10', 'Mfr.Date:', cell_format1)

worksheet.write('H7', ac_rg, cell_format20)
worksheet.write('H8', msn, cell_format20)
worksheet.write('H9', ac_model, cell_format20)
worksheet.write('H10', man_date, cell_format20)

worksheet.write('D12', f'TOTAL TIMES MSN {msn}', cell_format1)
worksheet.write('D13', f'TOTAL HOURS:   {ac_tsn_fh:,}', cell_format21)
worksheet.write('D14', f'TOTAL CYCLES:   {ac_tsn_fc:,}', cell_format21)

worksheet.write('D16', f'ESN1 {eng_lh_sn}', cell_format1)
worksheet.write('D17', f'TOTAL HOURS:   {eng_lh_tsn:,}', cell_format21)
worksheet.write('D18', f'TOTAL CYCLES:   {eng_lh_csn:,}', cell_format21)

worksheet.write('D20', f'ESN2 {eng_rh_sn}', cell_format1)
worksheet.write('D21', f'TOTAL HOURS:   {eng_rh_tsn:,}', cell_format21)
worksheet.write('D22', f'TOTAL CYCLES:   {eng_rh_csn:,}', cell_format21)

worksheet.write('D24', f'APU SN {apu_sn}', cell_format1)
worksheet.write('D25', f'TOTAL HOURS:   {apu_aot:,}', cell_format21)
worksheet.write('D26', f'TOTAL CYCLES:   {apu_acyc:,}', cell_format21)

cell_format5 = workbook.add_format({'bold': True, 'underline': True, 'align': 'center', 'font_size': 13})
cell_format51 = workbook.add_format({'bold': True, 'align': 'center', 'font_size': 13, 'border': True})
cell_format52 = workbook.add_format({'bold': True, 'align': 'center', 'font_size': 11, 'border': True})
cell_format53 = workbook.add_format({'align': 'center', 'font_size': 11, 'border': True})
cell_format54 = workbook.add_format({'align': 'center', 'font_size': 11, 'border': True, 'num_format': '#,##0'})

# Nose Landing Gear Table

worksheet.merge_range('B28:H28', 'Nose Landing Gear', cell_format5)

worksheet.merge_range('B29:D29', 'PART DESCRIPTION', cell_format51)
worksheet.write('E29', 'PN', cell_format51)
worksheet.write('F29', 'SN', cell_format51)
worksheet.write('G29', 'HOURS', cell_format51)
worksheet.write('H29', 'CYCLES', cell_format51)

worksheet.merge_range('B30:D30', 'NLG SHOCK STRUT', cell_format52)
worksheet.merge_range('B31:D31', 'NLG DRAG BRACE ASSY', cell_format52)
worksheet.merge_range('B32:D32', 'NLG LOCKING STAY', cell_format52)

worksheet.write('E30', nlg_strut_pn, cell_format53)
worksheet.write('F30', nlg_strut_sn, cell_format53)
worksheet.write('G30', nlg_strut_fh, cell_format54)
worksheet.write('H30', nlg_strut_fc, cell_format54)

worksheet.write('E31', nlg_brace_pn, cell_format53)
worksheet.write('F31', nlg_brace_sn, cell_format53)
worksheet.write('G31', nlg_brace_fh, cell_format54)
worksheet.write('H31', nlg_brace_fc, cell_format54)

worksheet.write('E32', nlg_lock_pn, cell_format53)
worksheet.write('F32', nlg_lock_sn, cell_format53)
worksheet.write('G32', nlg_lock_fh, cell_format54)
worksheet.write('H32', nlg_lock_fc, cell_format54)

# LH Main Landing Gear Table

worksheet.merge_range('B34:H34', 'LH Main Landing Gear', cell_format5)

worksheet.merge_range('B35:D35', 'PART DESCRIPTION', cell_format51)
worksheet.write('E35', 'PN', cell_format51)
worksheet.write('F35', 'SN', cell_format51)
worksheet.write('G35', 'HOURS', cell_format51)
worksheet.write('H35', 'CYCLES', cell_format51)

worksheet.merge_range('B36:D36', 'SHOCK STRUT ASSY', cell_format52)
worksheet.merge_range('B37:D37', 'MLG LOCKING STAY', cell_format52)
worksheet.merge_range('B38:D38', 'MLG SIDE STAY', cell_format52)

worksheet.write('E36', mlg_lh_strut_pn, cell_format53)
worksheet.write('F36', mlg_lh_strut_sn, cell_format53)
worksheet.write('G36', mlg_lh_strut_fh, cell_format54)
worksheet.write('H36', mlg_lh_strut_fc, cell_format54)

worksheet.write('E37', mlg_lh_lock_pn, cell_format53)
worksheet.write('F37', mlg_lh_lock_sn, cell_format53)
worksheet.write('G37', mlg_lh_lock_fh, cell_format54)
worksheet.write('H37', mlg_lh_lock_fc, cell_format54)

worksheet.write('E38', mlg_lh_side_pn, cell_format53)
worksheet.write('F38', mlg_lh_side_sn, cell_format53)
worksheet.write('G38', mlg_lh_side_fh, cell_format54)
worksheet.write('H38', mlg_lh_side_fc, cell_format54)

# RH Main Landing Gear Table

worksheet.merge_range('B40:H40', 'RH Main Landing Gear', cell_format5)

worksheet.merge_range('B41:D41', 'PART DESCRIPTION', cell_format51)
worksheet.write('E41', 'PN', cell_format51)
worksheet.write('F41', 'SN', cell_format51)
worksheet.write('G41', 'HOURS', cell_format51)
worksheet.write('H41', 'CYCLES', cell_format51)

worksheet.merge_range('B42:D42', 'SHOCK STRUT ASSY', cell_format52)
worksheet.merge_range('B43:D43', 'MLG LOCKING STAY', cell_format52)
worksheet.merge_range('B44:D44', 'MLG SIDE STAY', cell_format52)

worksheet.write('E42', mlg_rh_strut_pn, cell_format53)
worksheet.write('F42', mlg_rh_strut_sn, cell_format53)
worksheet.write('G42', mlg_rh_strut_fh, cell_format54)
worksheet.write('H42', mlg_rh_strut_fc, cell_format54)

worksheet.write('E43', mlg_rh_lock_pn, cell_format53)
worksheet.write('F43', mlg_rh_lock_sn, cell_format53)
worksheet.write('G43', mlg_rh_lock_fh, cell_format54)
worksheet.write('H43', mlg_rh_lock_fc, cell_format54)

worksheet.write('E44', mlg_rh_side_pn, cell_format53)
worksheet.write('F44', mlg_rh_side_sn, cell_format53)
worksheet.write('G44', mlg_rh_side_fh, cell_format54)
worksheet.write('H44', mlg_rh_side_fc, cell_format54)

# Closing workbook
workbook.close()

# Opening file
os.system(f'"{location}"')