from operator import index
import sys
import re
import shutil
from pathlib import Path
import time
from datetime import datetime, timedelta
import pandas as pd
pd.options.mode.chained_assignment = None


def write_log_file(pat_log, mes_log):
    print(mes_log)    # message in Terminal
    textfile = open(pat_log, 'a')
    textfile.write(mes_log + '\n')
    textfile.close()


def write_header_logfile(pat_log):
    cur_tim = datetime.now()
    sum_tim = bool(time.localtime(cur_tim.timestamp()).tm_isdst)
    if sum_tim:
        cur_tim = cur_tim - timedelta(hours=1)
    txt = '                                Kanton Uri - Amt für Umweltschutz (AfU)'
    cur_tim = cur_tim.strftime('%d.%m.%Y %H:%M')
    hea = lin + '\n' + cur_tim + txt + '\n' + lin
    write_log_file(pat_log, hea)


def write_footer_logfile(pat_log):
    fil_log = pat_log.name
    txt = fil_log + '                                                       Monitron AG'
    foo = txt + '\n' + lin
    write_log_file(pat_log, foo)


def read_message_csv(pat_mes):
    """read text of messages (error, warning, sucess)"""
    dfr_mes = pd.read_csv(str(pat_mes), delimiter=';')
    dfr_mes = dfr_mes.set_index('id')
    return dfr_mes


def get_file_list_xls(pat_log, dir_inp):
    """return list of xls-files"""
    lis_xls = list(Path(str(dir_inp)).glob('*.xls'))
    if not lis_xls:
        write_log_file(pat_log, dfr_mes.loc[3, lan_mes])
        sys.exit(1)
    return lis_xls


def colToExcel(col):
    """return correct row-index(1-999) and col-index(A-Z)"""
    excelCol = ''
    div = col
    while div:
        (div, mod) = divmod(div - 1, 26)
        excelCol = chr(mod + 65) + excelCol
    return excelCol


def read_fields_csv(pat_fie):
    """read position of input fields (f.e. A1, B3, etc.)"""
    dfr_fie = pd.read_csv(str(pat_fie), delimiter=';')
    dfr_fie = dfr_fie.set_index('id')
    return dfr_fie


def copy_template_csv(pat_tem_sta, pat_tem_str, pat_sta, pat_str):
    """copy template with header and generate empty dataframes"""
    shutil.copy(pat_tem_sta, pat_sta)
    shutil.copy(pat_tem_str, pat_str)
    dfr_sta = pd.read_csv(str(pat_sta), delimiter=';')
    dfr_str = pd.read_csv(str(pat_sta), delimiter=';')
    return dfr_sta, dfr_str


def check_values(pat_log, dfr_res, wri_csv):
    erh_dat = dfr_res.loc['erh_dat', 'result']
    dfr_res.loc['erh_dat', 'result'] = erh_dat.strftime("%d.%m.%Y")
    # dfr_res.loc['phw', 'name'] = 'pH_Wert_Messung'
    gel_nei = dfr_res.loc['gel_nei', 'result']
    match gel_nei:
        case 'schroff':
            gel_nei = 'schroff (90°-30°)'
        case 'stark':
            gel_nei = 'stark (30°-15°)'
        case 'mässig':
            gel_nei = 'mässig (15°-5°)'
        case 'schwach':
            gel_nei = 'schwach (5°-0°)'
        case _:
            gel_nei = ''
    bem = dfr_res.loc['bem', 'result']
    if 'PH' in bem.upper():
        dfr_res.loc['phw', 'result'] = float(re.search("(?i)pH *(\d*\.?\d+)[ |\;|\,]", bem)[1])
    else:
        dfr_res.loc['phw', 'result'] = float('nan')
    que_idn = str(dfr_res.loc['que_idn', 'result'])
    if 'NAN' in que_idn.upper() or 'X' in que_idn.upper():
        write_log_file(pat_log, dfr_mes.loc[4, lan_mes])
        wri_csv = 0
    aus_for = str(dfr_res.loc['aus_for', 'result'])
    if 'NAN' in aus_for.upper():
        write_log_file(pat_log, dfr_mes.loc[5, lan_mes])
        wri_csv = 0
    return dfr_res, wri_csv


def cross_check_fields(dfr_res, cou, i, gro_nam, pat_log, wri_csv):
    """check Quelle nicht bewertbar against other fields"""
    n_b_zer = str(dfr_res.loc['n_b_zer', 'result'])
    n_b_k_a = str(dfr_res.loc['n_b_k_a', 'result'])
    wea = dfr_res.loc['wea', 'result']
    web = dfr_res.loc['web', 'result']
    que_n_b = 0
    if n_b_zer.upper() == 'X' or n_b_k_a.upper() == 'X':
        que_n_b = 1
    if cou == 0 and que_n_b == 0:
        write_log_file(pat_log, dfr_mes.loc[6, lan_mes] + ' ' + gro_nam)
        wri_csv = 0
    elif cou > 1:
        write_log_file(pat_log, dfr_mes.loc[7, lan_mes] + ' ' + gro_nam)
        wri_csv = 0
    elif i == 3 and que_n_b == 1 and not pd.isna(wea):
        write_log_file(pat_log, dfr_mes.loc[8, lan_mes])
        wri_csv = 0
    elif i == 4 and que_n_b == 1 and not pd.isna(web):
        write_log_file(pat_log, dfr_mes.loc[9, lan_mes])
        wri_csv = 0
    return wri_csv


def check_single_choice(pat_fie, pat_log, dfr_res, wri_csv):
    """check single-choice groups 1, 2, 3, 4 from fields.csv"""
    dfr_sin = pd.read_csv(str(pat_fie), delimiter=';')
    dfr_sin = dfr_sin.set_index('single_choice')
    idx_max = int(dfr_sin.index.max() + 1)
    for i in range(1, idx_max):
        # series containing id of 1st, 2nd, ... group
        ser_sin = dfr_sin.loc[i, 'id']
        # group name = first token of name (f.e. 'Vernetzung')
        gro_nam = dfr_sin.loc[i, 'name'].iloc[0].partition('_')[0]
        nam_typ = dfr_sin.loc[i, 'name_typ'].iloc[0]
        zus_typ = dfr_sin.loc[i, 'zustand_typ'].iloc[0]
        nam = ''
        cou = 0
        for cur in ser_sin:
            nam = dfr_res.loc[cur, 'name']
            res = dfr_res.loc[cur, 'result']
            if res == 'x' or res == 'X':
                kat_nam = nam.partition('_')[2]
                # dfr_res.loc[nam_typ, 'name'] = gro_nam
                dfr_res.loc[nam_typ, 'result'] = kat_nam
                cou += 1
            elif res == 1:
                kat_nam = nam.split('_')[1]
                dfr_res.loc[nam_typ, 'result'] = kat_nam
                if not kat_nam == 'keine':
                    zus_nam = nam.split('_')[2]
                    dfr_res.loc[zus_typ, 'result'] = zus_nam
                # dfr_res.loc[nam_typ, 'name'] = gro_nam
                # dfr_res.loc[zus_typ, 'name'] = 'Zustand'
                cou += 1
        wri_csv = cross_check_fields(dfr_res, cou, i, gro_nam, pat_log, wri_csv)
    return dfr_res, wri_csv


def process_struktur_xls(pat_inp, pat_fie, pat_log, dir_che, dfr_fie):
    """read defined fields of xls file and return to dfr_res"""
    # write csv - 1...yes (default) / 2...no (FEHLER in message.csv)
    wri_csv = 1
    write_log_file(pat_log, pat_inp.name)
    dfr_str = pd.read_excel(str(pat_inp), sheet_name='Q_Bewertung_Struktur_D', header=None)
    dfr_str = dfr_str.rename(index=lambda x: x + 1, columns=lambda y: colToExcel(y + 1))
    dfr_res = dfr_fie[['name']]
    for ind, row in dfr_fie.iterrows():
        # skip last lines with empty fields
        if not pd.isna(dfr_fie.loc[ind, 'line']):
            dfr_res.loc[ind, 'result'] = dfr_str.loc[row['line'], row['column']]
    dfr_res, wri_csv = check_values(pat_log, dfr_res, wri_csv)
    dfr_res, wri_csv = check_single_choice(pat_fie, pat_log, dfr_res, wri_csv)
    fil_che = pat_inp.stem + '.csv'
    pat_che = dir_che / fil_che
    dfr_res.to_csv(pat_che, sep=';')
    write_log_file(pat_log, lin)
    return dfr_res, wri_csv


def write_standort_csv(pat_sta, dfr_sta, dfr_fie, dfr_res):
    """write csv-file for Wiski Import - Reiter Standort"""
    # new_row: dic with 'QU_AUSTRITTSFORM': 'Tümpelquelle' etc.
    new_row = pd.Series({dfr_fie.loc['que_idn', 'wiski_1']: dfr_res.loc['que_idn', 'result'],
                         dfr_fie.loc['aus_for', 'wiski_1']: dfr_res.loc['aus_for', 'result'],
                         dfr_fie.loc['han_lag', 'wiski_1']: dfr_res.loc['han_lag', 'result'],
                         dfr_fie.loc['abf_ric', 'wiski_1']: dfr_res.loc['abf_ric', 'result'],
                         dfr_fie.loc['gel_nei', 'wiski_1']: dfr_res.loc['gel_nei', 'result'],
                         dfr_fie.loc['fas_kei', 'wiski_1']: dfr_res.loc['fas_typ', 'result'],
                         dfr_fie.loc['fas_kei', 'wiski_2']: dfr_res.loc['fas_zus', 'result']})
    # append to empty dfr_sta
    dfr_sta = pd.concat([dfr_sta, new_row.to_frame().T], ignore_index=True)
    dfr_sta.to_csv(pat_sta, sep=';', mode='a', index=False, header=False)


def write_struktur_csv(pat_str, dfr_str, dfr_fie, dfr_res):
    """write csv-file for Wiski Import - Reiter Struktur"""
    new_row = pd.Series({dfr_fie.loc['que_idn', 'wiski_1']: dfr_res.loc['que_idn', 'result'],
                         dfr_fie.loc['auf_per', 'wiski_1']: dfr_res.loc['auf_per', 'result'],
                         dfr_fie.loc['auf_dat', 'wiski_1']: dfr_res.loc['auf_dat', 'result'],
                         dfr_fie.loc['que_gro', 'wiski_1']: dfr_res.loc['que_gro', 'result'],
                         dfr_fie.loc['que_ber', 'wiski_1']: dfr_res.loc['que_ber', 'result'],
                         dfr_fie.loc['que_lae', 'wiski_1']: dfr_res.loc['que_lae', 'result'],
                         dfr_fie.loc['mit_fli', 'wiski_1']: dfr_res.loc['mit_fli', 'result'],
                         dfr_fie.loc['qus_jah', 'wiski_1']: dfr_res.loc['qus_jah', 'result'],
                         dfr_fie.loc['ver_typ', 'wiski_1']: dfr_res.loc['ver_typ', 'result'],
                         dfr_fie.loc['dis_nac', 'wiski_1']: dfr_res.loc['dis_nac', 'result'],
                         dfr_fie.loc['anz_aus', 'wiski_1']: dfr_res.loc['anz_aus', 'result'],
                         dfr_fie.loc['str_bem', 'wiski_1']: dfr_res.loc['str_bem', 'result'],
                         dfr_fie.loc['str_wea', 'wiski_1']: dfr_res.loc['str_wea', 'result'],
                         dfr_fie.loc['str_web', 'wiski_1']: dfr_res.loc['str_web', 'result'],
                         dfr_fie.loc['str_weg', 'wiski_1']: dfr_res.loc['str_weg', 'result'],
                         # Weiter        
                         dfr_fie.loc['', 'wiski_2']: dfr_res.loc['', 'result']})
    # append to empty dfr_sta
    dfr_sta = pd.concat([dfr_sta, new_row.to_frame().T], ignore_index=True)
    dfr_sta.to_csv(pat_sta, sep=';', mode='a', index=False, header=False)


global lin
global dfr_mes
global lan_mes
lin = '---------------------------------------------------------------------------------------'
# define language here - later in .ini-file with configparse
lan_mes = 'messageGerman'
dir_cur = Path().resolve()
dir_ini = dir_cur / 'ini'
dir_inp = dir_cur / 'input'
dir_out = dir_cur / 'output'
dir_che = dir_out / 'check'
dir_log = dir_out / 'log'
dir_tem = dir_ini / 'templates'
pat_fie = dir_ini / 'fields.csv'
pat_mes = dir_ini / 'message.csv'
pat_tem_sta = dir_tem / 'KK_standort.csv'
pat_tem_str = dir_tem / 'KK_struktur.csv'
pat_sta = dir_out / 'KK_standort.csv'
pat_str = dir_out / 'KK_struktur.csv'
pat_log = dir_log / 'log_file.txt'
pat_log.unlink(missing_ok=True)
write_header_logfile(pat_log)
dfr_mes = read_message_csv(pat_mes)
dfr_fie = read_fields_csv(pat_fie)
dfr_sta, dfr_str = copy_template_csv(pat_tem_sta, pat_tem_str, pat_sta, pat_str)
lis_xls = get_file_list_xls(pat_log, dir_inp)
for pat_inp in lis_xls:
    dfr_res, wri_csv = process_struktur_xls(pat_inp, pat_fie, pat_log, dir_che, dfr_fie)
    if wri_csv:
        write_standort_csv(pat_sta, dfr_sta, dfr_fie, dfr_res)
        write_struktur_csv(pat_str, dfr_str, dfr_fie, dfr_res)
# write_footer_logfile(pat_log)

# -------------------------------------------------------------------------------
# WEITER - TODO !!!
# single choice x (bringt nichts) - ich brauche result (z.B. Einzelquelle) für Eintrag csv
# 1. token weglassen (delimiter _) z.B. Klassierung_bedingt naturnah  (d.h. 2. Unterstrich in fields weglöschen !!!)
# output: richtiges KK_standort.csv generieren, etc.

# Kosmetik: .ini-file with language path absolut, etc. (configparse)
# chained assignment nicht deaktivieren
# write csv
# only if no FEHLER (message.csv) occurs
