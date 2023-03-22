from operator import index
import os
import sys
import re
import shutil
from pathlib import Path
import time
from datetime import datetime, timedelta
import pandas as pd
pd.options.mode.chained_assignment = None


def write_log_file(pat_log, mes_log):
    # message in Terminal
    # print(mes_log)
    textfile = open(pat_log, 'a')
    textfile.write(mes_log + '\n')
    textfile.close()


def write_header_logfile(pat_log):
    cur_tim = datetime.now()
    sum_tim = bool(time.localtime(cur_tim.timestamp()).tm_isdst)
    if sum_tim:
        cur_tim = cur_tim - timedelta(hours=1)
    txt = '                                  Kanton Uri - Amt für Umweltschutz (AfU)'
    cur_tim = cur_tim.strftime('%d.%m.%Y %H:%M')
    hea = lin + '\n' + cur_tim + txt + '\n' + lin
    write_log_file(pat_log, hea)


def write_footer_logfile(pat_log):
    fil_log = pat_log.name
    txt = fil_log + '                                                                  Monitron AG'
    foo = txt + '\n' + lin
    write_log_file(pat_log, foo)


def read_message_csv(pat_mes):
    """read text of messages (error, warning, sucess)"""
    dfr_mes = pd.read_csv(str(pat_mes), delimiter=';')
    dfr_mes = dfr_mes.set_index('id')
    return dfr_mes


def get_list_of_input_files(pat_log, dir_inp, ext_xls):
    """return list of xls-files with extension"""
    lis_xls = list(Path(str(dir_inp)).glob(ext_xls))
    if not lis_xls:
        write_log_file(pat_log, dfr_mes.loc[3, lan_mes])
        sys.exit(1)
    num_inp = len(lis_xls)
    return lis_xls, num_inp


def remove_files_in_dir(pat_log, dir_tar, fil_ext):
    """remove files with extension .xyz in target directory"""
    if not dir_tar.exists():
        write_log_file(pat_log, dfr_mes.loc[13, lan_mes] + str(dir_tar))
        sys.exit(1)
    lis = list(Path(str(dir_tar)).glob(fil_ext))
    for fil in lis:
        fil.unlink()


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
    """copy templates standort, struktur, ... and generate empty dataframes"""
    shutil.copy(pat_tem_sta, pat_sta)
    shutil.copy(pat_tem_str, pat_str)
    dfr_sta = pd.read_csv(str(pat_sta), delimiter=';')
    dfr_str = pd.read_csv(str(pat_str), delimiter=';')
    return dfr_sta, dfr_str


def check_mandatory_fields(pat_fie, pat_log, dfr_res, wri_csv):
    """check mandatory fields acc. to fields.csv"""
    dfr = pd.read_csv(str(pat_fie), delimiter=';', usecols=['name', 'id', 'mandatory_field'])
    # exctract mandatory fields
    dfr = dfr[~dfr['mandatory_field'].isna()]
    for ind, row in dfr.iterrows():
        if 'NAN' in str(dfr_res.loc[row['id'], 'result']).upper():
            write_log_file(pat_log, dfr_mes.loc[5, lan_mes] + ' ' + row['name'])
            wri_csv = 0
    return dfr_res, wri_csv


def rename_fields(dfr_res):
    """Rename fields (Geländeneigung / mittl. Fliessgeschw."""
    # Geländeneigung
    match dfr_res.loc['gel_nei', 'result']:
        case 'schroff':
            dfr_res.loc['gel_nei', 'result'] = 'schroff (30 - 90°)'
        case 'stark':
            dfr_res.loc['gel_nei', 'result'] = 'stark (15 - 30°)'
        case 'mässig':
            dfr_res.loc['gel_nei', 'result'] = 'mässig (5 - 15°)'
        case 'schwach':
            dfr_res.loc['gel_nei', 'result'] = 'schwach (0 - 5°)'
        case _:
            dfr_res.loc['gel_nei', 'result'] = ''
    # mittl. Fliessgeschw.
    match dfr_res.loc['mit_fli', 'result']:
        case 'sehr schnell':
            dfr_res.loc['mit_fli', 'result'] = 'sehr schnell (> 1.5 m/s)'
        case 'schnell':
            dfr_res.loc['mit_fli', 'result'] = 'schnell (0.75 - 1.5 m/s)'
        case 'mässig':
            dfr_res.loc['mit_fli', 'result'] = 'mässig (0.25 - 0.75 m/s)'
        case 'langsam':
            dfr_res.loc['mit_fli', 'result'] = 'langsam (0.05 - 0.25 m/s)'
        case 'stehend':
            dfr_res.loc['mit_fli', 'result'] = 'stehend (< 0.05 m/s)'
        case _:
            dfr_res.loc['mit_fli', 'result'] = ''
    return dfr_res


def check_values(pat_log, dfr_res, wri_csv):
    """Extract pH-Wert, check Quellen-ID, evaluate Quelle nicht bewertbar"""
    auf_dat = dfr_res.loc['auf_dat', 'result']
    dfr_res.loc['auf_dat', 'result'] = auf_dat.strftime("%d.%m.%Y")
    # pH-Wert
    str_bem = dfr_res.loc['str_bem', 'result']
    if 'PH' in str_bem.upper():
        dfr_res.loc['phw', 'result'] = float(re.search("(?i)pH *(\d*\.?\d+)[ |\;|\,]", str_bem)[1])
    else:
        dfr_res.loc['phw', 'result'] = float('nan')
    # Quellen-ID
    que_idn = str(dfr_res.loc['que_idn', 'result'])
    if 'NAN' in que_idn.upper() or 'X' in que_idn.upper():
        write_log_file(pat_log, dfr_mes.loc[4, lan_mes])
        wri_csv = 0
    # Quelle nicht bewertbar
    n_b_zer = str(dfr_res.loc['n_b_zer', 'result'])
    n_b_k_a = str(dfr_res.loc['n_b_k_a', 'result'])
    if n_b_zer.upper() == 'X' and n_b_k_a.upper() == 'X':
        write_log_file(pat_log, dfr_mes.loc[10, lan_mes])
        que_n_b = 1
        wri_csv = 0
    elif n_b_zer.upper() == 'X' and n_b_k_a == 'nan':
        dfr_res.loc['qnb_typ', 'result'] = dfr_res.loc['n_b_zer', 'name'].partition('_')[2]
        que_n_b = 1
    elif n_b_k_a.upper() == 'X' and n_b_zer == 'nan':
        dfr_res.loc['qnb_typ', 'result'] = dfr_res.loc['n_b_k_a', 'name'].partition('_')[2]
        que_n_b = 1
    else:
        que_n_b = 0
    return dfr_res, wri_csv, que_n_b


def cross_check_fields(dfr_res, cou, i, gro_nam, pat_log, wri_csv, que_n_b):
    """check Quelle nicht bewertbar que_n_b against other fields"""
    str_wea = dfr_res.loc['str_wea', 'result']
    str_web = dfr_res.loc['str_web', 'result']
    if cou == 0 and que_n_b == 0:
        write_log_file(pat_log, dfr_mes.loc[6, lan_mes] + ' ' + gro_nam)
        wri_csv = 0
    elif cou > 1:
        write_log_file(pat_log, dfr_mes.loc[7, lan_mes] + ' ' + gro_nam)
        wri_csv = 0
    elif i == 3 and que_n_b == 1 and not pd.isna(str_wea):
        write_log_file(pat_log, dfr_mes.loc[8, lan_mes])
        wri_csv = 0
    elif i == 4 and que_n_b == 1 and not pd.isna(str_web):
        write_log_file(pat_log, dfr_mes.loc[9, lan_mes])
        wri_csv = 0
    return wri_csv


def check_single_choice(pat_fie, pat_log, dfr_res, que_n_b, wri_csv):
    """check single-choice groups 1, 2, 3, 4 from fields.csv"""
    dfr = pd.read_csv(str(pat_fie), delimiter=';')
    dfr = dfr.set_index('single_choice')
    idx_max = int(dfr.index.max() + 1)
    for i in range(1, idx_max):
        # series containing id of 1st, 2nd, ... 4th group
        ser_sin = dfr.loc[i, 'id']
        # group name = first token of name (f.e. 'Vernetzung')
        gro_nam = dfr.loc[i, 'name'].iloc[0].partition('_')[0]
        nam_typ = dfr.loc[i, 'name_typ'].iloc[0]
        zus_typ = dfr.loc[i, 'zustand_typ'].iloc[0]
        nam = ''
        cou = 0
        for cur in ser_sin:
            nam = dfr_res.loc[cur, 'name']
            res = dfr_res.loc[cur, 'result']
            if res == 'x' or res == 'X':
                kat_nam = nam.partition('_')[2]
                dfr_res.loc[nam_typ, 'result'] = kat_nam
                cou += 1
            elif res == 1:
                kat_nam = nam.split('_')[1]
                dfr_res.loc[nam_typ, 'result'] = kat_nam
                if not kat_nam == 'keine':
                    zus_nam = nam.split('_')[2]
                    dfr_res.loc[zus_typ, 'result'] = zus_nam
                cou += 1
        wri_csv = cross_check_fields(dfr_res, cou, i, gro_nam, pat_log, wri_csv, que_n_b)
    return dfr_res, wri_csv


def analyse_protocol_xls(pat_inp, pat_fie, pat_log, dir_che, dfr_fie):
    """read defined fields of xls file and return to dfr_res"""
    # write csv - 1...yes (default) / 2...no (FEHLER acc. to message.csv)
    wri_csv = 1
    write_log_file(pat_log, pat_inp.name)
    dfr = pd.read_excel(str(pat_inp), sheet_name='Q_Bewertung_Struktur_D', header=None)
    dfr = dfr.rename(index=lambda x: x + 1, columns=lambda y: colToExcel(y + 1))
    dfr_res = dfr_fie[['name']]
    for ind, row in dfr_fie.iterrows():
        # skip last lines with empty fields
        if not pd.isna(dfr_fie.loc[ind, 'line']):
            dfr_res.loc[ind, 'result'] = dfr.loc[row['line'], row['column']]
    dfr_res, wri_csv = check_mandatory_fields(pat_fie, pat_log, dfr_res, wri_csv)
    dfr_res = rename_fields(dfr_res)
    dfr_res, wri_csv, que_n_b = check_values(pat_log, dfr_res, wri_csv)
    dfr_res, wri_csv = check_single_choice(pat_fie, pat_log, dfr_res, que_n_b, wri_csv)
    fil_che = pat_inp.stem + '.csv'
    pat_che = dir_che / fil_che
    dfr_res.to_csv(pat_che, sep=';', encoding='ansi')
    write_log_file(pat_log, lin)
    return dfr_res, wri_csv


def write_standort_csv(pat_sta, dfr_sta, dfr_fie, dfr_res):
    """write csv-file for Wiski Import - Reiter Standort"""
    # new_row: dic with 'QU_AUSTRITTSFORM': 'Tümpelquelle' etc.
    new_row = pd.Series({dfr_fie.loc['que_idn', 'wiski']: dfr_res.loc['que_idn', 'result'],
                         dfr_fie.loc['aus_for', 'wiski']: dfr_res.loc['aus_for', 'result'],
                         dfr_fie.loc['han_lag', 'wiski']: dfr_res.loc['han_lag', 'result'],
                         dfr_fie.loc['abf_ric', 'wiski']: dfr_res.loc['abf_ric', 'result'],
                         dfr_fie.loc['gel_nei', 'wiski']: dfr_res.loc['gel_nei', 'result'],
                         dfr_fie.loc['fas_typ', 'wiski']: dfr_res.loc['fas_typ', 'result'],
                         dfr_fie.loc['fas_zus', 'wiski']: dfr_res.loc['fas_zus', 'result']})
    # append to empty dfr_sta
    dfr_sta = pd.concat([dfr_sta, new_row.to_frame().T], ignore_index=True)
    dfr_sta.to_csv(pat_sta, sep=';', mode='a', index=False, header=False, encoding='ansi')


def write_struktur_csv(pat_str, dfr_str, dfr_fie, dfr_res):
    """write csv-file for Wiski Import - Reiter Struktur"""
    new_row = pd.Series({dfr_fie.loc['que_idn', 'wiski']: dfr_res.loc['que_idn', 'result'],
                         dfr_fie.loc['auf_bea', 'wiski']: dfr_res.loc['auf_bea', 'result'],
                         dfr_fie.loc['auf_dat', 'wiski']: dfr_res.loc['auf_dat', 'result'],
                         dfr_fie.loc['que_gro', 'wiski']: dfr_res.loc['que_gro', 'result'],
                         dfr_fie.loc['que_ber', 'wiski']: dfr_res.loc['que_ber', 'result'],
                         dfr_fie.loc['que_lae', 'wiski']: dfr_res.loc['que_lae', 'result'],
                         dfr_fie.loc['mit_fli', 'wiski']: dfr_res.loc['mit_fli', 'result'],
                         dfr_fie.loc['qus_jah', 'wiski']: dfr_res.loc['qus_jah', 'result'],
                         dfr_fie.loc['ver_typ', 'wiski']: dfr_res.loc['ver_typ', 'result'],
                         dfr_fie.loc['dis_nac', 'wiski']: dfr_res.loc['dis_nac', 'result'],
                         dfr_fie.loc['anz_aus', 'wiski']: dfr_res.loc['anz_aus', 'result'],
                         dfr_fie.loc['str_bem', 'wiski']: dfr_res.loc['str_bem', 'result'],
                         dfr_fie.loc['str_wea', 'wiski']: dfr_res.loc['str_wea', 'result'],
                         dfr_fie.loc['str_web', 'wiski']: dfr_res.loc['str_web', 'result'],
                         dfr_fie.loc['str_weg', 'wiski']: dfr_res.loc['str_weg', 'result'],
                         dfr_fie.loc['bew_typ', 'wiski']: dfr_res.loc['bew_typ', 'result'],
                         dfr_fie.loc['ges_typ', 'wiski']: dfr_res.loc['ges_typ', 'result'],
                         dfr_fie.loc['rev_obj', 'wiski']: dfr_res.loc['rev_obj', 'result'],
                         dfr_fie.loc['qnb_typ', 'wiski']: dfr_res.loc['qnb_typ', 'result']})
    # append to empty dfr_str
    dfr_str = pd.concat([dfr_str, new_row.to_frame().T], ignore_index=True)
    dfr_str.to_csv(pat_str, sep=';', mode='a', index=False, header=False, encoding='ansi')


def remove_blank_line_csv(pat_sta):
    """remove last 2 character CR and LF to get rid of blank line at end of file"""
    with open(pat_sta, "rb+") as fil:
        fil.seek(-2, os.SEEK_END)
        fil.truncate()
        fil.close()


global lin
global dfr_mes
global lan_mes
lin = '-----------------------------------------------------------------------------------------'
# define language here
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
ext_xls = '*.xls'
ext_csv = '*.csv'


write_header_logfile(pat_log)
dfr_mes = read_message_csv(pat_mes)
remove_files_in_dir(pat_log, dir_che, ext_csv)
dfr_fie = read_fields_csv(pat_fie)
dfr_sta, dfr_str = copy_template_csv(pat_tem_sta, pat_tem_str, pat_sta, pat_str)
lis_xls, num_inp = get_list_of_input_files(pat_log, dir_inp, ext_xls)
cou_csv = 0
for pat_inp in lis_xls:
    dfr_res, wri_csv = analyse_protocol_xls(pat_inp, pat_fie, pat_log, dir_che, dfr_fie)
    if wri_csv:
        write_standort_csv(pat_sta, dfr_sta, dfr_fie, dfr_res)
        write_struktur_csv(pat_str, dfr_str, dfr_fie, dfr_res)
        cou_csv += 1
if cou_csv == 0:
    write_log_file(pat_log, dfr_mes.loc[12, lan_mes])
    remove_files_in_dir(pat_log, dir_out, ext_csv)
else:
    write_log_file(pat_log, dfr_mes.loc[11, lan_mes] + str(cou_csv))
    remove_blank_line_csv(pat_sta)
    remove_blank_line_csv(pat_str)
write_log_file(pat_log, lin)
# write_footer_logfile(pat_log)
print('\n---- START -------------------------------------------------')
print('Anzahl Strukturprotokolle in (Ordner input):  ' + str(num_inp))
print('Anzahl Einträge in KK_...csv (Ordner output): ' + str(cou_csv))
print('Anzahl fehlerhafte Protokolle (Ordner log):   ' + str(num_inp - cou_csv))
print('---- ENDE -------------------------------------------------\n')


# ---------------------------------------------------------------------------------------------
# TODO !!!

# Fehlermeld. Win durch eigene (verständliche Fehlermeldung)ersetzen
# z.B. wenn xls geöffnet ist und man es löschen will:
# [WinError 32] The process cannot access the file because it is being used by another process:
# '\\\\chdcrfile2\\HOME_BUR\\BUR\\THU\\MTPJO\\98_Ausbildung\\03_python_PJo\\box\\ur_xls_to_csv\\output\\check\\KK_standort_test_ok.csv'

# oder wenn man geöffnetes xls überschreiben:
# [Errno 13] Permission denied:
# '\\\\chdcrfile2\\HOME_BUR\\BUR\\THU\\MTPJO\\98_Ausbildung\\03_python_PJo\\box\\ur_xls_to_csv\\output\\KK_standort.csv'

# function rename_fields(dfr_res)
# Bezeichnungen in csv oder in dic auslagern

# Kosmetik: .ini-file with language=german, path to input / output, etc. (configparse)
# chained assignment nicht deaktivieren
