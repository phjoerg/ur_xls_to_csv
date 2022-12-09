from operator import index
import sys
import re
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
    txt = '                                      Kanton Uri - Amt für Umweltschutz'
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


def get_file_list_xls(dir_inp):
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


def check_values(dfr_res):
    erh_dat = dfr_res.loc['erh_dat', 'result']
    dfr_res.loc['erh_dat', 'result'] = erh_dat.strftime("%d.%m.%Y")
    dfr_res.loc['phw', 'name'] = 'pH_Wert_Messung'
    bem = dfr_res.loc['bem', 'result']
    if 'PH' in bem.upper():
        dfr_res.loc['phw', 'result'] = float(re.search("(?i)pH *(\d*\.?\d+)[ |\;|\,]", bem)[1])
    else:
        dfr_res.loc['phw', 'result'] = float('nan')
    que_idn = str(dfr_res.loc['que_idn', 'result'])
    if 'NAN' in que_idn.upper() or 'X' in que_idn.upper():
        write_log_file(pat_log, dfr_mes.loc[4, lan_mes])
    return dfr_res


def cross_check_fields(dfr_res, cou, i, gro_nam, pat_log):
    """check Quelle nicht bewertbar against other fields"""
    n_b_zer = str(dfr_res.loc['n_b_zer', 'result'])
    n_b_k_a = str(dfr_res.loc['n_b_k_a', 'result'])
    wea = dfr_res.loc['wea', 'result']
    web = dfr_res.loc['web', 'result']
    que_n_b = 0
    if n_b_zer.upper() == 'X' or n_b_k_a.upper() == 'X':
        que_n_b = 1
    if cou == 0 and que_n_b == 0:
        write_log_file(pat_log, dfr_mes.loc[5, lan_mes] + ' ' + gro_nam)
    elif cou > 1:
        write_log_file(pat_log, dfr_mes.loc[6, lan_mes] + ' ' + gro_nam)
    elif i == 3 and que_n_b == 1 and not pd.isna(wea):
        write_log_file(pat_log, dfr_mes.loc[7, lan_mes])
    elif i == 4 and que_n_b == 1 and not pd.isna(web):
        write_log_file(pat_log, dfr_mes.loc[8, lan_mes])


def check_single_choice(pat_fie, pat_log, dfr_res):
    """check single-choice groups"""
    dfr_sin = pd.read_csv(str(pat_fie), delimiter=';')
    dfr_sin = dfr_sin.set_index('single_choice')
    idx_max = int(dfr_sin.index.max() + 1)
    for i in range(1, idx_max):
        # series containing id of 1st, 2nd, ... group
        ser_sin = dfr_sin.loc[i, 'id']
        # group name = first token of name (f.e. 'Vernetzung')
        gro_nam = dfr_sin.loc[i, 'name'].iloc[0].partition('_')[0]
        res = ''
        cou = 0
        for cur in ser_sin:
            mem = dfr_res.loc[cur, 'result']
            if mem == 'x' or mem == 'X' or mem == 1:
                # res = dfr_res.loc[cur, 'name']      res (string) ev. nec. for csv ?
                cou += 1
        cross_check_fields(dfr_res, cou, i, gro_nam, pat_log)


def process_spring_report_xls(pat_inp, pat_fie, dfr_fie):
    """read defined fields of xls file and return to dfr_res"""
    write_log_file(pat_log, pat_inp.name)
    dfr_str = pd.read_excel(str(pat_inp), sheet_name='Q_Bewertung_Struktur_D', header=None)
    dfr_str = dfr_str.rename(index=lambda x: x + 1, columns=lambda y: colToExcel(y + 1))
    dfr_res = dfr_fie[['name']]
    for ind, row in dfr_fie.iterrows():
        dfr_res.loc[ind, 'result'] = dfr_str.loc[row['line'], row['column']]
    dfr_res = check_values(dfr_res)
    check_single_choice(pat_fie, pat_log, dfr_res)
    fil_che = pat_inp.stem + '.csv'
    pat_che = dir_che / fil_che
    dfr_res.to_csv(pat_che, sep=';')
    write_log_file(pat_log, lin)


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
pat_fie = dir_ini / 'fields.csv'
pat_mes = dir_ini / 'message.csv'
pat_log = dir_log / 'log_file.txt'
pat_log.unlink(missing_ok=True)
write_header_logfile(pat_log)
dfr_mes = read_message_csv(pat_mes)
dfr_fie = read_fields_csv(pat_fie)
lis_xls = get_file_list_xls(dir_inp)
for pat_inp in lis_xls:
    process_spring_report_xls(pat_inp, pat_fie, dfr_fie)
# write_footer_logfile(pat_log)

# -------------------------------------------------------------------------------
# WEITER - TODO !!!
# .ini-file with language path absolut, etc. (configparse)

# read csv single choice
# single choice - warnungen bestehend einpflegen
# logfile erzeugen immer eine Zeile mit xls ---
# nä. Zeile ok oder Meldungen (Warnungen / Error)
# Simon: Was sind Pflichtfelder - wo soll py Alarm schlagen?
# chained assignment nicht deaktivieren

# split up input fields to several pandas series (row) for each csv wiski output (row)
# first check input fields necessary for csv
# then check input fields for remaining fields

# format Wert A/B, Aufwertung and pH-Wert to 2 NKS (nicht so dringend)


# kla = ''
# kla_cou = 0
# kla_ein = dfr_str.loc[9, 'S']
# kla_sys = dfr_str.loc[9, 'U']
# kla_kom = dfr_str.loc[9, 'W']
# if kla_ein.upper() == 'X':
#     kla = 'Einzelquelle'
#     kla_cou += 1
# if kla_sys.upper() == 'X':
#     kla = 'Quellsystem'
#     kla_cou += 1
# if kla_kom.upper() == 'X':
#     kla = 'Quellkomplex'
#     kla_cou += 1
# if kla_cou == 0:
#     sys.exit('ERROR - fehlendes X bei Vernetzung')
# elif kla_cou > 1:
#     sys.exit('ERROR - Mehrfachauswahl bei Vernetzung')
#
# fot_dok = dfr_str.loc[17, 'Q']
# fot_idn = dfr_str.loc[17, 'S']
# if fot_dok == 'x' or fot_dok == 'X':
#     fot_dok = True
# else:
#     fot_dok = False
#     if fot_idn:
#         sys.exit('ERROR - fehlendes X bei Fotos / Dokumente')
# nut = ''
# nut_tri = dfr_str.loc[19, 'Q']
# nut_sch = dfr_str.loc[19, 'S']
# nut_kul = dfr_str.loc[19, 'W']
# if nut_tri.upper() == 'X':
#     nut = 'Trinkwassernutzung'
# if nut_sch.upper() == 'X':
#     if nut:
#         nut = nut + ', Schutzstatus'
#     else:
#         nut = 'Schutzstatus'
# if nut_kul.upper() == 'X':
#     if nut:
#         nut = nut + ', Kulturhistorische Bedeutung'
#     else:
#         nut = 'Kulturhistorische Bedeutung'
# # Protokoll - Fuss -------------------------------------------------------------------------------------------------
# rev_obj = dfr_str.loc[124, 'G']
# if rev_obj.upper() == 'JA':
#     rev_obj = True
# elif rev_obj.upper() == 'NEIN':
#     rev_obj = False
# else:
#     sys.exit('ERROR - fehlende Eingabe (J/N) bei Revitalisierungsobjekt')
# kla = ''
# kla_cou = 0
# kla_nat = dfr_str.loc[127, 'G']
# kla_bed = dfr_str.loc[128, 'G']
# kla_mae = dfr_str.loc[129, 'G']
# kla_ges = dfr_str.loc[130, 'G']
# kla_sta = dfr_str.loc[131, 'G']
# if kla_nat.upper() == 'X':
#     kla = 'naturnah'
#     kla_cou += 1
# if kla_bed.upper() == 'X':
#     kla = 'bedingt naturnah'
#     kla_cou += 1
# if kla_mae.upper() == 'X':
#     kla = 'mässig beeinträchtigt'
#     kla_cou += 1
# if kla_ges.upper() == 'X':
#     kla = 'geschädigt'
#     kla_cou += 1
# if kla_sta.upper() == 'X':
#     kla = 'stark geschädigt'
#     kla_cou += 1
# if kla_cou == 0:
#     sys.exit('ERROR - fehlendes X bei Vernetzung')
# elif kla_cou > 1:
#     sys.exit('ERROR - Mehrfachauswahl bei Vernetzung')
# bew = ''
# bew_1__ = dfr_str.loc[127, 'M']
# bew_2__ = dfr_str.loc[128, 'M']
# bew_3__ = dfr_str.loc[129, 'M']
# bew_4__ = dfr_str.loc[130, 'M']
# bew_5__ = dfr_str.loc[131, 'M']
# # WEITER !!! if... keine Mehrfachauswahl und kein Leerlassen erlaubt
# # def für mehrfachausw., etc. NICHT erlaubt mit dic kla_nat = {dfr_str.loc[127, 'G'], 'naturnah'}
# # weiteres für mehrfachausw., etc. erlaubt
# n_b = ''
# n_b_zer = dfr_str.loc[129, 'U']
# n_b_k_A = dfr_str.loc[131, 'U']
# # if...  Mehrfachauswahl und Leerlassen erlaubt

# wer_a__ = float(dfr_str.loc[122, 'I'])
# wer_b__ = float(dfr_str.loc[122, 'V'])
# wer_auf = float(dfr_str.loc[124, 'V'])
# wer_ges = float(dfr_str.loc[126, 'V'])

# print('xxx')
