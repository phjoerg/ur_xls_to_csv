from operator import index
import sys
import pandas as pd
import numpy as np
import re
from pathlib import Path


def get_file_list_xls(dir_inp):
    """return list of xls-files"""
    lis_xls = list(Path(str(dir_inp)).glob('*.xls'))
    if not lis_xls:
        sys.exit('ERROR - Keine xlsx-Dateien in Input-klazeichnis')
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
    """read csv-table with row / col of input fields"""
    dfr_fie = pd.read_csv(str(pat_fie), delimiter=';')
    dfr_fie = dfr_fie.set_index('id')
    return dfr_fie


def read_file_xls(pat_inp, dfr_fie):
    """read defined fields of xls file and return to dfr_res and output/check/csv"""
    dfr_str = pd.read_excel(str(pat_inp), sheet_name='Q_Bewertung_Struktur_D', header=None)
    dfr_str = dfr_str.rename(index=lambda x: x + 1, columns=lambda y: colToExcel(y + 1))
    dfr_res = dfr_fie[['name']]
    dfr_res['result'] = pd.Series()
    for ind, row in dfr_fie.iterrows():
        res = dfr_str.loc[row['line'], row['column']]
        dfr_res['result'][ind] = res
    fil_che = pat_inp.stem + '.csv'
    pat_che = dir_che / fil_che
    dfr_res.to_csv(pat_che, sep=';')
    return dfr_res

    # pro_nam = dfr_str.loc[1, 'A']
    # if not pro_nam.__contains__('Struktur'):
    #     sys.exit('ERROR - Keine deutsches Quell-Protokoll')
    # que_ktn = dfr_str.loc[1, 'R']
    # que_idn = dfr_str.loc[1, 'U']
    # que_nam = dfr_str.loc[3, 'C']
    # erh_dat = dfr_str.loc[3, 'N'].strftime("%d.%m.%Y")
    # koo_x__ = dfr_str.loc[3, 'T']
    # koo_y__ = dfr_str.loc[3, 'V']
    # flu_nam = dfr_str.loc[5, 'C']
    # hoe_ueM = dfr_str.loc[5, 'N']
    # bea_nam = dfr_str.loc[5, 'T']
    # aus_for = dfr_str.loc[9, 'D']
    # han_lag = dfr_str.loc[11, 'D']
    # abf_ric = dfr_str.loc[13, 'D']
    # gel_nei = dfr_str.loc[15, 'D']
    # que_sch = dfr_str.loc[17, 'D']
    # mit_fli = dfr_str.loc[19, 'D']
    # que_gro = dfr_str.loc[9, 'J']
    # que_ber = dfr_str.loc[11, 'J']
    # que_lae = dfr_str.loc[13, 'J']
    # was_tem = dfr_str.loc[15, 'J']
    # que_l_s = dfr_str.loc[17, 'J']
    # lei_fae = dfr_str.loc[19, 'J']
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
    # dis_nac = dfr_str.loc[11, 'R']
    # anz_aus = dfr_str.loc[11, 'V']
    # phw = dfr_str.loc[13, 'Q']
    # phw = float(re.search("(?i)pH *(\d*\.?\d+)[ |\;|\,]", phw)[1])
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


print('**** START **********************************************************************************')

dir_cur = Path().resolve()

dir_inp = dir_cur / 'input'
dir_out = dir_cur / 'output'
dir_che = dir_out / 'check'
pat_fie = dir_inp / 'fields.csv'
dfr_fie = read_fields_csv(pat_fie)
lis_xls = get_file_list_xls(dir_inp)
for pat_inp in lis_xls:
    dfr_res = read_file_xls(pat_inp, dfr_fie)
    print('xxx')

print('**** END ************************************************************************************')

# TODO
# input csv for row col number of xlsx
# split up input fields to several pandas series (row) for each csv output (row)
# first check input fields necessary for csv
# then check input fields for remaining fields

# pro_nam = dfr_str.loc[1, 'A']
#     que_ktn = dfr_str.loc[1, 'R']
#     que_idn = dfr_str.loc[1, 'U']
#     que_nam = dfr_str.loc[3, 'C']
#     erh_dat = dfr_str.loc[3, 'N']
#     kox = dfr_str.loc[3, 'T']
#     koy = dfr_str.loc[3, 'V']
#     flu_nam = dfr_str.loc[5, 'C']
#     hoe_ueM = dfr_str.loc[5, 'N']
#     bea_nam = dfr_str.loc[5, 'T']
#     aus_for = dfr_str.loc[9, 'D']
#     han_lag = dfr_str.loc[11, 'D']
#     abf_ric = dfr_str.loc[13, 'D']
#     gel_nei = dfr_str.loc[15, 'D']
#     qus_zei = dfr_str.loc[17, 'D']
#     mit_fli = dfr_str.loc[19, 'D']
#     que_gro = dfr_str.loc[9, 'J']
#     que_ber = dfr_str.loc[11, 'J']
#     que_lae = dfr_str.loc[13, 'J']
#     was_tem = dfr_str.loc[15, 'J']
#     qus_wer = dfr_str.loc[17, 'J']
#     lei_fae = dfr_str.loc[19, 'J']
#     kla_ein = dfr_str.loc[9, 'S']
#     kla_sys = dfr_str.loc[9, 'U']
#     kla_kom = dfr_str.loc[9, 'W']
#     dis_nac = dfr_str.loc[11, 'R']
#     anz_aus = dfr_str.loc[11, 'V']
#     phw = dfr_str.loc[13, 'Q']
#     fot_dok = dfr_str.loc[17, 'Q']
#     fot_idn = dfr_str.loc[17, 'S']
#     nut_tri = dfr_str.loc[19, 'Q']
#     nut_sch = dfr_str.loc[19, 'S']
#     nut_kul = dfr_str.loc[19, 'W']
#     rev_obj = dfr_str.loc[124, 'G']
#     kla_nat = dfr_str.loc[127, 'G']
#     kla_bed = dfr_str.loc[128, 'G']
#     kla_mae = dfr_str.loc[129, 'G']
#     kla_ges = dfr_str.loc[130, 'G']
#     kla_sta = dfr_str.loc[131, 'G']
#     bew_1__ = dfr_str.loc[127, 'M']
#     bew_2__ = dfr_str.loc[128, 'M']
#     bew_3__ = dfr_str.loc[129, 'M']
#     bew_4__ = dfr_str.loc[130, 'M']
#     bew_5__ = dfr_str.loc[131, 'M']
#     wer_a__ = dfr_str.loc[122, 'I']
#     wer_b__ = dfr_str.loc[122, 'V']
#     wer_auf = dfr_str.loc[124, 'V']
#     wer_ges = dfr_str.loc[126, 'V']
#     n_b_zer = dfr_str.loc[129, 'U']
#     n_b_k_A = dfr_str.loc[131, 'U']
#     co1 = ['a', 'b', 'c', 'd']
#     df1 = pd.DataFrame([que_idn, que_nam, kox, koy], dtype=str, columns=co1)
# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.concat.html#pandas.concat
