#!usr/bin/env python3
# -*- coding:utf-8 -*-

import pandas as pd
from prettytable import PrettyTable as Pr

table_src = '20_years_of_gaming_export.xls'


def get_wb(source):
    workbook = pd.read_excel(source, sheet_name='tsvexport_eipRKY')  # reading worksheet of workbook
    prof_estimate = workbook['Профф. Критика'].to_list()   # read column to list
    gamers_estimate = workbook['Сред. пользовательская оценка'].to_list()  # read next column to list
    return mark_lst_gam(gamers_estimate, prof_estimate)


def mark_lst_gam(gam, prof):
    okay = 0
    amazing = 0
    masterpiece = 0
    great = 0
    good = 0
    awful = 0
    disaster = 0
    unbearable = 0
    mediocre = 0
    bad = 0
    painful = 0
    y = 0
    n = 0
    for i in gam:
        """Add values to lists of gamers estimate"""
        if i == 'Okay':     # iterate the gamers estimate list for searching  an estimate value
            okay = okay+1
        elif i == 'Amazing':
            amazing = amazing+1
        elif i == 'Great':
            great = great+1
        elif i == 'Good':
            good = good+1
        elif i == 'Masterpiece':
            masterpiece = masterpiece+1
        elif i == 'Awful':
            awful = awful+1
        elif i == 'Disaster':
            disaster = disaster+1
        elif i == 'Unbearable':
            unbearable = unbearable+1
        elif i == 'Mediocre':
            mediocre = mediocre+1
        elif i == 'Bad':
            bad = bad+1
        elif i == 'Painful':
            painful = painful+1
    for j in prof:
        """Add values to reviewers estimate lists"""
        if j == 'Y':
            y = y+1
        else:
            n = n+1
    """Calculate to rounded percentage of relation gamers and reviewers estimates"""
    total_neg_gam = str(round((awful+disaster+unbearable+mediocre+bad+painful)/(len(gam)/100), 2))+' %'
    total_pos_gam = str(round((okay+amazing+great+good+masterpiece) / (len(gam)/100), 2))+' %'
    total_neg_pro = str(round(n/(len(prof)/100), 2))+' %'
    total_pos_pro = str(round(y/(len(prof)/100), 2))+' %'
    percent_pos_rel_gam = round(((okay+amazing+great+good+masterpiece)
                                - (awful+disaster+unbearable+mediocre+bad+painful))
                                / (awful+disaster+unbearable+mediocre+bad+painful), 2)
    percent_rev_relation_estimate = round(((n/(len(prof)/100))-(y/(len(prof)/100)))/(y/(len(prof)/100)), 2)
    st_rel_gam = str(round(((okay+amazing+great+good+masterpiece)
                            - (awful+disaster+unbearable+mediocre+bad+painful))
                           / (awful+disaster+unbearable+mediocre+bad+painful), 2))+' %'
    total_relation_percent = str(round((percent_rev_relation_estimate-percent_pos_rel_gam)/percent_pos_rel_gam, 2))+' %'
    perc_rev_rel_est = str(round(((n/(len(prof)/100))-(y/(len(prof)/100)))/(y/(len(prof)/100)), 2))+' %'
    table = Pr(["Name", "Total estimate"])
    table.add_row(['total gamer positive estimate: ', total_pos_gam])
    table.add_row(['total gamer negative estimate: ', total_neg_gam])
    table.add_row(['total reviewer positive estimate: ', total_pos_pro])
    table.add_row(['total reviewer negative estimate: ', total_neg_pro])
    table.add_row(["relation of negative/positive gamer's estimate: ", st_rel_gam])
    table.add_row(["relation of negative/positive reviewer's estimate: ", perc_rev_rel_est])
    table.add_row(["total relation percent: ", total_relation_percent])
    return print(table)


get_wb(table_src)
