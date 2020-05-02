# -*- coding: utf-8 -*-
"""
Created on Thu Apr 30 11:19:43 2020

@author: Vasil
"""

import numpy as np
import pandas as pd
import datetime as dt  # , pytzq , grid,

from dateutil.relativedelta import relativedelta
CB = str(dt.date.today()-dt.timedelta(30))
OB = str(dt.date.today()+relativedelta(months=-18))


# In[ ]:


def generate_case(g):
    """Функция для заполнения поля case
    g - группа записей в виде датафрейма
    """

    rules = {
        "Q-UNSELLABLE": [
            "P-SELLABLE",
            "P-DEFECTIVE",
            "P-DISTRIBUTOR_DAMAGED",
            "P-CUSTOMER_DAMAGED",
            "P-EXPIRED",
            # "P-EXPIRED",
        ],
        "Q-WAREHOUSE_DAMAGED": [
            "P-SELLABLE",
            "P-DEFECTIVE",
            "P-DISTRIBUTOR_DAMAGED",
            "P-CUSTOMER_DAMAGED",
            "P-EXPIRED",
            # "P-EXPIRED",
        ],
        "Q-CARRIER_DAMAGED": [
            "P-SELLABLE",
            "P-DEFECTIVE",
            "P-DISTRIBUTOR_DAMAGED",
            "P-CUSTOMER_DAMAGED",
            "P-EXPIRED",
            # "P-EXPIRED",
        ],
        "P-UNSELLABLE": [
            "Q-SELLABLE",
            # "Q-WAREHOUSE_DAMAGED",
            # "Q-CARRIER_DAMAGED",
        ],
        "P-WAREHOUSE_DAMAGED": [
            "Q-SELLABLE",
            # "P-WAREHOUSE_DAMAGED",
            # "P-CARRIER_DAMAGED",
        ],
        "P-CARRIER_DAMAGED": [
            "Q-SELLABLE",
            # "P-WAREHOUSE_DAMAGED",
            # "P-CARRIER_DAMAGED",
        ],
    }

    # Предварительно заполняем столбец case
    # значениями "<reason>-<status>".
    g.case = g.reason.str.cat(g.disposition, sep='-')

    # Проверяем есть ли значения из case
    # среди ключей в словаре правил.
    # Если найдено хотя бы одно совпадение (.any()),
    # то продолжаем работать с этой группой.
    if g.case.isin(rules).any():

        # Мы знаем, что в группе записей для одного товара
        # в пределах одного документа может быть только
        # одна запись с хорошим состоянием.
        # Получим значение.
        good_condition = g[g.case.isin(rules)].case.values[0]
        # good_condition = g.query('case in @rules').case.values[0]

        # Достаем из словаря правил список подозрительных
        # значений для данного хорошего состояния.
        suspicious_conditions = rules[good_condition]  # <<<<<<<<<<<<

        # Перезаполняем case.
        # Если значение ячейки case есть в списке
        # подозрительных проводок для данного ключа, то
        # заполняем case - "<good_condition>:<suspicious_condition>",
        # иначе записываем nan.
        g.case = np.where(g.case.isin(suspicious_conditions),  # <<<<<<<<<<<
                          good_condition + ':' + g.case,
                          np.nan)

    else:
        g.case = np.nan

    return g


# In[ ]:
def data_processing(df_rec, df_adj, df_rei):
    df_adj = df_adj.sort_values(by='transaction_item_id')
    df_adj = df_adj.reset_index(drop=True)

    # In[ ]:

    Max_diff, document_number = 10000, 1
    df_adj['tr_id'] = 0
    df_adj.tr_id = np.where(df_adj.transaction_item_id.diff(periods=1) <= Max_diff,
                            None,
                            np.where(df_adj.transaction_item_id.diff(periods=-1)
                                     >= -Max_diff,
                                     df_adj.transaction_item_id,
                                     0
                                     ).astype('int64')
                            )

    df_adj.tr_id.ffill(inplace=True)
    
    # In[ ]:

    df_work = df_adj[df_adj.tr_id != 0]
    df_work['case'] = ""
    # df_work.loc[:, ('case',)] = ""

    # In[ ]:
    df_work = (df_work.groupby(['tr_id'])
               # Применяем функцию для заполнения case
               # Будут заполнены только ячейки, удовлетворяющие правилам.
                      .apply(generate_case)
                      .dropna(subset=['case'])
               )

    # In[ ]:

    # Берем из df_work столбцы 'id' и 'quantity'.
    # Группируем по id товара и суммируем quantity
    df_quantity_sum = (df_work.loc[:, ['fnsku', 'quantity']]
                              .groupby('fnsku')
                              .sum()
                              .reset_index()
                       )

    # In[ ]:

    df_quantity_sum.rename(columns={'quantity': '17_Regraded_to_sell'},
                           inplace=True)

    # In[ ]:

    df_rec = pd.merge(df_rec, df_quantity_sum, how='left', on='fnsku').fillna(0)

    # In[ ]:

    checking_cases = ['E', '5', 'Damaged_by_FC', 'Disposed_of', 'M', 'Misplaced', ]
    df_work = df_adj[df_adj.reason.isin(checking_cases) &
                     (df_adj['adjusted_date'] < CB) &
                     (df_adj['adjusted_date'] > OB)]

    df_work = pd.pivot_table(df_work,
                             columns='reason',
                             values='quantity',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    # In[ ]:

    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    # In[ ]:
    
    df_work = pd.pivot_table(
                        df_adj[
                                (df_adj.reason.isin(
                                        ['F',
                                         'N',
                                         'O',
                                         '1',
                                         '2',
                                         # '3',
                                         # '4',
                                         'D',
                                         'G',
                                         'I',
                                         'K',
                                         ]) &
                                    ((df_adj['adjusted_date'] > OB))
                                 )
                        ],
                        columns='reason',
                        values='quantity',
                        index='fnsku',
                        aggfunc=np.sum,
                        fill_value=0
            )

    # In[ ]:

    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    # In[ ]:
    
    df_work = pd.pivot_table(
                        df_rei[
                               df_rei.reason.isin(
                                       ['Lost_Warehouse',
                                        'Damaged_Warehouse',
                                        # 'Reimbursement_Reversal',
                                        ]
                                ) & (df_rei.approval_date > OB)
                        ],
                        values=
                        # [
                        'quantity_reimbursed_total',
                        # 'quantity_reimbursed_cash',
                        # 'quantity_reimbursed_inventory',
                        # ],
                        columns='reason',
                        index='fnsku',
                        aggfunc=np.sum,
                        fill_value=0
            )

    # In[ ]:

    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    # In[ ]:

    df_work = pd.pivot_table(
        df_rei[df_rei.approval_date > OB],  # [
        # df_rei.reason.isin(['Reimbursement_Reversal'])
        # ],
        values=
        # [
        # 'quantity_reimbursed_total',
        # 'quantity_reimbursed_cash',
        'quantity_reimbursed_inventory',
        # ],
        # columns='reason',
        index='fnsku',
        aggfunc=np.sum,
        fill_value=0
            )

    # In[ ]:

    df_rec = pd.merge(df_rec, -df_work, how='left', on='fnsku').fillna(0)

    # In[ ]:

    df_work = df_adj.loc[(df_adj.disposition == 'WAREHOUSE_DAMAGED')
                         & df_adj.reason.isin(['M', 'F', 'N', 'O'])
                         & (df_adj.adjusted_date > OB),
                         ['fnsku', 'quantity']
                         ]
    df_work.rename(columns={'quantity': '15_DMG_lost'}, inplace=True)

    # In[ ]:

    df_work = pd.pivot_table(df_work,
                             # columns='reason',
                             values='15_DMG_lost',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    # In[ ]:

    df_rec = pd.merge(df_rec, -df_work, how='left', on='fnsku').fillna(0)


    # In[ ]:

    # берем только те диспозалы, что не могли реимбурсить
    df_work = df_adj[
        (df_adj.reason == 'D')
        & df_adj.disposition.isin(
            [
                'SELLABLE',
                'DEFECTIVE',
                'DISTRIBUTOR_DAMAGED',
                'UNSELLABLE',
                'CUSTOMER_DAMAGED'
            ]
                                 )
        & (df_adj.adjusted_date > OB)
                  ]

    df_work = pd.pivot_table(
        df_work,
        columns='reason',
        values='quantity',
        index='fnsku',
        aggfunc=np.sum,
        fill_value=0
    )

    # In[ ]:
    df_work = df_work.rename(columns={'D': '12_Dispsd_sellbl'})

    # In[ ]:
    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    # In[ ]:
    # нужно сверить tr.out + disposal of damaged
    # с реимбурсами за порчу.

    df_work = df_adj[
        (df_adj.reason == 'D')
        & df_adj.disposition.isin(
            [
                'WAREHOUSE_DAMAGED',
                'CARRIER_DAMAGED',
                'UNSELLABLE',
            ]
                                 )
        & (df_adj.adjusted_date > OB)
                  ]

    df_work = pd.pivot_table(
        df_work,
        columns='reason',
        values='quantity',
        index=['fnsku', 'disposition'],
        aggfunc=np.sum,
        fill_value=0
            )
    df_work = df_work.rename(columns={'D': 'Disp_reimbrsble'})
    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    # In[ ]:
    df_rec['05_Tr.OUT_reasn'] = pd.DataFrame(
        df_rec.get("Damaged_Warehouse", default=0)
                .add(df_rec.get("Disp_reimbrsble", default=0))
            ).join(-df_rec.O).min(axis=1)

    # In[ ]:
    df_rec['10_LOST'] = (
        df_rec.get('M', default=0)
        + df_rec.get('F', default=0)
        + df_rec.get('Lost_Warehouse', default=0)
        + df_rec.get('discrepant_quantity', default=0)
        + df_rec.get('N', default=0)
        + df_rec.get('quantity_reimbursed_inventory', default=0)
        + df_rec.get('O', default=0)
        + df_rec.get('05_Tr.OUT_reasn', default=0)
        )

    df_rec['20_DMG'] = (
        df_rec.get('E', default=0)
        + df_rec.get('12_Dispsd_sellbl', default=0)
        + df_rec.get('17_Regraded_to_sell', default=0)  # Q
        + df_rec.get('K', default=0)
        + df_rec.get('G', default=0)
        + df_rec.get('I', default=0)
        + df_rec.get('Damaged_Warehouse', default=0)
        + df_rec.get('15_DMG_lost', default=0)
        + df_rec.get('quantity', default=0)
        )

    df_rec['01_TOTL'] = (df_rec['10_LOST'].where(df_rec['10_LOST'] <= 0, 0)
                           + df_rec['20_DMG'].where(df_rec['20_DMG'] <= 0, 0)
                           )

    # In[ ]:
    df_rec.rename(
        columns={'quantity': 'D_Regraded',
                 'E': '11_Damaged',
                 'D': 'Disposed',
                 'M': '02_Misplaced',
                 '5': '13_Unrecoverable',
                 'F': '02_Found',
                 'quantity_reimbursed_inventory': '04_Tr.IN_reasn',
                 'G': '14_Unexplained',
                 'K': 'Defect_dmg',
                 'N': '03_Tr.IN',
                 'O': '05_Tr.OUT',
                 'Lost_Warehouse': '07_Reimb_Loss',
                 'Damaged_Warehouse': '16_Reimb_Dmg',
                 'discrepant_quantity': '02_Discrepancy',
                 'fnsku': '00_fnsku'
                 }, inplace=True
        )

    # In[ ]:
    df_rec.sort_values(by='01_TOTL', inplace=True)
    df_rec.sort_index(axis=1, inplace=True)

    # In[]:
    print("===df_rec_pivoting===")
    
    df_rec = df_rec[df_rec["01_TOTL"] < 0]
    print("===pivoting01==")
    df_adj_filtered = df_adj[df_adj.fnsku.isin(df_rec["00_fnsku"])]
    df_adj_filtered['disposition'].replace({
        'SELLABLE': 'Z__SELLABLE',
        'CARRIER_DAMAGED': 'Zz_CARRIER_DAMAGED',
        'UNSELLABLE': 'Zz_UNSELLABLE',
        'WAREHOUSE_DAMAGED': 'Zz_WAREHOUSE_DAMAGED', 
        }, inplace=True)
    table = pd.pivot_table(df_adj_filtered,
                           values='quantity', columns=['disposition'],
                           index = ["fnsku", 'adjusted_date', "fulfillment_center_id", 
                                    "transaction_item_id", "tr_id", "reason"],
                           aggfunc=np.sum)

    return (df_rec, table)


def excel_writer(folders_path, client_folder, df_rec, table):
    new_file_path = folders_path / client_folder / f"{client_folder}_snapshot_.xlsx"
    with pd.ExcelWriter(str(new_file_path), engine='xlsxwriter') as writer:
        df_rec.to_excel(writer, sheet_name='Sheet01', index=False)
        table.to_excel(writer, sheet_name='Pivot_ADJ')
    return new_file_path
