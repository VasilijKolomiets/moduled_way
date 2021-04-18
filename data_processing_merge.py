"""
Created on Thu Apr 30 11:19:43 2020.

@author: Vasil
"""

import numpy as np
import pandas as pd
# pd.options.mode.chained_assignment = None  # default='warn'

import datetime as dt
from dateutil.relativedelta import relativedelta
import pivottablejs


# In[ ]:


def generate_case(g):
    """Функция для заполнения поля case.

    g - группа записей в виде датафрейма
    """
    pos_reasons = [
        "P-SELLABLE",
        "P-DEFECTIVE",
        "P-DISTRIBUTOR_DAMAGED",
        "P-CUSTOMER_DAMAGED",
        "P-EXPIRED",
    ]
    neg_reasons = [
        "Q-SELLABLE",
        "Q-DEFECTIVE",
        "Q-DISTRIBUTOR_DAMAGED",
        "Q-CUSTOMER_DAMAGED",
        "Q-EXPIRED",
    ]

    rules = {
        "Q-UNSELLABLE": pos_reasons,
        "Q-WAREHOUSE_DAMAGED": pos_reasons,
        "Q-CARRIER_DAMAGED": pos_reasons,
        "P-UNSELLABLE": neg_reasons,
        "P-WAREHOUSE_DAMAGED": neg_reasons,
        "P-CARRIER_DAMAGED": neg_reasons,
    }
    # Предварительно заполняем столбец case
    # значениями "<reason>-<status>".
    g.case = g.reason.str.cat(g.disposition, sep='-')

    # Проверяем есть ли значения из case
    # среди ключей в словаре правил.
    # Если найдено хотя бы одно совпадение (.any()),
    # то продолжаем работать с этой группой.
    if g.case.isin(rules).any():
        keys_matching = g[g.case.isin(rules)].case.values

        # Достаем из словаря правил список подозрительных
        # значений для каждого ключа из набора.
        for key in keys_matching:
            suspicious_conditions = rules[key]
            # Перезаполняем case.
            # Если значение ячейки case есть в списке
            # подозрительных проводок для данного ключа, то
            # заполняем case - "<good_condition>:<suspicious_condition>",
            # иначе записываем nan.
            g.case1 = np.where(g.case.isin(suspicious_conditions),  # <<<<<<<<<<<
                               key + ':' + g.case,
                               np.nan)

    return g


# In[ ]:
def data_processing(df_rec, df_adj, df_rei, OB=None):

    #  чтобы срезать бызы по датам от до OB до CB
    # CB = str(dt.date.today())
    today_minus_18_monthes = str(dt.date.today() + relativedelta(months=-18))
    OB = max(OB, today_minus_18_monthes) if OB else today_minus_18_monthes

    df_adj = df_adj.sort_values(by='transaction_item_id')
    df_adj = df_adj.reset_index(drop=True)

    df_adj = df_adj[  # (df_adj['adjusted_date'] <= CB) &
        (df_adj['adjusted_date'] >= OB)]

    df_rei = df_rei[  # (df_rei['approval_date'] <= CB) &
        (df_rei['approval_date'] >= OB)]

    rei_fnsku = set(df_rei["fnsku"])
    rec_fnsku = set(df_rec["fnsku"])

    df_rec = pd.concat([df_rec, *[pd.DataFrame([
        ['sku', fnsku, 'asin', 'product_name', 'condition', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], ],
        columns=df_rec.columns.values) for fnsku in (rei_fnsku - rec_fnsku)]
    ],
        ignore_index=True)

# In[ ]:

    Max_diff = 10000
    df_adj['tr_id'] = 0
    df_adj.tr_id = np.where(
        df_adj.transaction_item_id.diff(periods=1) <= Max_diff,
        None,
        np.where(df_adj.transaction_item_id.diff(periods=-1) >= -Max_diff,
                 df_adj.transaction_item_id,
                 0,
                 ).astype('int64')
    )
    df_adj.tr_id.ffill(inplace=True)

    df_work = df_adj[df_adj.tr_id != 0].copy()
    df_work['case'], df_work['case1'] = np.nan, np.nan
    if df_work.empty:
        return {'df_rec': None, 'table': None, "df_adj_filtered": None}

    df_work = (df_work.groupby(['tr_id'])
               # Применяем функцию для заполнения case
               # Будут заполнены только ячейки, удовлетворяющие правилам.
                      .apply(generate_case)
                      .dropna(subset=['case1'])
               )

    # Берем из df_work столбцы 'id' и 'quantity'.
    # Группируем по id товара и суммируем quantity
    df_quantity_sum = (df_work.loc[:, ['fnsku', 'quantity']]
                              .groupby('fnsku')
                              .sum()
                              .reset_index()
                       )

    df_quantity_sum.rename(columns={'quantity': '17_Regraded_to_sell'},
                           inplace=True)

    df_rec = pd.merge(df_rec, df_quantity_sum, how='left', on='fnsku').fillna(0)

    checking_cases1 = ['E', '5', 'Damaged_by_FC', 'Disposed_of', 'M', 'Misplaced', ]
    df_work = df_adj[df_adj.reason.isin(checking_cases1)].copy()

    df_work = pd.pivot_table(df_work,
                             columns='reason',
                             values='quantity',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    #  df_rec.join(df_work, on='fnsku').fillna(0)
    #  df_rec = pd.merge(df_rec, df_work, how='left', left_on='fnsku', right_index=True).fillna(0)
    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    checking_cases2 = ['F', 'N', 'O', '1', '2', 'D', 'G', 'I', 'K', ]    # '3', # '4',
    df_work = pd.pivot_table(df_adj[df_adj.reason.isin(checking_cases2)],
                             columns='reason',
                             values='quantity',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    df_rei = df_rei[["reason", "fnsku", 'quantity_reimbursed_total',
                     'quantity_reimbursed_inventory']]
    df_work = pd.pivot_table(
        df_rei[
            df_rei.reason.isin(
                [
                    'Lost_Warehouse',
                    'Damaged_Warehouse',
                ]
            )
        ],
        values='quantity_reimbursed_total',
        columns='reason',
        index='fnsku',
        aggfunc=np.sum,
        fill_value=0,
    )

    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    df_work = pd.pivot_table(df_rei,
                             values='quantity_reimbursed_inventory',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    df_rec = pd.merge(df_rec, -df_work, how='left', on='fnsku').fillna(0)

    df_work = df_adj.loc[(df_adj.disposition == 'WAREHOUSE_DAMAGED')
                         & df_adj.reason.isin(['M', 'F']),
                         ['fnsku', 'quantity']
                         ].copy()

    df_work.rename(columns={'quantity': '15_DMG_lost'}, inplace=True)

    df_work = pd.pivot_table(df_work,
                             # columns='reason',
                             values='15_DMG_lost',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    df_rec = pd.merge(df_rec, -df_work, how='left', on='fnsku').fillna(0)

    # берем только те диспозалы, что не могли реимбурсить
    df_work = df_adj[
        (df_adj.reason == 'D')
        & df_adj.disposition.isin([
            'SELLABLE',
            'DEFECTIVE',
            'DISTRIBUTOR_DAMAGED',
            'UNSELLABLE',
            'CUSTOMER_DAMAGED',
        ])
    ].copy()

    df_work = pd.pivot_table(
        df_work,
        columns='reason',
        values='quantity',
        index='fnsku',
        aggfunc=np.sum,
        fill_value=0
    )

    df_work = df_work.rename(columns={'D': '12_Dispsd_sellbl'})
    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    # нужно сверить tr.out + disposal of damaged
    # с реимбурсами за порчу.
    df_work = df_adj[(df_adj.reason == 'D')
                     & df_adj.disposition.isin([
                         'WAREHOUSE_DAMAGED',
                         'CARRIER_DAMAGED',
                         'UNSELLABLE',
                     ])
                     ].copy()

    df_work = pd.pivot_table(
        df_work,
        columns='reason',
        values='quantity',
        index=['fnsku', 'disposition'],
        aggfunc=np.sum,
        fill_value=0
    )         # .sort_index().reset_index().set_index('fnsku')
    df_work = df_work.rename(columns={'D': 'Disp_reimbrsble'})
    df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    df_rec['O'] = df_rec.get('O', default=0)  # for if "O" not exists

    df_rec['05_Tr.OUT_reasn'] = pd.DataFrame(
        df_rec.get("Damaged_Warehouse", default=0)
        .add(df_rec.get("Disp_reimbrsble", default=0))
    ).join(-df_rec.O).min(axis=1)

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

    df_rec.sort_values(by='01_TOTL', inplace=True)
    df_rec.sort_index(axis=1, inplace=True)

    print("===df_rec_pivoting===")
    # ----------  df_rec = df_rec[df_rec["01_TOTL"] < 0  ]
    df_rec = df_rec[df_rec.iloc[:, 1:df_rec.columns.get_loc("asin")].copy().apply(abs).sum(axis=1)
                    > 0]

    print("===pivoting01==")
    df_adj_filtered = df_adj[df_adj.fnsku.isin(df_rec["00_fnsku"])].copy()
    df_adj_filtered['disposition'].replace({
        'SELLABLE': 'Z__SELLABLE',
        'CARRIER_DAMAGED': 'Zz_CARRIER_DAMAGED',
        'UNSELLABLE': 'Zz_UNSELLABLE',
        'WAREHOUSE_DAMAGED': 'Zz_WAREHOUSE_DAMAGED',
    }, inplace=True)
    try:
        table = pd.pivot_table(df_adj_filtered,
                               values='quantity', columns=['disposition'],
                               index=["fnsku", 'adjusted_date',
                                      "fulfillment_center_id",
                                      "transaction_item_id", "tr_id", "reason"],
                               aggfunc=np.sum,
                               margins=True,
                               )
    except ValueError:  # ValueError: No objects to concatenate   with #margins=True,
        table = pd.DataFrame({'Z__SELLABLE': {0: 'No matching data'}})

    return {'df_rec': df_rec, 'table': table, "df_adj_filtered": df_adj_filtered}


def excel_writer(folders_path, client_folder, client_name, files):

    new_xlsx_path = folders_path / client_folder / f"{client_name}_snapshot_.xlsx"
    with pd.ExcelWriter(str(new_xlsx_path), engine='xlsxwriter') as writer:
        files['df_rec'].to_excel(writer, sheet_name='Sheet01', index=False)
        files['table'].to_excel(writer, sheet_name='Pivot_ADJ',  merge_cells=False)

    new_html_path = folders_path / client_folder / f'{client_name}_pivottablejs.html'
    pivottablejs.pivot_ui(files["df_adj_filtered"],
                          rows=["fnsku", 'adjusted_date', "fulfillment_center_id",
                                "transaction_item_id", "tr_id", "reason"],
                          cols=['disposition'],
                          aggregatorName="Integer Sum",
                          vals=['quantity'],
                          rendererName="Col Heatmap",
                          outfile_path=str(new_html_path)
                          )

    return {"xlsx": new_xlsx_path, "html": new_html_path}
