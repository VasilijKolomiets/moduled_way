"""
Created on Thu Apr 30 11:19:43 2020.

@author: Vasil
"""

import numpy as np
import pandas as pd
# pd.options.mode.chained_assignment = None  # default='warn'

import csv

import datetime as dt
from dateutil.relativedelta import relativedelta
import pivottablejs

from pathlib import Path


from copy import copy, deepcopy

SNAPSHOTS_PATTERNS = ["rec", "adj", "rei"]
RETURNED_NOT_ADDED = ["eve", "ret", "rei"]   # 2022
# In[ ]:


def rename_df_columns(df_):
    """Rename the df columns  by replacing character "-" with "_" for avoid syntax problems."""
    dict_ = {x: x.replace("-", "_") for x in df_.columns}
    df_.rename(columns=dict_, inplace=True)


def file_names_reader(path_to_files_folder: Path,
                      separators: str = " _-",
                      patterns: list = SNAPSHOTS_PATTERNS,
                      ) -> dict:
    """
    Search files in specified directory
    - path_to_files_folder: Path -
    according patterns in parameter
    - patterns -
    """
    import re

    wrk_dir = path_to_files_folder
    p_compiled = {p: re.compile(f'.*[{separators}]{p}.*\.csv') for p in patterns}
    files = dict()
    for kp, pc in p_compiled.items():
        files[kp] = [f for f in wrk_dir.iterdir() if f.is_file() and pc.search(f.name.lower())]
        if len(files[kp]) > 1:
            raise Exception(f'несколько файлов с подстрокой {kp}')
        if len(files[kp]) == 0:
            raise Exception(f'неn файлов с подстрокой {kp}')

    return {key: file_names[0] for key, file_names in files.items()}


#   ===== 2022 starts(2) =========

def reshape(df: pd.DataFrame, shape: dict) -> pd.DataFrame:
    """
    Reshape input pd.DataFrame according to given 'shape' dictionary.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame for reshapimg.
    shape : dict
        Dictionary with structure:
            {
               recolumn: dict(
                        'col_name_to rename_01': 'new_column_name_01',
                        'col_name_to rename_02': 'new_column_name_02',
                        ...
                    ),
               add_columns=('adding_column_name_01', 'adding_column_name_02', ...),
               sort_columns=['column_to_sort_on_1', 'column_to_sort_on_2', ...]
            }


    Returns
    -------
    df : pd.DataFrame
        Reshaped DataFrame.

    Example:

    >>> reshape(
        pd.DataFrame({'c':[1,2,3], 'b':[3,4,5], 'a':[8,7,6]}),
        {'recolumn':{'b': '01_b'}, 'add_columns': ('f','d'), sort_columns=('a')}
        )
    >>>
       01_b  a  c     d     f
    0     3  8  1  None  None
    1     4  7  2  None  None
    2     5  6  3  None  None
    """
    if shape['add_columns']:
        df.loc[:, shape['add_columns']] = None

    shape['recolumn'] and df.rename(columns=shape['recolumn'], inplace=True)
    df.sort_index(axis=1, inplace=True)    # columns by names

    # , ascending=[True, False], na_position='first'
    shape['sort_columns'] and df.sort_values(by=shape['sort_columns'], inplace=True)

    return df

#   ===== 2022 ends(2) =========

# In[1]:


def check_repare_csv_header(file_name, en_coding='latin1'):
    HAVE_TO_REWRITE = False
    with open(file_name, mode='r', encoding=en_coding) as file:
        first_line = file.readline()
        if first_line[0] not in {'"', "'"}:   # starts not like literal !
            ind = first_line.find('"')
            if ind != -1:
                first_line = first_line[ind:]
            elif first_line.find("'") != -1:
                first_line = first_line[first_line.find("'"):]
            else:
                return None   # TODO: change dict_response!!
            #  HERE WE ARE after deleting first extrasimbols (like ?)
            lines = [line for line in file]
            lines.insert(0, first_line)
            HAVE_TO_REWRITE = True

    if HAVE_TO_REWRITE:
        with open(file_name, mode='w', encoding=en_coding) as f:
            f.writelines(lines)


def find_delimiter(filename, delimiters=(',', '\t')):
    sniffer = csv.Sniffer()
    with open(filename) as fp:
        delimiter = sniffer.sniff(fp.read(1024*64), delimiters=delimiters).delimiter
    return delimiter


def files_reading(path_to_files_folder, patterns: list = SNAPSHOTS_PATTERNS):   # 2022

    my_response = {             # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "files": {},
    }

    #  !chardetect direct_adj_20285364484018375.csv
    en_codings = ['utf-8', "latin1", "cp1252", "cp1251", "cp1250"]

    files = file_names_reader(path_to_files_folder, patterns=patterns)
    sep_files = dict.fromkeys(files.keys())
    readed = False
    for en_coding in en_codings:
        if readed:
            break
        try:
            for key, file_name in files.items():
                check_repare_csv_header(file_name)
                # separator:  '\t' = tab or "," = comma
                separator = find_delimiter(file_name, delimiters=(',', '\t'))
                sep_files[key] = pd.read_csv(str(file_name),
                                             # 2022 error_bad_lines=False,
                                             sep=separator,
                                             encoding=en_coding,
                                             )
                rename_df_columns(sep_files[key])
            readed = True
        except ValueError as error:   # TODO: break ???
            my_response["exit_message"] = f'{error} while reading {file_name}'
    else:
        my_response["exit_message"] = f'wrong encoding. Not from {en_codings}'

    if readed:
        my_response["exit_is_ok"] = True
        my_response["files"] = sep_files

    return my_response


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

def data_processing(rec=None, adj=None, rei=None, OB=None):

    my_response = {           # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "files": {'df_rec': None,
                  'table': None,
                  "df_adj_filtered": None,
                  },
    }

    df_rec, df_adj, df_rei = rec, adj, rei
    #  чтобы срезать бызы по датам от до OB до CB
    CB = str(dt.date.today())   # + relativedelta(months=-1)   #TODO: 2022
    today_minus_18_monthes = str(dt.date.today() + relativedelta(months=-18))
    OB = max(OB, today_minus_18_monthes) if OB else today_minus_18_monthes

    df_adj = df_adj.sort_values(by='transaction_item_id')
    df_adj = df_adj.reset_index(drop=True)

    df_adj = df_adj[df_adj['adjusted_date'] >= OB]
    df_rei = df_rei[df_rei['approval_date'] >= OB]

    rei_fnsku = set(df_rei["fnsku"].unique())
    rec_fnsku = set(df_rec["fnsku"].unique())

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
        my_response["files"] = {'df_rec': df_rec, 'table': None,
                                "df_adj_filtered": df_adj,
                                }  # TODO: 21/11/30
        my_response['exit_message'] = "nothing_to_do"  # TODO: 21/11/30
        return my_response

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

    df_quantity_sum.rename(columns={'quantity': '17_Regraded_to_sell'}, inplace=True)
    df_rec = pd.merge(df_rec, df_quantity_sum, how='left', on='fnsku').fillna(0)

    # TODO:     -     M E - 30 days+
    df_adj = df_adj.query(
        "((adjusted_date < @CB) and reason in ('M', 'E', '5')) or (reason not in ('M', 'E', '5'))"
    )
    # adding new work columns  # done 21.11.11
    new_adj_columns = ['unrec_M', 'unrec_F', 'unrec_E', 'M_damaged',
                       '10__lost_to_compare', '10__lost_check',
                       '20__damaged_to_compare', '20__damaged_check',
                       ]
    df_adj.loc[:, new_adj_columns] = 0
    df_adj.loc[df_adj.reason == 'M', 'unrec_M'] = df_adj.unreconciled
    df_adj.loc[df_adj.reason == 'F', 'unrec_F'] = df_adj.unreconciled
    df_adj.loc[df_adj.reason == 'E', 'unrec_E'] = df_adj.unreconciled
    df_adj.loc[(df_adj.reason == 'M')
               & df_adj.disposition.isin(["WAREHOUSE_DAMAGED", "CARRIER_DAMAGED"]),
               'M_damaged'] = df_adj.unreconciled
    df_adj.loc[:, '10__lost_to_compare'] = df_adj.unrec_M - df_adj.unrec_F
    df_adj.loc[:, '20__damaged_to_compare'] = df_adj.unrec_E + df_adj.M_damaged
    # TODO:     -     M E - 30 days+

    #  на это действует 30 дней
    checking_cases1 = ['E', '5', 'Damaged_by_FC', 'Disposed_of', 'M', 'Misplaced', ]
    df_work = df_adj[df_adj.reason.isin(checking_cases1)].copy()

    df_work = pd.pivot_table(df_work,
                             columns='reason',
                             values='quantity',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    df_rec = df_rec.join(df_work, on='fnsku').fillna(0)
    #  df_rec = pd.merge(df_rec, df_work, how='left', left_on='fnsku', right_index=True).fillna(0)
    #  df_rec = pd.merge(df_rec, df_work, how='left', on='fnsku').fillna(0)

    checking_cases2 = ['F', 'N', 'O', '1', '2', 'D', 'G', 'I', 'K', ]    # '3', # '4',
    df_work = pd.pivot_table(df_adj[df_adj.reason.isin(checking_cases2)],
                             columns='reason',
                             values='quantity',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )
    df_rec = df_rec.join(df_work, on='fnsku').fillna(0)

    # TODO: 21/11/11
    df_adj_agg = df_adj[['fnsku', '10__lost_to_compare',
                         '20__damaged_to_compare']].groupby('fnsku').sum()
    df_rec = df_rec.join(df_adj_agg, on='fnsku').fillna(0)

    # TODO: 21/11/11

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

    df_rec = df_rec.join(df_work,  on='fnsku').fillna(0)

    df_work = pd.pivot_table(df_rei,
                             values='quantity_reimbursed_inventory',
                             index='fnsku',
                             aggfunc=np.sum,
                             fill_value=0
                             )

    df_rec = df_rec.join(-df_work, on='fnsku').fillna(0)

    df_work = df_adj.loc[(df_adj.disposition == 'WAREHOUSE_DAMAGED')
                         & df_adj.reason.isin(['M', 'F']),
                         ['fnsku', 'quantity']
                         ].copy()

    df_work.rename(columns={'quantity': '15_DMG_lost'}, inplace=True)

    if df_work.empty:                           # TODO: 21/11/30
        df_rec.loc[:, '15_DMG_lost'] = 0  # adding zero column    # TODO: 21/11/30

    else:
        df_work = pd.pivot_table(df_work,
                                 # columns='reason',
                                 values='15_DMG_lost',
                                 index='fnsku',
                                 aggfunc=np.sum,
                                 fill_value=0
                                 )
        # indexed for non overlaping fnsku    # TODO: 21/11/30
        df_rec = df_rec.join(-df_work, on='fnsku').fillna(0)

    # берем только те дисапозоны, что не могли реимбурсить
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
    df_rec = df_rec.join(df_work, on='fnsku').fillna(0)

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
    ).sort_index().reset_index().set_index('fnsku')

    if not df_work.empty:
        df_work = df_work.rename(columns={'D': 'Disp_reimbrsble'})
        df_rec = df_rec.join(df_work, on='fnsku').fillna(0)

        df_rec.reset_index(inplace=True)
        df_rec.set_index(['fnsku', 'disposition'], inplace=True)
        df_rec.reset_index(inplace=True)

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

    # TODO: 21/11/30
    columns_to_rename = {'quantity': 'D_Regraded',
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
                         }
    missing_columns = [name for name in columns_to_rename.keys() if name not in df_rec.columns]
    df_rec.loc[:, missing_columns] = 0  # adding missing reasons columns
    # TODO: 21/11/30
    df_rec.rename(columns=columns_to_rename, inplace=True)

    # TODO: 21/11/11
    df_rec.loc[:, '10__lost_check'] = (df_rec['02_Found'] + df_rec['02_Misplaced']
                                       + df_rec['07_Reimb_Loss'])
    df_rec.loc[:, '20__damaged_check'] = (df_rec['11_Damaged'] + df_rec['15_DMG_lost']
                                          + df_rec['16_Reimb_Dmg']
                                          # + df_rec['17_Regraded_to_sell'] - ОДНООБРАЗИЕ СРАВНЕНИЯ
                                          )

    # TODO: 21/11/11

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

    # TODO: 21/11/11
    # 3-d sheet named 'adj' is df_adj with М+Е anrecociled.  Sort on fnsku field.
    df_adj.drop(columns=new_adj_columns, axis=1, inplace=True)
    df_adj_filtered_ME = df_adj.query("(reason in ('M', 'E')) and (unreconciled > 0)")
    df_adj_filtered_ME = df_adj_filtered_ME.sort_values(by=['fnsku', ])   # , inplace=True)

    # TODO: 21/11/11
    my_response["exit_is_ok"] = True
    my_response["files"] = {'df_rec': df_rec, 'table': table,
                            "df_adj_filtered": df_adj_filtered,
                            "df_adj_filtered_ME": df_adj_filtered_ME}  # TODO: 21/11/11

    return my_response


def excel_writer(path_to_files_folder, client_name: str, files):

    new_xlsx_path = path_to_files_folder / f"{client_name}_snapshot_.xlsx"
    with pd.ExcelWriter(str(new_xlsx_path), engine='xlsxwriter') as writer:
        files['df_rec'].to_excel(writer, sheet_name='Sheet01', index=False)
        files['table'].to_excel(writer, sheet_name='Pivot_ADJ',  merge_cells=False)
        # TODO: 21/11/11
        files['df_adj_filtered_ME'].to_excel(writer, sheet_name='adj_filtered_ME', index=False)

    new_html_path = path_to_files_folder / f'{client_name}_pivottablejs.html'
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


# In[ 2]:


def fee_file_names_reader(path_to_files_folder: Path) -> dict:
    """
    Search files in specified directory 'path_to_files_folder'.

    Parameters:
    - path_to_files_folder: Path -

    Retuns
    - dict with readed elements:
        fee_prev: df,               (from .csv)
        settlements: [df, df, df]    (from many  .txt)
    """
    wrk_dir = path_to_files_folder

    files = dict(fee_prev=None, settlements=[])

    files = [f for f in wrk_dir.iterdir() if f.is_file() and f.suffix in ('.txt', '.csv')]
    csv_f = [f for f in files if f.suffix == '.csv']
    txt_f = [f for f in files if f.suffix == '.txt']

    return {'fee_prev': csv_f, 'settlements': txt_f}


def orders_fees_input(folder=r".\Brada_Fee", subfolder="txt"):     # TODO: to module
    """Read and join all files with orders details (Settlements)."""
    my_response = {             # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "df_file": {},
    }

    folder = Path(folder) / subfolder

    dtypes = {
        'sku': str,
        'settlement-id': str,
        'currency': str,
        'transaction-type': str,
        'order-id': str,
        'merchant-order-id': str,
        'adjustment-id': str,
        'marketplace-name': str,
        'shipment-id': str,
        'fulfillment-id': str,
        # 'promotion-id': str,
        'order-item-code': str,
        'merchant-order-item-id': str,
        'merchant-adjustment-item-id': str,

        'amount-type': str,
        'amount-description': str,
        'amount': str,  # np.float32,
        'quantity-purchased': str,  # np.float32,
    }
    '''
                    'quantity_purchased': np.int32,
                    'amount': np.int32,
    '''

    columns_to_read = list(dtypes.keys())
    columns_to_read[-5:-5] = ['posted-date']

    en_codings = ["latin1", "cp1252", "cp1251", "cp1250"]
    readed = False
    for en_coding in en_codings:
        if readed:
            break
        try:
            txt_files = (pd.read_csv(f, sep="\t", skiprows=[1, ], decimal=",",
                                     encoding=en_coding,
                                     dtype=dtypes,
                                     parse_dates=['posted-date'],
                                     usecols=columns_to_read,
                                     )
                         for f in folder.glob("*.txt")
                         )

            readed = True
        except ValueError as error:
            print(error)
    else:
        my_response["exit_message"] = f'wrong encoding. Not from {en_codings}'

    if not readed:
        return my_response

    df = pd.concat(txt_files)
    rename_df_columns(df)

    columns_to_take = ['sku', 'settlement_id', 'marketplace_name', 'order_id', 'shipment_id',
                       'transaction_type', 'amount_type', 'amount_description',
                       'quantity_purchased', 'amount', 'posted_date', "currency",
                       ]
    df = df.query("amount_description == 'FBAPerUnitFulfillmentFee' ")[columns_to_take]

    # TODO: ask Yury for this block:

    """df_all = df_all.append(df_work[df_work['amount-description']
                                       .isin(['FBACustomerReturnPerOrderFee',
                                              'FBACustomerReturnPerUnitFee',
                                              'FBACustomerReturnWeightBasedFee',
                                              'FBAPerOrderFulfillmentFee',
                                              'FBAPerUnitFulfillmentFee',
                                              'FBAWeightBasedFee',
                                              ])
                                       ])

        df_all.sku = df_all.groupby('order_id').sku.ffill()
        df_all.sku = df_all.groupby('order_id').sku.bfill()

        # %% Dropping all SKUs left with NaN

        df_all.dropna(subset=['sku'], inplace=True)
    """
    # for future and EU compatibility df.marketplace_name.str.upper().str[-2:]
    df.loc[:, "currency"] = "USD"   # for future development
    df.loc[:, "country"] = "USA"

    my_response["exit_is_ok"] = True
    my_response["df_file"] = df
    return my_response


def fee_files_reading(path_to_files_folder):

    my_response = {             # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "files": {
            "FEE": None,
            "settlements": None,
        },
    }

    #  !chardetect direct_adj_20285364484018375.csv
    en_codings = ["latin1", "cp1252", "cp1251", "cp1250"]

    coma = ","  # source file field separators: '\t' = tab or "," = comma

    fee_files = fee_file_names_reader(path_to_files_folder)
    if fee_files['fee_prev'].__len__() > 1:
        my_response["exit_message"] = "More than 1 CSV file durimg FEE processing!!!"
        return my_response

    #  settlements readig & merging
    settlements = orders_fees_input(path_to_files_folder, subfolder='')
    if not settlements["exit_is_ok"]:
        my_response["exit_message"] = settlements["exit_message"] + '(from settlements reading)'
        return my_response

    my_response["files"]["settlements"] = settlements['df_file']

    readed = False
    for en_coding in en_codings:
        if readed:
            break
        try:
            check_repare_csv_header(fee_files['fee_prev'][0])
            df_fee_preview = pd.read_csv(str(fee_files['fee_prev'][0]),
                                         error_bad_lines=False,
                                         sep=coma,
                                         encoding=en_coding,
                                         )

            rename_df_columns(df_fee_preview)
            readed = True
        except ValueError as error:
            print(error)
    else:
        my_response["exit_message"] = f'wrong encoding. Not from {en_codings}'

    if readed:
        my_response["exit_is_ok"] = True
        my_response["files"]["FEE"] = df_fee_preview

    return my_response


def fee_data_processing(FEE=None, settlements=None, SPLIT_DATE=dt.datetime(2021, 6, 1)):
    """Process the data."""
    import fee_calc_2106 as fee_new
    import fee_calc as fee_old

    from copy import deepcopy

    files_dict = {"old_rooles": {'fee_preview': None, 'fee_diff': None},
                  "new_rooles": {'fee_preview': None, 'fee_diff': None},
                  }

    my_response = {             # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "files": {
            'USA': deepcopy(files_dict),
            # 'EU': deepcopy(files_dict),
        },
    }

    df_fee = FEE
    try:
        for country in ['USA', ]:       # TODO: countryes dict fee calc  # TODO: try ... exc
            print("2 new data_cleaning")
            df_fee_prev_new = fee_new.data_cleaning(df_fee)
            df_fee_prev_new = fee_new.fee_flag_calculation(df_fee_prev_new, country)
            print("2 old")
            df_fee_prev_old = fee_old.data_cleaning(df_fee)
            df_fee_prev_old = fee_old.fee_flag_calculation(df_fee_prev_old, country)

            print("4___orders_fees_input")
            df_orders_fee = settlements
            df_orders_fee_old = df_orders_fee.query("posted_date < @SPLIT_DATE")
            df_orders_fee_new = df_orders_fee.query("posted_date >= @SPLIT_DATE")

            # =============================================================================
            #      5. form the fee diference if it is
            # =============================================================================
            print("5 old  pick_difference")
            fee_diff_old = fee_old.pick_difference(df_orders_fee_old, df_fee_prev_old)
            print("5 new")
            fee_diff_new = fee_new.pick_difference(df_orders_fee_new, df_fee_prev_new)

            my_response["files"][country]["old_rooles"]['fee_preview'] = df_fee_prev_old
            my_response["files"][country]["old_rooles"]['fee_diff'] = fee_diff_old

            my_response["files"][country]["new_rooles"]['fee_preview'] = df_fee_prev_new
            my_response["files"][country]["new_rooles"]['fee_diff'] = fee_diff_new

            my_response["exit_is_ok"] = True

    except Exception as inst:
        my_response["exit_message"] = f'{__name__}: виникла помилка {inst}.'

    return my_response


def fee_excel_writer(folders_path, client_name, files):

    xl_files_path_dict = {"old_rooles":  None,   "new_rooles":  None}

    my_response = {             # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "xlsx_files": {
            'USA': copy(xl_files_path_dict),
            # 'EU': copy(xl_files_path_dict),
        },
    }

    try:
        for country, files_dict in files.items():
            for rooles in files_dict.keys():
                new_xlsx_path = folders_path / f"{client_name}_{country}_{rooles}_fees.xlsx"
                with pd.ExcelWriter(str(new_xlsx_path), engine='xlsxwriter') as writer:
                    files[country][rooles]['fee_preview'].to_excel(writer,
                                                                   sheet_name='fee_preview',
                                                                   index=False)
                    files[country][rooles]['fee_diff'].to_excel(writer,
                                                                sheet_name='fee_diff',
                                                                index=False,
                                                                merge_cells=False)

                my_response["xlsx_files"][country][rooles] = new_xlsx_path
        my_response["exit_is_ok"] = True
    except Exception as inst:
        my_response["exit_message"] = f'{__name__}: виникла помилка {inst}.'

    return my_response


#  ==================================== 2022 start (1) ================
def returned_not_added(in_data_folder=r"D:\_\BI\==data=="):
    RETURNED_NOT_ADDED = ["eve", "ret", "rei"]   # 2022
    path_to_files_folder = Path(in_data_folder)
    z = files_reading(path_to_files_folder, patterns=RETURNED_NOT_ADDED)

    if not z['exit_is_ok']:
        print(z['exit_message'], 'from  ==path_to_files_folder==')
        return z
    #
    #
    #   events columns:
    #        snapshot-date   transaction-type   fnsku   sku   product-name
    #        fulfillment-center-id   quantity   disposition
    #
    #        snapshot-date :6   transaction-type   fnsku  :1   sku   :2   product-name :3
    #        fulfillment-center-id :4   quantity :9   disposition :5
    #
    #
    #   return columns:
    #      return-date   order-id    sku   asin   fnsku   product-name
    #      quantity  fulfillment-center-id   detailed-disposition   reason   status
    #      license-plate-number   customer-comments
    #
    #      return-date :6   order-id :8    sku :2   asin   fnsku :1   product-name :3
    #      quantity :9   fulfillment-center-id :4   detailed-disposition :5   reason   status
    #      license-plate-number   customer-comments :13
    #
    #
    #   reimburse columns:
    #       approval-date   reimbursement-id   case-id   amazon-order-id   reason
    #       sku   fnsku   asin   product-name   condition   currency-unit
    #       amount-per-unit   amount-total   quantity-reimbursed-cash
    #       quantity-reimbursed-inventory   quantity-reimbursed-total :10
    #       original-reimbursement-id   original-reimbursement-type
    #
    #
    #  FINAL COLUMNS:
    #       01_fnsku   02_sku   03_product_name   04_fulfillment_center_id	  05_disposition
    #	    06_date   07_type_e_r   08_order_id
    #       09_returned_q   10_reimbursed_q   11_to_quest_q   12_event_q
    #       13_customer_comments
    #
    #

    SORT_COLUMNS = ['01_fnsku', '04_ff_center_id', '05_disposition', '06_date']
    FINAL_COLUMNS = [
        '01_fnsku', '02_sku', '03_product_name',
        '04_ff_center_id', '05_disposition',
        '06_date', '06_group_date',
        '07_opertn_id', '08_order_id',
        '09_returned_q', '10_reimbursed_q', '11_event_q',
    ]

    events_shaping = dict(
        recolumn=dict(
            fnsku='01_fnsku',   sku='02_sku',   product_name='03_product_name',
            fulfillment_center_id='04_ff_center_id',   disposition='05_disposition',
            snapshot_date='06_date', quantity='11_event_q',
        ),
        add_columns=('07_opertn_id', '08_order_id',
                     '09_returned_q', '10_reimbursed_q', '06_group_date',
                     ),
        sort_columns=SORT_COLUMNS,
    )

    returns_shaping = dict(
        recolumn=dict(
            fnsku='01_fnsku',   sku='02_sku',   product_name='03_product_name',
            fulfillment_center_id='04_ff_center_id', detailed_disposition='05_disposition',
            return_date='06_date',  license_plate_number='07_opertn_id',
            order_id='08_order_id', quantity='09_returned_q',
            customer_comments='13_customer_comments',
        ),
        add_columns=('10_reimbursed_q', '11_event_q', '06_group_date'),
        sort_columns=SORT_COLUMNS,
    )

    reimburse_shaping = dict(
        recolumn=dict(
            fnsku='01_fnsku',   sku='02_sku',   product_name='03_product_name',
            approval_date='06_date',
            reimbursement_id='07_opertn_id', amazon_order_id='08_order_id',
            quantity_reimbursed_total='10_reimbursed_q',
        ),
        add_columns=('04_ff_center_id', '05_disposition',
                     '09_returned_q', '11_event_q', '12_to_quest_q',
                     '13_customer_comments',),
        sort_columns=[],
    )

    df_shapings = dict(
        eve=events_shaping,
        ret=returns_shaping,
        rei=reimburse_shaping,
    )

    # take from events  'CustomerReturns' only!
    z['files']['eve'] = z['files']['eve'][z['files']['eve'].transaction_type == "CustomerReturns"]

    for key in z['files'].keys():
        print(key)
        z['files'][key] = reshape(z['files'][key].copy(), df_shapings[key])

    df_events = z['files']['eve']
    # filter 'reimburses' with  order_id
    df_reimburses_with_order_id = z['files']['rei'][
        ~z['files']['rei']['08_order_id'].isnull()][
            ['01_fnsku', '08_order_id', '10_reimbursed_q']].copy()

    df_reimbrs = z['files']['rei'].loc[df_reimburses_with_order_id.index]

    # Count returned items and reimbursed items for reimbursed order_ids.
    df_reimburses_items_sum = df_reimburses_with_order_id.groupby(by=[
        '01_fnsku', '08_order_id']).sum()

    df_returns = z['files']['ret']
    df_rets = z['files']['ret'][['01_fnsku', '08_order_id', '09_returned_q']].copy()
    df_returns_items_sum = df_rets.groupby(by=['01_fnsku', '08_order_id']).sum()
    # Delete returns if sums are equal

    reimbursed_returns_index = df_returns_items_sum.index.intersection(
        df_reimburses_items_sum.index)

    to_analize = df_returns_items_sum.loc[
        reimbursed_returns_index, :][
            (df_returns_items_sum.loc[reimbursed_returns_index, '09_returned_q']
             > df_reimburses_items_sum.loc[reimbursed_returns_index, '10_reimbursed_q']
             )]

    to_del = df_returns_items_sum.loc[
        reimbursed_returns_index, :][
            (df_returns_items_sum.loc[reimbursed_returns_index, '09_returned_q']
             <= df_reimburses_items_sum.loc[reimbursed_returns_index, '10_reimbursed_q']
             )]

    for key, row in to_del.reset_index().iterrows():
        df_returns = df_returns.drop(index=df_returns[
            (df_returns['01_fnsku'] == row['01_fnsku'])
            & (df_returns['08_order_id'] == row['08_order_id'])
        ].index
        )

    for key, row in to_analize.reset_index().iterrows():
        # set warehouses to reimburses:    #TODO: insert DATE to ''06_group_date'' field
        reimbursed_returns_filter = ((df_returns['01_fnsku'] == row['01_fnsku'])
                                     & (df_returns['08_order_id'] == row['08_order_id']))
        # take from first selected rows 'warehouse, return_date' values:
        warehouse, return_date = df_returns[reimbursed_returns_filter].iloc[
            [0]][['04_ff_center_id', '06_date']].values.tolist()[0]

        df_reimbrs.loc[
            df_reimbrs[(df_reimbrs['01_fnsku'] == row['01_fnsku'])
                       & (df_reimbrs['08_order_id'] == row['08_order_id'])].index,
            ['04_ff_center_id', '06_group_date']] = warehouse, return_date[:10]

    df_reimbrs = df_reimbrs[~df_reimbrs['04_ff_center_id'].isnull()]

    values = {
        '09_returned_q': 0, '10_reimbursed_q': 0, '11_event_q': 0,
    }

    df_reimbrs.fillna(value=values, inplace=True)
    df_returns.fillna(value=values, inplace=True)
    df_events.fillna(value=values, inplace=True)
    df_events.loc[:, '07_opertn_id'] = 'event'
    # group date filling for 'events' / 'return'  - the same date without timezone
    df_returns.loc[:, '06_group_date'] = df_returns['06_date'].str[:10]
    df_events.loc[:, '06_group_date'] = df_events['06_date'].str[:10]
    # for 'nothing to do' situation
    if not df_reimbrs.empty and '06_group_date' not in df_reimbrs.columns:
        df_reimbrs.loc[:, '06_group_date'] = None

    if df_reimbrs.empty:
        df_result = pd.concat(
            [df_returns[FINAL_COLUMNS],
             df_events[FINAL_COLUMNS],
             ],
            ignore_index=True).sort_values(by=SORT_COLUMNS)
    else:
        df_result = pd.concat(
            [df_returns[FINAL_COLUMNS],
             df_reimbrs[FINAL_COLUMNS],
             df_events[FINAL_COLUMNS],
             ],
            ignore_index=True).sort_values(by=SORT_COLUMNS)

    return df_result


def groping_files_reading(in_data_folder=None, ext: str = 'csv'):
    ...


def returnless_refunds(in_data_folder=r"D:\_\BI\==data=="):
    RETURLESS_REFUNDS = ['ret', 'rei']   # 2022
    path_to_files_folder = Path(in_data_folder)
    z = files_reading(path_to_files_folder, patterns=RETURLESS_REFUNDS)


...
#  ==================================== 2022 end (1) ================

if __name__ == "__main__":
    folder = Path(r'D:\_\BI')
    client = '0653_'
    clients_folder = folder / client
    df_final = returned_not_added(in_data_folder=clients_folder)

    df_final.to_excel(clients_folder / F'{client}_ORNABI.xlsx', index=False)
