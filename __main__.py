import os, sys, time, xlwings, pandas as pd
from sim_parser import sim_oven

if __name__ == '__main__':
    start_time = time.time()

    print('~ FlashMe: CO Daily Flash Report Builder ~\n\tcoded by @jjonamos\n')
    excel_file = os.path.join(
        os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__)),
        'Daily_Flash_Report.xlsm'
    )
    oven = sim_oven(excel_file)
    work_book = xlwings.Book(excel_file)
    work_book.macro('Clear_Sims')()
    for process in oven.process_folder_dictionary.keys():
        oven.cook(process)
        data = None
        with xlwings.App(visible = False):   
            work_sheet = work_book.sheets(process)
            work_sheet.range('A2').options(header = False, index = False).value = oven.cooked_sim_list
            headers = work_sheet.range(work_sheet.tables[process + '_SIMs']).expand('right').value[0]
            miss_col_index = headers.index('Miss')
            operator_col_index = headers.index('Operator Follow-up Miss')
            false_col_index = headers.index('False Resolution Miss')
            sla_col_index = headers.index('SLA Miss')
            miss_column = work_sheet.range(work_sheet.tables[process + '_SIMs']).expand('down').columns[miss_col_index].value[1:]
            operator_column = work_sheet.range(work_sheet.tables[process + '_SIMs']).expand('down').columns[operator_col_index].value[1:]
            false_column = work_sheet.range(work_sheet.tables[process + '_SIMs']).expand('down').columns[false_col_index].value[1:]
            sla_column = work_sheet.range(work_sheet.tables[process + '_SIMs']).expand('down').columns[sla_col_index].value[1:]
            data = {
                'Miss': miss_column,
                'Operator Follow-up Miss': operator_column,
                'False Resolution Miss': false_column,
                'SLA Miss': sla_column
            }
        missDf = pd.DataFrame(data)
        CiNciDf = oven.getCiNci(missDf)
    # oven.cook('FIF')
    # with xlwings.App(visible=False):
    #     work_sheet = work_book.sheets('FIF')
    #     work_sheet.range('A2').options(header = False, index = False).value = oven.cooked_sim_list
    # oven.labels.save()
    work_book.macro('Generate_Email_Report')()
    work_book.save()
    #work_book.close()

    process_time = time.time() - start_time
    process_minutes = process_time // 60
    process_seconds = process_time - process_minutes * 60
    print(f'Process Time: {process_minutes} minutes, {process_seconds} seconds.')