import os, sys, time, xlwings
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
        oven.process = process
        oven.cook()
        with xlwings.App(visible = False):   
            work_sheet = work_book.sheets(process)
            work_sheet.range('A2').options(header = False, index = False).value = oven.cooked_list

    # oven.process = ('ORSA_Warnings_Miss')
    # oven.cook()
    # with xlwings.App(visible=False):
    #     work_sheet = work_book.sheets('ORSA_Warnings_Miss')
    #     work_sheet.range('A2').options(header = False, index = False).value = oven.cooked_list

    oven.labels.save()
    work_book.macro('Generate_Email_Report')()
    work_book.save()
    #work_book.close()

    process_time = time.time() - start_time
    process_minutes = process_time // 60
    process_seconds = process_time - process_minutes * 60
    print(f'Process Time: {process_minutes} minutes, {process_seconds} seconds.')