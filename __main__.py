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
        for step in range(1, 5):
            oven.cook(process, step)
            
            # Access the correct worksheet
            work_sheet = work_book.sheets(process)
            
            # Find the next empty row in the worksheet
            last_row = work_sheet.range('A' + str(work_sheet.cells.last_cell.row)).end('up').row
            
            # Determine where to paste the data
            start_row = last_row + 1 if last_row >= 2 else 2
            
            # Paste data from the current cooked_sim_list
            work_sheet.range(f'A{start_row}').options(header=False, index=False).value = oven.cooked_sim_list

            # Clear the DataFrame after pasting
            oven.cooked_sim_list.drop(oven.cooked_sim_list.index, inplace=True)


    # oven.cook('ORSA_Valids')
    # with xlwings.App(visible=False):
    #     work_sheet = work_book.sheets('ORSA_Valids')
    #     work_sheet.range('A2').options(header = False, index = False).value = oven.cooked_sim_list

    oven.labels.save()
    work_book.macro('Generate_Email_Report')()
    work_book.save()
    #work_book.close()

    process_time = time.time() - start_time
    process_minutes = process_time // 60
    process_seconds = process_time - process_minutes * 60
    print(f'Process Time: {process_minutes} minutes, {process_seconds} seconds.')