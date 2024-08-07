import pandas, xlwings
from api_caller import maxis
from json import loads
from labels_handler import label

class sim_oven:
    def __init__(self, excel_file):
        self.cooked_sim_list = pandas.DataFrame(columns = ['Issue Url', 'Title', 'Labels', 'Data Status', 'Operator Follow-up Miss', 'False Resolution Miss', 'SLA Miss'])
        self.labels = label(excel_file)
        self._maxis = maxis()
        self._sim_endpoint = self._start_token = ''
        self._cooked_sim_list_row = None
        self.process_folder_dictionary = {}
        self.__init_process_folder_dictionary(excel_file)

    def __init_process_folder_dictionary(self, excel_file):
        with xlwings.App(visible = False):
            work_sheet = xlwings.Book(excel_file).sheets('Py_Variables')
            key_list = [str(entry.value) for entry in work_sheet['Process_Folder_Dictionary[Key]']]
            value_list = [str(entry.value) for entry in work_sheet['Process_Folder_Dictionary[Value]']]
            for key, value in zip(key_list, value_list): self.process_folder_dictionary[key] = value

    def cook(self, process_name = ''):
        self.cooked_sim_list.drop(self.cooked_sim_list.index, inplace = True)
        self._cooked_sim_list_row = 0
        self.__init_sim_endpoint(process_name)
        self.__cook(process_name)

    def __init_sim_endpoint(self, process_name):
        date_range_map = {
            'AFS' : '[NOW-28DAYS TO NOW-2DAYS]',
            'FIF' : '[NOW-123DAYS TO NOW-1DAYS]',
            'ORSA_Intervention' : '[NOW-123DAYS TO NOW-1DAYS]'
        }
        process_id = (
            self.process_folder_dictionary[process_name] if process_name != ''
            else '+OR+'.join(value for value in self.process_folder_dictionary.values())
        )
        sim_status = 'Resolved'
        date_range = date_range_map.get(process_name, '[NOW-28DAYS TO NOW-1DAYS]')
        sort_order = 'lastUpdatedDate+desc'
        self._sim_endpoint = f'issues?q=containingFolder:({process_id})+status:({sim_status})+createDate:({date_range})&sort={sort_order}'

    def __cook(self, process_name):
        self.__update_raw_sim_list()
        self.__cook_sims(process_name)
    
    def __update_raw_sim_list(self):
        self._maxis.get(self._sim_endpoint + self._start_token)
        self._raw_sim_list = loads(self._maxis.response)
        self._start_token = f'&startToken={self._raw_sim_list['startToken']}' if self._raw_sim_list['startToken'] else ''

    def __cook_sims(self, process_name):
        for raw_sim in self._raw_sim_list['documents']:
            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Issue Url'] = f'https://issues.amazon.com/issues/{raw_sim['aliases'][0]['id']}'
            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Title'] = raw_sim['title']
            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Labels'] = self.__cook_labels(raw_sim['labels'])
            if 'customFields' in raw_sim:
                self.cooked_sim_list.at[self._cooked_sim_list_row, 'Data Status'] = self.__cook_data_status(raw_sim['customFields']['string'])
                if 'checkbox' in raw_sim['customFields']:
                    for checkbox in raw_sim['customFields']['checkbox']:
                        checked_values = self.__cook_checkboxes(checkbox)
                        if checkbox['id'] == 'operator_follow_up_miss' or checkbox['id'] == 'operator_follow_up_miss_':
                            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Operator Follow-up Miss'] = checked_values
                        elif checkbox['id'] == 'false_resolution' or checkbox['id'] == 'miss' or checkbox['id'] == 'false_resolution_miss':
                            self.cooked_sim_list.at[self._cooked_sim_list_row, 'False Resolution Miss'] = checked_values                    
                        elif checkbox['id'] == 'sla_miss':
                            self.cooked_sim_list.at[self._cooked_sim_list_row, 'SLA Miss'] = checked_values
            self._cooked_sim_list_row += 1
        print(f'Downloaded {self._cooked_sim_list_row}/{self._raw_sim_list['totalNumberFound']} {process_name} SIMs')
        if self._start_token != '': self.__cook(process_name)

    def __cook_labels(self, label_id_list):
        cooked_labels = new_label_id_list = ''
        for label_id in label_id_list:
            if not label_id['id'] in self.labels.dictionary: new_label_id_list += f'{label_id['id']}+OR+'
        if new_label_id_list != '':
            self._maxis.get(f'labels?q=id:({new_label_id_list[:-4]})') #Remove trailing '+OR+'
            new_label_list = loads(self._maxis.response)
            for new_label in new_label_list['documents']:
                self.labels.dictionary[new_label['id']] = new_label['label'][0]['text']
        for label_id in label_id_list:
            cooked_labels += f'{self.labels.dictionary[label_id['id']]},'
        return cooked_labels

    def __cook_data_status(self, string_list):
        cooked_data_status = ''
        for string in string_list:
            if string['id'] == 'data_status': cooked_data_status += string['id']; break
        return cooked_data_status
    
    def __cook_checkboxes(self, checkbox_list):
        cooked_checkboxes = ''
        for checkbox in checkbox_list['value']:
            if checkbox['checked'] == True: cooked_checkboxes += f'{checkbox['value']},'
        return cooked_checkboxes