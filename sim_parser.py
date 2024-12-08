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

    def cook(self, process_name = '', iteration = 1):
        self.iteration = iteration
        if iteration == 1: 
            self._cooked_sim_list_row = 0
            self.running_total = 0
        self.checkpoint = 0
        self.__init_sim_endpoint(process_name)
        self.__cook(process_name)

    def __init_sim_endpoint(self, process_name):
        process_id = (
            self.process_folder_dictionary.get(process_name, '')
            if process_name
            else '+OR+'.join(value for value in self.process_folder_dictionary.values())
        )
        
        process_label_exclusions = [
            '226e8e9c-1486-451b-9609-07ef2c302b00',
            '95c6328f-8a83-49b0-a7dc-b47b73f1b4f7',
            '3a50ffcd-0084-4217-9b90-bbf0e3d8871f',
            'beadfedd-2c7c-4599-879f-93d435d27f34',
            '04e2cd43-f916-40bc-aba3-96936b05f4e5',
            '226e8e9c-1486-451b-9609-07ef2c302b00',
            '2fa0a86d-3465-4e49-8d2f-8ed682f4e615',
            '1759e83f-4d79-4d80-ab46-b85ace8d9e85',
            '8cb7f94b-6805-422b-bf32-5f2238432840',
            '84ed1b0f-65d8-43ff-af5b-1cb54557c56b',
            'bf860273-c56a-4a7a-8440-9f30dd07237e',
            '36b64a96-4457-4166-9553-f656f11dbeb4',
            '379bcb27-4812-4a5b-85d3-160ff0779060',
            '4ae952fa-ff2c-4f37-b6f1-59c70cd5e5a7',
            '44cb5c79-1dda-4b7a-9082-2ceddef0f1e5',
            '9f89815e-5125-474a-b637-75df57f81a16'
        ]
        exclusion_query = ' OR '.join(process_label_exclusions)

        date_ranges = {
            1: '[2024-01-01T07:00:00.000Z TO 2024-04-01T06:59:00.000Z]',
            2: '[2024-04-01T07:00:00.000Z TO 2024-07-01T06:59:00.000Z]',
            3: '[2024-07-01T07:00:00.000Z TO 2024-10-01T06:59:00.000Z]',
            4: '[2024-10-01T07:00:00.000Z TO NOW]'
        }
        date_range = date_ranges.get(self.iteration, '')

        folder_query = f'containingFolder:({process_id})'
        status_query = f'status:(Resolved)'
        date_query = f'createDate:({date_range})'
        title_exclusion = '-title:(partial OR pilot OR training OR test)'
        label_exclusion = f'-label:({exclusion_query})'
        sort_query = 'sort=lastUpdatedDate+desc'

        self._sim_endpoint = f'issues?q={folder_query}+{status_query}+{date_query}+{title_exclusion}+{label_exclusion}&{sort_query}'

    def __cook(self, process_name):
        self.__update_raw_sim_list()
        self.__cook_sims(process_name)
    
    def __update_raw_sim_list(self):
        self._maxis.get(self._sim_endpoint + self._start_token)
        self._raw_sim_list = loads(self._maxis.response)
        if self.checkpoint == 0:
            self.running_total += self._raw_sim_list['totalNumberFound']
            self.checkpoint = 1
        self._start_token = f'&startToken={self._raw_sim_list['startToken']}' if self._raw_sim_list['startToken'] else ''

    def __cook_sims(self, process_name):
        for raw_sim in self._raw_sim_list['documents']:
            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Issue Url'] = f'https://issues.amazon.com/issues/{raw_sim['aliases'][0]['id']}'
            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Title'] = raw_sim['title']
            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Labels'] = self.__cook_labels(raw_sim['labels'])
            if 'customFields' in raw_sim:
                if 'checkbox' in raw_sim['customFields']:
                    for checkbox in raw_sim['customFields']['checkbox']:
                        checked_values = self.__cook_checkboxes(checkbox)
                        if checkbox['id'] == 'operator_follow_up_miss' or checkbox['id'] == 'operator_follow_up_miss_':
                            self.cooked_sim_list.at[self._cooked_sim_list_row, 'Operator Follow-up Miss'] = checked_values
                        elif checkbox['id'] == 'false_resolution' or checkbox['id'] == 'miss' or checkbox['id'] == 'false_resolution_miss':
                            self.cooked_sim_list.at[self._cooked_sim_list_row, 'False Resolution Miss'] = checked_values                    
                        elif checkbox['id'] == 'sla_miss':
                            self.cooked_sim_list.at[self._cooked_sim_list_row, 'SLA Miss'] = checked_values
                if 'string' in raw_sim['customFields']:
                    self.cooked_sim_list.at[self._cooked_sim_list_row, 'Data Status'] = self.__cook_data_status(raw_sim['customFields']['string'])
                    self.cooked_sim_list.at[self._cooked_sim_list_row, 'SLA Miss'] = self.__cook_string_sla_miss(raw_sim['customFields']['string'])
            self._cooked_sim_list_row += 1
        print(f'Downloaded {self._cooked_sim_list_row}/{self.running_total} {process_name} SIMs | iteration {self.iteration}')
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
            if string['id'] == 'data_status': cooked_data_status += string['value']; break
        return cooked_data_status
    
    def __cook_checkboxes(self, checkbox_list):
        cooked_checkboxes = ''
        for checkbox in checkbox_list['value']:
            if checkbox['checked'] == True: cooked_checkboxes += f'{checkbox['value']},'
        return cooked_checkboxes
    
    def __cook_string_sla_miss(self, string_list):
        cooked_sla_miss = ''
        for string in string_list:
            if string['id'] == 'sla_miss' and string['value'] != 'N/A': cooked_sla_miss = string['value']; break
        return cooked_sla_miss