import pandas, xlwings
from api_caller import maxis
from json import loads
from labels_handler import label

class sim_oven:
    def __init__(self, excel_file):
        self.process = ''
        self.cooked_sim_list = pandas.DataFrame(columns = ['Issue Url', 'Title', 'Labels', 'Data Status', 'Operator Follow-up Miss', 'False Resolution Miss', 'SLA Miss'])
        self.cooked_sim_list_ORSA = pandas.DataFrame(columns = ['ID','Issue Url', 'Title', 'Labels', 'Next Step Action', 'Data Status', 'Operator Follow-up Miss', 'False Resolution Miss', 'SLA Miss'])
        self.cooked_list = None
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

    def cook(self):
        self.cooked_list.drop(self.cooked_list.index, inplace=True) if self.cooked_list is not None else None
        self._cooked_sim_list_row = 0
        self.__init_sim_endpoint()
        self.__cook()

    def __init_sim_endpoint(self):
        date_range_map = {
            'FIF' : '[NOW-123DAYS TO NOW]',
            'ORSA_Intervention' : '[NOW-123DAYS TO NOW]'
        }
        process_id = (
            self.process_folder_dictionary[self.process] if self.process != ''
            else '+OR+'.join(value for value in self.process_folder_dictionary.values())
        )
        sim_status = '' if self.process == "ORSA_Warnings_Miss" else "+status:(Resolved)"
        date_range = date_range_map.get(self.process, '[NOW-28DAYS TO NOW]')
        sort_order = 'lastUpdatedDate+desc'
        folder = "labels" if self.process == "ORSA_Warnings_Miss" else "containingFolder"
        self._sim_endpoint = f'issues?q={folder}:({process_id}){sim_status}+createDate:({date_range})&sort={sort_order}'

    def __cook(self):
        valid_processes = {"ORSA_Valids", "ORSA_Invalids", "ORSA_Warnings_Miss"}
        self.cooked_list = self.cooked_sim_list_ORSA if self.process in valid_processes else self.cooked_sim_list
        while True:
            self.__update_raw_sim_list()
            self.__cook_sims()

            if not self._start_token:
                break
    
    def __update_raw_sim_list(self):
        self._maxis.get(self._sim_endpoint + self._start_token)
        self._raw_sim_list = loads(self._maxis.response)
        self._start_token = f'&startToken={self._raw_sim_list['startToken']}' if self._raw_sim_list['startToken'] else ''

    def __cook_sims(self):
        for raw_sim in self._raw_sim_list['documents']:
            valid_processes = {"ORSA_Valids", "ORSA_Invalids", "ORSA_Warnings_Miss"}
            
            issue_url = f'https://issues.amazon.com/issues/{raw_sim["aliases"][0]["id"]}'
            self.cooked_list.at[self._cooked_sim_list_row, 'Issue Url'] = issue_url
            self.cooked_list.at[self._cooked_sim_list_row, 'Title'] = raw_sim['title']
            self.cooked_list.at[self._cooked_sim_list_row, 'Labels'] = self.__cook_labels(raw_sim['labels'])

            if self.process in valid_processes:
                self.cooked_list.at[self._cooked_sim_list_row, 'Next Step Action'] = raw_sim['next_step']['action']

            # Process custom fields if they exist
            if 'customFields' in raw_sim:
                self.__process_custom_fields(raw_sim['customFields'])

            self._cooked_sim_list_row += 1

        total_found = self._raw_sim_list['totalNumberFound']
        print(f'Downloaded {self._cooked_sim_list_row}/{total_found} {self.process} SIMs')

    def __process_custom_fields(self, custom_fields):
        valid_processes = {"ORSA_Valids", "ORSA_Invalids", "ORSA_Warnings_Miss"}
        if 'checkbox' in custom_fields:
            for checkbox in custom_fields['checkbox']:
                checked_values = self.__cook_checkboxes(checkbox)
                checkbox_id = checkbox['id']

                if checkbox_id in {'operator_follow_up_miss', 'operator_follow_up_miss_'}:
                    self.cooked_list.at[self._cooked_sim_list_row, 'Operator Follow-up Miss'] = checked_values
                elif checkbox_id in {'false_resolution', 'miss', 'false_resolution_miss'}:
                    self.cooked_list.at[self._cooked_sim_list_row, 'False Resolution Miss'] = checked_values
                elif checkbox_id == 'sla_miss':
                    self.cooked_list.at[self._cooked_sim_list_row, 'SLA Miss'] = checked_values

        if 'string' in custom_fields:
            string_fields = custom_fields['string']
            self.cooked_list.at[self._cooked_sim_list_row, 'Data Status'] = self.__cook_data_status(string_fields)
            self.cooked_list.at[self._cooked_sim_list_row, 'SLA Miss'] = self.__cook_string_sla_miss(string_fields)
            if self.process in valid_processes:
                self.cooked_list.at[self._cooked_sim_list_row, 'ID'] = next(
                    (field.get("value") for field in string_fields if field["id"] == "contact_id"),
                    None
                )

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