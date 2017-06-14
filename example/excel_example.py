class ExcelExample:
        def __init__(self, file_name, head=[]):
            """

            :file_name: excel file name.
            :param head: list data.It contain the first line of excel file.
            """
            self.file_name = file_name if '.xlsx' in file_name else file_name+'.xlsx'
            self.excel_init = Excel(file_name, head)
            self.head = head

        def write_one_line(data_dict):
            self.excel_init.write_one_raw(self.format(data_dict))

        def finish(self):
            self.excel_init.close()

        def format(self, message):
            """
            Map to the first list of the excel file.
            :param message: dict.Values are map to columns by the keys which must be same as self.head
            :return: list
            """
            data = []
            for key in MESSAGE_HEAD:
                data.append(message[key])
            return data

if __name__ == '__main__':
    excel_head = ['first', 'second', 'third']
    excel_data = [ 
        {'first': 1, 'second': 2, 'third': 3},
        {'first': 4, 'second': 5, 'third': 6},
        {'first': 7, 'second': 8, 'third': 9}
    ]
    excel_example = ExcelExample('example.xlsx', head=excel_head)
    for i_excel_data in excel_data:
        excel_example.write_one_line(i_excel_data)
    excel_exmaple.finish()
