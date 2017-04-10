import xlsxwriter


class Excel:
    def __init__(self, file_name, head_raw, head_raw_position=0):
        self.workbook = xlsxwriter.Workbook(file_name)
        self.worksheet = self.workbook.add_worksheet()
        self.bold = self.workbook.add_format({'bold': True})
        self.head_raw_position = head_raw_position
        self.content_raw_position = head_raw_position + 1
        self.write_head_raw(head_raw)

    def write_head_raw(self, head_raw):
        for col in range(len(head_raw)):
            self.worksheet.write(self.head_raw_position, col, head_raw[col], self.bold)

    def write_one_raw(self, raw_content):
        for col in range(len(raw_content)):
            self.worksheet.write(self.content_raw_position, col, raw_content[col])
        self.content_raw_position += 1

    def writer_multic_raws(self, raw_list):
        for i_raw in raw_list:
            self.write_one_raw(i_raw)

    def close(self):
        self.workbook.close()

