import xlrd
import xlwt

class GetInfoFromExcel:
    """从excel文档获取bug记录"""
    def __init__(self, filename):
        self.workbook = xlrd.open_workbook(filename=filename)

    def get_data(self, table):
        """根据bug模板中的数据结构，获取对应数据，并返回"""
        row, cols = table.nrows, table.ncols
        headers = {}
        data = {}
        for i in range(row):
            for j in range(cols):
                if i == 0:
                   headers.setdefault(j, table.cell(i, j).value)
                else:
                    value = table.cell(i, j).value
                    data.setdefault("{0}-{1}".format(i, j), value)
                # if headers.get(j) == "计划完成日期" and i != 0 :
                #     pass
        return headers, data


class WriteToExcel:
    """写入excel"""
    def __init__(self, filename, sheetname):
        self.filename = filename
        self.f = xlwt.Workbook(encoding="utf-8")
        self.table = self.f.add_sheet(sheetname=sheetname, cell_overwrite_ok=True)

    def write_via_row(self, data_list, startrow, startcol, keyname=[]):
        """
        将传入的参数以行的形式进行写入；
        若数据格式为 [a, b, c,d,e ...], 则写入时，固定行，将列表的数据写入不同列
        若数据格式为 [[a, b, c], [d, e, f] ...], 则写入时，每个元素列表写入不同的行，元素列表中的每个元素的数据写入不同列
        :param data_list: 数据值的列表
        :param startrow: 列表写入的开始行，从0开始
        :param startcol: 列表写入的开始列，从0开始
        :param keyname: 若传入的data_list格式为：[{k1:v1, k2:v2}], 可根据传入的 keyname的列表数据， 依次获取写入数据
        :return: 
        """
        for i, value in enumerate(data_list):
            # data_list = [[v1, v2, v3], [v4, v, v6] ...]
            if isinstance(value, list):
                for j, data in enumerate(value):
                    self.table.write(startrow+i, startcol+j, data)
            # data_list = [{v1:v11, v2:v22, v3:v33}, {v4:v44, v5:v55, v6:v66} ...]
            elif isinstance(value, dict):
                for j, name in enumerate(keyname):
                    self.table.write(startrow+i, startcol+j, value.get(name))
                if not keyname:
                    self.table.write(startrow, startcol + i, value)
            else:
                self.table.write(startrow, startcol+i, value)
        return

    def write_via_column(self, data_list, startrow, startcol):
        """
        将传入的参数以行的形式进行写入；
        若数据格式为 [a, b, c,d,e ...], 则写入时，固定列，将列表的数据写入不同行
        若数据格式为 [[a, b, c], [d, e, f] ...], 则写入时，每个元素列表写入不同的列，元素列表中的每个元素的数据写入不同行
        :param data_list: 数据值的列表
        :param startrow: 列表写入的开始行，从0开始
        :param startcol: 列表写入的开始列，从0开始
        :return: 
        """
        for i, value in enumerate(data_list):
            if isinstance(value, list):
                for j, data in enumerate(value):
                    self.table.write(startrow+j, startcol+i, data)
            else:
                self.table.write(startrow+i, startcol, value)
        return

    def close_file(self):
        self.f.save(filename_or_stream=self.filename)


if __name__ == "__main__":
    get_info_from_excel = GetInfoFromExcel(filename="config/jira_template.xls")
    table = get_info_from_excel.workbook.sheet_by_name("bugs")


