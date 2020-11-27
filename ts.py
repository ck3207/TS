import xlrd
import xlwt

class TsDataDeal:
    def __init__(self, workbook, keyword_list):
        self.keyword_list = keyword_list
        self.workbook = xlrd.open_workbook(filename=workbook)
        self.keyword_dict = {}
        for k in self.keyword_list:
            self.keyword_dict.setdefault(k, [])

    def classify(self, sheet_name="修改单导出表"):
        """读取excel中，ts单数据，并将其分类"""
        target_sheet = self.workbook.sheet_by_name(sheet_name)
        nrows = target_sheet.nrows
        ncols = target_sheet.ncols
        for r in range(nrows):
            for c in range(ncols):
                cell_value = target_sheet.cell(r, c).value
                for k in self.keyword_list:
                    if c == 0 and target_sheet.cell(r, c+1).value.count(k) >= 1:
                        self.keyword_dict.get(k).append({cell_value: target_sheet.cell(r, c+1).value})
        return self.keyword_dict

    def printf(self):
        for k, v in self.keyword_dict.items():
            print(k)
            for each in v:
                print(each)

    def save_to_excel(self, book_name="需求汇总.xls", needs_common_data=True):
        """将ts单信息数据录入到excel表格中"""
        work_book = xlwt.Workbook()
        for k, v_list in self.keyword_dict.items():
            if needs_common_data == False and k == self.keyword_list[0]:
                continue
            work_sheet = work_book.add_sheet(k)
            i, j = 0, 0
            for each_ts in v_list:
                for ts, info in each_ts.items():
                    work_sheet.write(i, j, ts)
                    work_sheet.write(i, j+1, info)
                i += 1
        work_book.save(book_name)

    def combine_common_ts(self):
        """整合通用需求"""
        common_list = []
        for i, info in enumerate(self.keyword_list):
            if i == 0:
                common_list = self.keyword_dict.get(info)
            else:
                self.keyword_dict.setdefault(info, self.keyword_dict.get(info).extend(common_list))
        return self.keyword_dict

if __name__ == "__main__":
    ts_data_deal = TsDataDeal("ModifyDetail2012255746.xlsx", ["通用", "中邮", "万和", "太平洋", "财达", "联储"])
    ts_data_deal.classify()
    # ts_data_deal.printf()
    ts_data_deal.combine_common_ts()
    ts_data_deal.save_to_excel(book_name="需求汇总.xls", needs_common_data=False)
