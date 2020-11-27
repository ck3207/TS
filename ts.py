import re

import xlrd
import xlwt

class Print:
    """打印各类数据"""
    def print_data(self, data):
        if isinstance(data, tuple) or isinstance(data, list):
            for each in data:
                print(each)
        elif isinstance(data, dict):
            for k, v in data.items():
                print(k, v)

    def print_kv_via_defined_word(self, data, connected_word=["-", ".zip"]):
        if isinstance(data, dict):
            for k, v in data.items():
                print(k + connected_word[0] + v + connected_word[1])

class TsDataDeal:
    def __init__(self, workbook, keyword_list):
        self.keyword_list = keyword_list    # 券商标识
        self.workbook = xlrd.open_workbook(filename=workbook)
        self.keyword_dict = {}  # 需求归类
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
                    # 判断修改说明里面是否标识了券商关键字
                    if c == 0 and target_sheet.cell(r, c+1).value.count(k) >= 1:
                        # 对符合的数据进行归类数据
                        self.keyword_dict.get(k).append({cell_value: target_sheet.cell(r, c+1).value})
        return self.keyword_dict

    def printf(self):
        for k, v in self.keyword_dict.items():
            print(k)
            for each in v:
                print(each)

    def save_to_excel(self, book_name="需求汇总.xls", needs_common_data=True):
        """将ts单信息数据录入到excel表格中
        @book_name
        @needs_common_data 是否导出通用的需求，True 导出通用，False 不导出通用
        """
        work_book = xlwt.Workbook()
        for k, v_list in self.keyword_dict.items():
            # 若无需导出通用需求 或 当前查看的为通用需求时略过
            if needs_common_data == False and k == self.keyword_list[0]:
                continue
            work_sheet = work_book.add_sheet(k)
            i, j = 0, 0
            for each_ts in v_list:
                for ts, info in each_ts.items():
                    # 表格第一行写ts修改单编号， 第二行写修改单说明
                    work_sheet.write(i, j, ts)
                    work_sheet.write(i, j+1, info)
                i += 1
        work_book.save(book_name)

    def combine_common_ts(self):
        """整合通用需求，遍历通用需求，写入各个独立的券商需求里"""
        common_list = []
        for i, info in enumerate(self.keyword_list):
            if i == 0:
                common_list = self.keyword_dict.get(info)
            else:
                self.keyword_dict.setdefault(info, self.keyword_dict.get(info).extend(common_list))
        return self.keyword_dict


class GetLatestIntegrationPackages(Print):
    """根据ts导出的集成包中，找出各个包路径下的最新包"""
    def __init__(self, workbook=""):
        self.workbook = xlrd.open_workbook(filename=workbook)
        self.all_integration_packages = []  # 所有集成包列表
        self.target_integration_packages = {}  # 筛选后的最新集成包列表, 格式为：{package: YYMMddhhmmss-SVN12345}

    def distinct_integration_package(self, sheet_name="修改单导出表"):
        target_sheet = self.workbook.sheet_by_name(sheet_name)
        nrows = target_sheet.nrows
        ncols = target_sheet.ncols
        if ncols > 1:
            raise Exception("导出的修改单列表不符合规范，导出文件应只有集成版本一列。")
        for nr in range(nrows):
            cell_value = target_sheet.cell(nr, 0).value
            self.all_integration_packages.append(cell_value)
        return self.all_integration_packages

    def get_integration_packages(self, integration_packages_list=[]):
        reg_packages = re.compile("[\w-]+\.zip")  # 匹配集成包版本路径
        # reg_package_split = re.compile("([\w-]+)-(\d+)-(SVN\d+)\.zip")  # 匹配集成包路径各个部分(包名,集成时间，SVN版本)
        reg_package_split = re.compile("([\w-]+)-(\d+-SVN\d+)\.zip")  # 匹配集成包路径各个部分(包名,版本)
        for integration_package in integration_packages_list:
            reg_packages_result = reg_packages.findall(integration_package)
            for package in reg_packages_result:
                # package_name, integration_time, svn_version = reg_package_split.match(package).groups()
                package_name, version = reg_package_split.match(package).groups()
                # print(package_name, version)
                # 若匹配的包不在目标数据中则添加，若存在与目标数据中，则比较两个版本哪个高，并保留版本高的数据
                if not self.target_integration_packages.get(package_name):
                    self.target_integration_packages.setdefault(package_name, version)
                elif self.target_integration_packages.get(package_name) < version:
                    self.target_integration_packages.update({package_name: version})

        return self.target_integration_packages


if __name__ == "__main__":
    # ts_data_deal = TsDataDeal("ModifyDetail2012255746.xlsx", ["通用", "中邮", "万和", "太平洋", "财达", "联储"])
    # ts_data_deal.classify()
    # # ts_data_deal.printf()
    # ts_data_deal.combine_common_ts()
    # ts_data_deal.save_to_excel(book_name="需求汇总.xls", needs_common_data=False)

    integration = GetLatestIntegrationPackages(workbook="ModifyDetail865521882.xlsx")
    integration.get_integration_packages(integration.distinct_integration_package())
    integration.print_kv_via_defined_word(data=integration.target_integration_packages)