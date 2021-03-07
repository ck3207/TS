import re
import sys

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

    @staticmethod
    def print_kv_via_defined_word(data, connected_word="-"):
        if isinstance(data, dict):
            for k, v in data.items():
                if "docker" in k:
                    k = "/".join(k.split("/")[:-1])
                if v.endswith("online.zip") and "h5-fans-cnpsec" in v:
                    print(k + connected_word + v[:-len("online.zip")] + "offline.zip")
                elif v.endswith("offline.zip") and "h5-fans-cnpsec" in v:
                    print(k + connected_word + v[:-len("offline.zip")] + "online.zip")
                print(k + connected_word + v)

class TsDataDeal:
    def __init__(self, workbook, keyword_list=["通用", "中邮", "万和", "太平洋", "财达", "联储", "华创"]):
        self.keyword_list = keyword_list    # 券商标识
        self.workbook = xlrd.open_workbook(filename=workbook)
        self.keyword_dict = {}  # 需求按券商归类
        for k in self.keyword_list:
            self.keyword_dict.setdefault(k, [])

        self.all_integration_packages = []  # 所有集成包列表
        self.target_integration_packages = {}  # 筛选后的最新集成包数据, 格式为：{package: YYMMddhhmmss-SVN12345}

    def classify(self, sheet_name="修改单导出表"):
        """读取excel中，ts单数据(修改单编号， 修改原因)，并将其分类"""
        target_sheet = self.workbook.sheet_by_name(sheet_name)
        nrows = target_sheet.nrows
        ncols = target_sheet.ncols
        if target_sheet.cell(0, 0).value != "修改单编号" or target_sheet.cell(0, 1).value != "修改原因":
            raise Exception("读取的excel不符合规范！Excel的sheet名称需为修改单导出表，\
            单元格需从第0行开始，单元格第0列需为修改单编号列，单元格第一列需为修改原因列。")
        for r in range(nrows):
            for c in range(ncols)[:2]:
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
            work_sheet.write(0, 0, "修改单号")
            work_sheet.write(0, 1, "修改原因")
            i, j = 1, 0
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

    def get_all_integration_package(self, sheet_name="修改单导出表"):
        target_sheet = self.workbook.sheet_by_name(sheet_name)
        nrows = target_sheet.nrows
        ncols = target_sheet.ncols
        target_nc = None    # 集成说明列
        for nc in range(ncols):
            if target_sheet.cell(0, nc).value == "集成说明":
                target_nc = nc
                break
        if not target_nc:
            raise Exception("读取的excel不符合规范！Excel的sheet名称需为修改单导出表，\
            单元格需从第0行开始，单元格第0行的第n列需存在名为集成说明列。")
        for nr in range(nrows):
            cell_value = target_sheet.cell(nr, target_nc).value
            self.all_integration_packages.append(cell_value)
        return self.all_integration_packages

    def get_integration_packages(self, integration_packages_list=[]):
        """通过集成说明字段数据中，匹配提取出版本信息;
        根据版本路径的dirname为key， basename为value，更新同一key的最新value"""
        reg_packages = re.compile("[\/\.:\u4e00-\u9fa5\w\[\]-]+\.zip")  # 匹配集成包版本路径
        # reg_packages = re.compile("(\S+SVN\S{1,30})$")  # 匹配集成包版本路径
        reg_packages_docker = re.compile("[\/\._:\u4e00-\u9fa5\w\[\]-]+SVN\S+$")  # 匹配集成包版本路径
        for integration_package in integration_packages_list:
            for integration_package in integration_package.split("\n"):
                integration_package = integration_package.strip()
                # 对于docker的集成包路径特殊化处理
                if "docker" in integration_package:
                    reg_packages = reg_packages_docker
                reg_packages_result = reg_packages.findall(integration_package)
                for package in reg_packages_result:
                    # 集成包版本
                    version = package.split("/")[-1]
                    prefix_package_path = "/".join(package.split("/")[:-1])
                    # 对于docker的集成包路径特殊化处理
                    # artifactory.hundsun.com/xxx/smartwhale/smart:SVN15730-20201130152218
                    # artifactory.hundsun.com/xxx/smartwhale/smart-data:SVN15730-20201130152409
                    if "docker" in package:
                        # 集成包路径前缀
                        tmp_list = package.split("/")[:-1]
                        tmp_list.append(version.split(":")[0])
                        prefix_package_path = "/".join(tmp_list)

                    # 若匹配的包不在目标数据中则添加，若存在与目标数据中，则比较两个版本哪个高，并保留版本高的数据
                    # self.prefix_package_path.setdefault(package_name, prefix_package_path)
                    if not self.target_integration_packages.get(prefix_package_path):
                        self.target_integration_packages.setdefault(prefix_package_path, version)
                    elif self.target_integration_packages.get(prefix_package_path) < version:
                        self.target_integration_packages.update({prefix_package_path: version})

        return self.target_integration_packages


if __name__ == "__main__":
    try:
        filename = sys.argv[1]
    except:
        filename = "ModifyDetail-642286889.xlsx"
    ts_data_deal = TsDataDeal(filename, ["通用", "中邮", "万和", "太平洋", "财达", "联储", "华创"])
    ts_data_deal.classify()
    # ts_data_deal.printf()
    ts_data_deal.combine_common_ts()
    ts_data_deal.save_to_excel(book_name="需求汇总.xls", needs_common_data=False)

    ts_data_deal.get_integration_packages(ts_data_deal.get_all_integration_package())
    Print.print_kv_via_defined_word(data=ts_data_deal.target_integration_packages, connected_word="/")