"""
Author: xxjjtttt
Date: 2024-01-22 17:11:56
LastEditTime: 2024-01-22 23:13:25
LastEditors: xxjjtttt
Description: https://github.com/xxjjtttt
FilePath: \\python\\xlsx master\\Task\\Task.py
屎山代码什么时候重构
"""

# python -m pip install upgrade pip
# pip install openpyxl==3.1.2

import json
from Main.File.Database import Database
from Main.File.Sourcefile import Sourcefile
from Main.File.Targetfile import Targetfile


class CoreTask:
    properties: dict = {}
    database: Database
    sourcefile: Sourcefile
    targetfile: Targetfile

    def __init__(self) -> None:
        # I don't want to bind properties to this class when I create this class
        pass

    def init(self) -> None:
        self.properties = json.load(open("./Properties/classify.json", "r", encoding="utf-8"))
        # self.properties = json.load(open("./Properties/count.json", "r", encoding="utf-8"))
        database_path: str = str((self.properties["task"]["file"]["database"]["path"]))
        database_active_sheet: list = list((self.properties["task"]["file"]["database"]["active_sheet"]))
        sourcefile_active_sheet: list = list((self.properties["task"]["file"]["sourcefile"]["active_sheet"]))
        sourcefile_path: str = str(self.properties["task"]["file"]["sourcefile"]["path"])
        targetfile_path: str = str(self.properties["task"]["file"]["targetfile"]["path"])
        key_column: int = int(self.properties["task"]["file"]["sourcefile"]["key_column"])
        index_column: int = int(self.properties["task"]["file"]["sourcefile"]["index_column"])
        have_header: bool = bool(self.properties["task"]["file"]["sourcefile"]["have_header"])
        headers: list = list(self.properties["task"]["file"]["targetfile"]["headers"])
        self.database = Database(database_path, database_active_sheet)
        self.sourcefile = Sourcefile(sourcefile_path, key_column, index_column, have_header, sourcefile_active_sheet)
        self.targetfile = Targetfile(targetfile_path, key_column, index_column, have_header, headers)
        self.targetfile.save()

    def close(self) -> None:
        self.database.close()
        self.sourcefile.close()
        self.targetfile.close()

    def classify(self) -> None:
        header_list:list = self.sourcefile.get_header_list()
        self.targetfile.set_header(header_list)
        for source_name in self.sourcefile.get_sheetnames():
            self.sourcefile.change_working_sheet(source_name)
            for source_row in range(self.sourcefile.get_max()[0]):
                source = self.sourcefile.get_key_value()
                is_classified = False
                if source == ".":
                    self.sourcefile.next_working_row()
                    continue
                for data_name in self.database.get_sheetnames():
                    self.database.change_working_sheet(data_name)
                    for data_row in range(self.database.get_max()[0]):
                        data_list = self.database.get_data_list()
                        for index, data in enumerate(data_list):
                            if index == 0:
                                continue
                            if data == source:
                                self.targetfile.change_working_sheet(data_list[0])
                                source_list = self.sourcefile.get_source_list()
                                self.targetfile.add_row(source_list)
                                self.sourcefile.next_working_row()
                                is_classified = True
                                break
                        if is_classified:
                            break
                    if is_classified:
                        break
                    else:
                        self.sourcefile.next_working_row()
        self.sourcefile.throw_rubbish()

    # count 任务完成
    def count(self) -> None:
        sheetname_list = self.targetfile.get_sheetnames()
        for sheetname in sheetname_list:
            self.database.change_working_sheet(sheetname)
            self.targetfile.change_working_sheet(sheetname)
            self.sourcefile.change_working_sheet(sheetname)
            for data_row in range(self.database.get_max()[0]):
                self.sourcefile.reset_working_row()
                data_list = self.database.get_data_list()
                count: int = 0
                for source_row in range(self.sourcefile.get_max()[0]):
                    source = self.sourcefile.get_key_value()
                    if source == ".":
                        self.sourcefile.next_working_row()
                        continue
                    for index, data in enumerate(data_list):
                        if index == 0:
                            continue
                        if data == source:
                            self.sourcefile.tag_row()
                            count += 1
                            break
                    self.sourcefile.next_working_row()
                self.targetfile.add_row([data_list[0], count])
                self.targetfile.save()


# a = CoreTask()
# a.init()
# a.classify()
# a.targetfile.save()
# # count
# # a.count()
# # a.targetfile.save()
