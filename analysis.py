import re

import xlwings as xw

###########本文件的作用#######
#获取表格的有效行列数确定解析范围
#根据解析范围排查未填写空格，并提示责任人修改，然后程序结束（这个过程测试人员要修改到正确才可以）
#解析范围内的内容都填写完成，开始解析，并将新用例生成新表格
from interval import Interval


class ExcelReportAnalysis(object):
    def __init__(self):
        #打开已有的工作表
        # visible参数控制创建文件时是否可见，True是可见，False是不可见
        app = xw.App(visible=False, add_book=False)
        self.app = app
        # 不显示Excel消息框
        self.app.display_alerts = False
        # 关闭屏幕更新,可加快宏的执行速度
        self.app.screen_updating = False
        # path = r'C:\Users\test\Desktop\Testcase2jira Sypnase\Intretech-VV-0202 咕咕机_Android_V3.4.0测试报告_20211021.xlsx'
        # path = r'C:\Users\test\Desktop\Testcase2jira Sypnase\TestPlanReport（测试计划执行报告）.xls'
        path = r'C:\Users\test\Desktop\Testcase2jira Sypnase\Intretech-VV-0202 绘本台灯SF1-自研_趣学伴_Android V2.5.0_测试用例_20210929.xlsx'
        # path = r'C:\Users\test\Desktop\Testcase2jira Sypnase\test.xlsx'
        # path = r'C:\Users\test\Desktop\Testcase2jira Sypnase\Intretech-VV-0202 Clean Handle-BAT-MCU-HC89F0411P-Image-V3.6_测试报告_2021.11.5.xlsx'
        uploadpath = 'F:\PycharmProjects\TestCase2SypnaseRT\document\excel.xlsx'
        wb = app.books.open(path)
        uploadwb = app.books.open(uploadpath)
        self.wb = wb
        self.uploadwb = uploadwb

    # 完成
    def finish_analysis(self): #完成
        # self.wb.save()  #解析excel阶段对excel的操作都不保存
        self.uploadwb.save('upload.xlsx')
        self.wb.close()
        self.uploadwb.close()
        # 退出excel程序，
        self.app.quit()
        # 通过杀掉进程强制Excel app退出
        self.app.kill()

    #把测试报告里面所有的测试用例sheet的有效范围列出来。返回字典{‘sheetname：（行，列）’}
    def testcase_sheet_area(self): #完成  #TODO 这边应该把表格里的
        # 获取总sheet数量
        sheet_count = self.wb.sheets.count
        # 测试用例部分的sheet合集，及每个sheet页面里的有效行列数
        sheet_shape = {}
        #获取测试用例sheet页数合集
        testcase_sheet = []
        #获取每个sheet页面的有效列数
        sheet_shape_col = {}


        ###获取测试报告中的所有测试用例sheet总数###
        for num in range(4,sheet_count+1):
            #从第4个sheet开始找页面中是否有操作步骤，预期结果，如果有就是测试用例，如果没有可能是“目录”，如咕咕机；后面几页也是一样的判断方式  TODO：雷蛇的用例怎么处理
            sheet = self.wb.sheets(num)
            sheet_name = sheet.name
            for i in range(ord('A'), ord('O')):   #ord('A')= 65 ord('O')=79  chr(65)='A'
                cellvalue = sheet.range(chr(i) + str(1)).value  # 这边需要用到正则表达式，不然可能找不到。因为大家编辑excel的时候可能将空格带入
                # print("单元格位置",chr(i) + str(1),cellvalue)

                match_premise = re.findall("前置条件",str(cellvalue))
                match_step = re.findall("操作步骤",str(sheet.range(chr(i+1) + str(1)).value))
                match_expected = re.findall("预期输出", str(sheet.range(chr(i + 2) + str(1)).value))
                #页面中要有“前提”“操作”“预期”才能算是测试用例
                if len(match_premise) != 0 or len(match_step) != 0 or len(match_expected) != 0:
                    testcase_sheet.append(sheet.name)
                    # print('测试用例总页数：', len(testcase_sheet), '', testcase_sheet)
                    # sheet_name = sheet.name
                    sheet_shape_col[sheet_name] = chr(i+2)
                    # print(sheet_shape_col)
                    break


        ###获取每个sheet页面里的有效行列值###
        for name in testcase_sheet:
            # print(name)
            sheet = self.wb.sheets(name)
            # 读取行列
            info = sheet.used_range
            nrows = info.last_cell.row
            ncols = info.last_cell.column

            #####获取有效行，列数######
            valid_row = nrows
            # 获取A列不为空的行数，从获得的行数倒数筛查整行是否为空----为了获得最大行数的准确性，需要增加获取“前提”“操作”“预期”3列的列号，然后加上这3列对应的每行值的判断。#TODO 放在后续优化
            while (sheet.range('A' + str(valid_row)).expand('right').value) == None:
                #如果用A列向右获取整行，如果第二格为空，就会停止。所以要加“前提”“操作步骤”“预期”的判断，这几行为空才是真的整行为空
                # print(premise_col,step_col,expected_col)
                valid_row -= 1
            # print('有效行数为：',valid_row)

            #添加“测试人员误操作，出现尾部离最后一行用例的任意一行中整行就第一格有内容，后面无内容的情况”的过滤   TODO：这边可能需要根据实际情况再优化一下
            premise_col = ord(sheet_shape_col[name]) -2
            step_col = ord(sheet_shape_col[name]) - 1
            expected_col = sheet_shape_col[name]
            # print("premise_col=",chr(premise_col),"step_col=",chr(step_col),"expected_col",expected_col)
            premise_value = sheet.range(chr(premise_col) + str(valid_row)).value
            step_value = sheet.range(chr(step_col) + str(valid_row)).value
            expected_value = sheet.range(expected_col + str(valid_row)).value
            # print("前提：",premise_value,"操作：",step_value,"预期",expected_value)
            if premise_value == None and step_value == None and expected_value == None:
                valid_row -= 1
                #过滤完在循环判断下，从下往上的行数是否都有内容
                while (sheet.range('A' + str(valid_row)).expand('right').value) == None:
                    valid_row -= 1
                # print('扣除测试人员误操作引起的多余行数，有效行数为：',valid_row)

            # 判断目前获取到的A列最后一行是否有合并单元格（如果是合并单元格，合并的部分除了提一条，后面都是空值，有效行数就要加上这部分合并的行数）
            combine = sheet.range('A' + str(valid_row)).merge_cells
            #combine返回True就是有合并
            if combine:
                combine_area = sheet.range('A' + str(valid_row)).merge_area
                # print("当前sheet实际有效总行数：", combine_area.last_cell.row)
                valid_row = combine_area.last_cell.row
            else:
                # print("当前sheet实际有效总行数：", valid_row)
                pass

            sheet_shape[name] = (valid_row, sheet_shape_col[name])
            # print(sheet_shape)

            # 做一个判断，如果得到的非空白行总数>实际有用行总数的10，就加一条提示语，如果表格底下有很多空白无用行，提示大家删除
            if nrows - valid_row > 10 :
                print(name,"sheet尾部存在无效行，请删除")
        # print(sheet_shape)
        #返回 测试用例有效sheet集合，每个有效sheet对应的有效范围,每个sheet页面列头所对应的行值（用于和新表格内容对齐）
        return testcase_sheet,sheet_shape

    #根据每个sheet范围，检查有效范围内是否所有单元格都有填写内容。有填写不完整的情况，提示测试人员补充更新
    #返回不符合项的位置和不符合项内容
    #完成
    def testcase_form_check(self): #完成
        sheet_shape = self.testcase_sheet_area()
        error_message = ''
        for name in sheet_shape[0]:
            sheet = self.wb.sheets(name)
            sheet_area_col = sheet_shape[1][name][1]
            sheet_area_row = sheet_shape[1][name][0]
            # print("当前sheet名称:",name,"列数：",sheet_area_col,)
            #以列为基准进行过滤，检查是否有未填写项
            for c in range(ord("A"),ord(sheet_area_col)+1):
                # print('当前循环的是',chr(c))
                # for r in range(1,sheet_area_row+1):
                r = 1
                while (r < sheet_area_row+1):
                    # print('r = ',r)
                    cellvalue = sheet.range(chr(c)+str(r)).value
                    # print('cellvalue = ',cellvalue)
                    #判断单元格是否有值
                    if cellvalue == None:
                        cellvalue_combine = sheet.range(chr(c)+str(r)).merge_cells
                        # print('cellvalue_combine=',cellvalue_combine)
                        # 如果返回True就是有合并单元格

                        if cellvalue_combine:
                            # cellvalue_combine_area = sheet.range(chr(c)+str(r)).merge_area
                            last_combine_len = sheet.range(chr(c) + str(r - 1)).merge_area.last_cell.row
                            cellvalue_combine_len = sheet.range(chr(c)+str(r)).merge_area.last_cell.row
                            ####判断合并单元格的首格为空
                            # 获取为空的单元格的上一行的合并单元格长度，如果长度和本格的合并长度相等，那说明是同一个合并部分；如果上一行的合并长度≠本格的合并长度，说明本格是本格以下合并的首格，需要有内容，没有内容就说明本格以下合并的部分内容为空
                            if last_combine_len < cellvalue_combine_len:
                                # print("合并单元格首格，【", name, "】的", chr(c) + str(r), "为空")
                                error_message += "【" + name + "】的" + chr(c) + str(r) + "为空" + '\n'
                            else:
                                pass
                            r = cellvalue_combine_len+1
                            # print('加了合并单元格的r = ',r)
                        else:
                            # print("【",name,"】的",chr(c)+str(r),"为空")
                            error_message += "【" + name + "】的" + chr(c)+str(r) + "为空" + '\n'
                            r += 1
                    else:
                        # print("过滤的单元格位置",chr(c)+str(r))
                        # print(chr(c)+str(r),'的值为',cellvalue)
                        r += 1



        return error_message



    ###获取每个sheet页面列头所在的列######
    ######获取每个sheet页面列头内容对应位置，返回字典，格式如{“听听”:{'模块':'A','功能点':'B'},"我的":{'模块':'A','功能点':'C'}}
    #完成
    def testcase_column_head(self): #完成
        sheet_shape = self.testcase_sheet_area()
        # 获取每个sheet页面列头所在的列
        column_head = {}  # 总的，由每个sheet合起来的
        sheet_column_head = {}  # 每个sheet里的列头信息
        for sheet_name in sheet_shape[0]:
            sheet = self.wb.sheets(sheet_name)
            for i in range(ord('A'), ord('O')):
                cellvalue = sheet.range(chr(i) + str(1)).value  # 这边需要用到正则表达式，不然可能找不到。因为大家编辑excel的时候可能将空格带入
                # print("单元格位置",chr(i) + str(1),cellvalue)

                match_module = re.findall("模块", str(cellvalue))
                if len(match_module) != 0 :
                    sheet_column_head["模块"] = chr(i)

                match_function = re.findall("功能点", str(cellvalue))
                if len(match_function) != 0 :
                    sheet_column_head["功能点"] = chr(i)

                match_testcase_version = re.findall("测试用例版本号", str(cellvalue))
                if len(match_testcase_version) != 0 :
                    sheet_column_head["测试用例版本号"] = chr(i)

                match_testcase_id = re.findall("用例编号", str(cellvalue))
                if len(match_testcase_id) != 0 :
                    sheet_column_head["用例编号"] = chr(i)

                match_testcase_name = re.findall("用例名称", str(cellvalue))
                if len(match_testcase_name) != 0 :
                    sheet_column_head["用例名称"] = chr(i)

                match_importance_level = re.findall("重要级别", str(cellvalue))
                if len(match_importance_level) != 0 :
                    sheet_column_head["重要级别"] = chr(i)

                match_premise = re.findall("前置条件",str(cellvalue))
                if len(match_premise) != 0 :
                    sheet_column_head["前置条件"] = chr(i)

                match_step = re.findall("操作步骤",str(cellvalue))

                if len(match_step) != 0 :
                    sheet_column_head["操作步骤"] = chr(i)

                match_expected = re.findall("预期输出", str(cellvalue))
                if len(match_expected) != 0 :
                    sheet_column_head["预期输出"] = chr(i)

            column_head[sheet_name] = sheet_column_head
        print(column_head)
        return column_head



    #要上传的内容放新生成的表格里
    #根据测试用例版本号过滤出要上传的新用例,主要分为3个步骤：
    # 1.过滤新用例到新表格里
    # 2.新表格用seleium创建到jira的同时将生成的用例jira写到新表格里
    # 3.根据新表格里的新用例jiraid更新《测试报告》中的用例编号
    #《测试报告》中完成状态为“open”的用例在jira上可以设置为“锁定”或者“不适用”
    #todo  完成，但是需要优化，代码写得有带乱！！！
    def testcase_to_upload(self,testcase_version='V2.5.0'): #未传入参数则默认上传所有版本的用例
        # upload_path = r'F:\PycharmProjects\TestCase2SypnaseRT\document\uplod excel模板.xlsx'
        # app = xw.App(visible=False, add_book=False)
        # upload_excel = app.books.open(upload_path)
        update_excel_sheet = self.uploadwb.sheets(1)
        report_excel = self.testcase_sheet_area()
        report_column_head = self.testcase_column_head()
        report_excel_sheetname = report_excel[0]
        newtestcase_row_dirt = {} #用来放新测试用例的行号，格式为{'sheetname':[2,4,5,6,11,19]}
        new_excel_row_dirt= {} #弄一个dirt = {‘sheetname’：{L1:76,L2:78,L3:245}},然后新值写入时，新表格对应的行就是‘L1’中的第1行,用字符切片s[1:]取
        new_excel_row_site = {}
        for sheetname in report_excel_sheetname:
            print('当前表名',sheetname)
            report_excel_sheet = self.wb.sheets(sheetname)
            newtestcase_row_list = []
            new_excel_row_site = {}
            #各个要素在新表格中对应的列号
            update_excel_sheet_col_dirt = {"模块":'B',"功能点":'C',"测试用例版本号":'D',"用例编号":'E',"用例名称":'F',"重要级别":'G',"前置条件":'H',"操作步骤":'I',"预期输出":'J'}

            #各个要素在《测试报告》中对应的列号
            module_value_col = report_column_head[sheetname]['模块']
            function_value_col = report_column_head[sheetname]['功能点']
            testcase_version_col = report_column_head[sheetname]['测试用例版本号']
            testcase_id_col = report_column_head[sheetname]['用例编号']
            testcase_name_col = report_column_head[sheetname]['用例名称']
            importance_level_col = report_column_head[sheetname]['重要级别']
            premise_col = report_column_head[sheetname]['前置条件']
            step_col = report_column_head[sheetname]['操作步骤']
            expected_col = report_column_head[sheetname]['预期输出']
            #根据传参的用例版本号，过滤要上传的新用例
            report_excel_sheetname_row = report_excel[1][sheetname][0]
            for row in range(2,report_excel_sheetname_row):
                testcase_version_value = report_excel_sheet.range(testcase_version_col+str(row)).value
                # print('testcase_version_value=',testcase_version_value)
                match_testcase_version = re.findall(str(testcase_version),str(testcase_version_value))
                if len(match_testcase_version) != 0: #新用例
                    print('新用例')
                    #######将《测试报告》中的新用例行号写到新表格里#######
                    # 获取新表格目前的行数
                    update_excel_sheet_row = update_excel_sheet.used_range.last_cell.row
                    update_excel_sheet.range('L' + str(update_excel_sheet_row + 1)).value = row
                    update_excel_sheet.range('A' + str(update_excel_sheet_row + 1)).value = sheetname
                    newtestcase_row_list.append(row)
                    new_excel_row_site['L' + str(update_excel_sheet_row + 1)] = row
                    print(sheetname,'新用例的行号',row)
                    self.uploadwb.save('upload.xlsx')

                    #################################################

                    # for i in range(ord('A'), ord('O')):  # ord('A')= 65 ord('O')=79  chr(65)='A'
                    #     #遍历《测试报告》中需要导入的每条用例内容
                    #     cellvalue = report_excel_sheet.range(chr(i) + str(row)).value
                    #     print('cellvalue=',cellvalue)
                    #     #判断是否有值
                    #     #单元格没有值
                    #     if cellvalue == None:
                    #         combine = report_excel_sheet.range(chr(i) + str(row)).merge_cells
                    #         #判断是否有合并单元格
                    #         #单元格合并
                    #         if combine:
                    #             print(chr(i) + str(row),"合并",combine)
                    #             combine_row = row
                    #             # TODO 这部分需要优化，太慢了，可以将《测试用例》中新用例对应的行在新表格中写一列，然后根据这个行号，以列获取
                    #             #比如：新用例在《测试报告》
                    #             while (report_excel_sheet.range(chr(i) + str(combine_row)).value == None):
                    #                 combine_row -= 1
                    #                 print(chr(i) + str(combine_row),'向上获取值')
                    #             combine_cellvalue = report_excel_sheet.range(chr(i) + str(combine_row)).value
                    #             print('combine_cellvalue=',combine_cellvalue)
                    #             #TODO 进行写入新表格
                    #             #获取新表格目前的行数
                    #             update_excel_sheet_row = update_excel_sheet.used_range.last_cell.row
                    #             # 判断《测试报告》sheet页面里的内容对应在新表格的哪列，用于下面的值填写在新表格正确的位置上
                    #             #写入模块的值
                    #             if module_value_col == chr(i):
                    #                 print('新旧表格列标头相同')
                    #                 report_column = report_column_head[sheetname]
                    #                 for k, v in report_column.items():
                    #                     if v == module_value_col:
                    #                         print('update_excel_sheet_col_dirt[k]=',update_excel_sheet_col_dirt[k])
                    #                         update_excel_sheet_col = update_excel_sheet_col_dirt[k]
                    #                         print('update_excel_sheet_col=',update_excel_sheet_col)
                    #                         update_excel_sheet.range(update_excel_sheet_col+str(update_excel_sheet_row+1)).value = combine_cellvalue
                    #                         upload_excel.save(r'F:\PycharmProjects\TestCase2SypnaseRT\document\upload.xlsx')
                    #
                    #         #单元格没有合并---一般是不会出现这个的，因为会先检查格式再进行转换，格式没有通过检查，不会到转换
                    #         else:
                    #             print(sheetname,"位置：",chr(i) + str(row),"没有内容。")
                    #     #单元格有值，TODO 直接将整行对应值写到新表格里
                    #     else:
                    #         print('单元格有值')
                    #         pass

                #不导入的用例
                else:
                    # print(sheetname,"第",str(row),"行的用例不导入")
                    pass

            newtestcase_row_dirt[sheetname]=newtestcase_row_list  #取出来的值为{‘sheetname’:[2,4,6,7,9]}
            new_excel_row_dirt[sheetname]=new_excel_row_site  #取出来的值{'sheetname':{'L1':23,'L2':24}
            print('newtestcase_row_dirt=',newtestcase_row_dirt)
            print('new_excel_row_dirt=',new_excel_row_dirt)

        #再遍历一遍《测试报告》表名，找各个单元格内容
        for sheetname in report_excel_sheetname:
            report_excel_sheet = self.wb.sheets(sheetname)
            newtestcase_row_list = newtestcase_row_dirt[sheetname]  #举例：内容为[2,3,4,6,8,10]类似的格式
            report_excel_sheet_maxrow = report_excel[1][sheetname][0]
            report_column = report_column_head[sheetname]
            upload_excel_write_dirt = {}
            report_excel_col_list = [ord(report_column_head[sheetname]['模块']),ord(report_column_head[sheetname]['功能点']),ord(report_column_head[sheetname]['测试用例版本号']),ord(report_column_head[sheetname]['用例编号']),ord(report_column_head[sheetname]['用例名称']),ord(report_column_head[sheetname]['重要级别']),ord(report_column_head[sheetname]['前置条件']),ord(report_column_head[sheetname]['操作步骤']),ord(report_column_head[sheetname]['预期输出'])]
            print(report_excel_col_list)
            # for c in range(ord('A'), ord('O')):
            for c in report_excel_col_list:
                print('现在循环的是',chr(c),'列')
                # for row in range(2,report_excel_sheet_maxrow):
                #     # combine_area = Interval(2,row)  #python区间库，有在这个区间就返回true，左右都是闭合区间，可以加参数设置为开区间
                #     #先读单元格的值
                #     cellvalue = report_excel_sheet.range(chr(i)+str(row)).value
                #     if cellvalue == None: #‘A2’“B2”等第2行的值一定有
                #         #存在合并单元格，要去取合并的单元格的第一个格子的值
                #         if row ==2:
                #             print('第2行就没有值，有问题')
                #         else:
                #
                #             pass
                #         cellvalue = ''
                #         pass
                #     else:
                #         pass
                cellvalue = report_excel_sheet.range(chr(c)+str(2)).value
                cell_site = chr(c)+str(2)
                print('cellvalue=',cellvalue)
                combine_lower = 2
                combine_upper = report_excel_sheet.range(chr(c)+str(2)).merge_area.last_cell.row
                combine_area = (combine_lower,combine_upper)
                for i in range (0,len(newtestcase_row_list)):
                    while newtestcase_row_list[i] > combine_upper:
                        print(sheetname,'newtestcase_row_list[i] > combine_upper','newtestcase_row_list[i]=',newtestcase_row_list[i],"combine_upper",combine_upper)
                        cellvalue = report_excel_sheet.range(chr(c)+str(combine_upper+1)).value
                        cell_site = chr(c)+str(combine_upper+1)
                        #对合并单元格的值做一个判断
                        if cellvalue != None:
                            pass
                        else:
                            print(sheetname,chr(c)+str(combine_upper+1),'合并的第一个格的值为空，有问题')
                        combine_lower = combine_upper + 1
                        combine_upper = report_excel_sheet.range(chr(c)+str(combine_lower)).merge_area.last_cell.row
                        print("combine_upper=",combine_upper)
                        combine_area = (combine_lower,combine_upper)
                        print("更新单元格合并区间：",chr(c),combine_area,'此时的值=',cellvalue)

                    if newtestcase_row_list[i] in Interval(combine_lower,combine_upper):
                        print('此时值在区间内')
                        #######写入upload表格的位置确认
                        # 行号
                        write_row_string = [k for k, v in new_excel_row_dirt[sheetname].items() if
                                            v == newtestcase_row_list[i]]  # TODO 取到的是数组，如['L34']，还要切片
                        write_row = write_row_string[0][1:]
                        # print('写入的行号：', write_row)
                        # 列号
                        # 各个要素在《测试报告》中对应的列号
                        update_excel_sheet_col_dirt = {"模块": 'B', "功能点": 'C', "测试用例版本号": 'D', "用例编号": 'E', "用例名称": 'F',"重要级别": 'G', "前置条件": 'H', "操作步骤": 'I', "预期输出": 'J'}
                        # module_value_col = report_column_head[sheetname]['模块']
                        # function_value_col = report_column_head[sheetname]['功能点']
                        # testcase_version_col = report_column_head[sheetname]['测试用例版本号']
                        # testcase_id_col = report_column_head[sheetname]['测试用例版本号']
                        # testcase_name_col = report_column_head[sheetname]['用例名称']
                        # importance_level_col = report_column_head[sheetname]['重要级别']
                        # premise_col = report_column_head[sheetname]['前置条件']
                        # step_col = report_column_head[sheetname]['操作步骤']
                        # expected_col = report_column_head[sheetname]['预期输出']

                        ##判断当前循环的列对应新旧表格是哪个列
                        write_col = ''
                        if chr(c) == report_column_head[sheetname]['模块']:
                            write_upload_excel_module_value_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_module_value_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_module_value_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue
                        if chr(c) == report_column_head[sheetname]['功能点']:
                            write_upload_excel_function_value_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_function_value_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_function_value_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue
                        if chr(c) == report_column_head[sheetname]['测试用例版本号']:
                            write_upload_excel_testcase_version_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_testcase_version_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_testcase_version_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue
                        if chr(c) == report_column_head[sheetname]['用例编号']:
                            write_upload_excel_testcase_id_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_testcase_id_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_testcase_id_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue
                            #把测试用例编号在《测试报告》中的位置写入upload exccel中的K列,值的格式如“E4","E5"
                            update_excel_sheet.range('K' + str(write_row)).value = cell_site
                        if chr(c) == report_column_head[sheetname]['用例名称']:
                            write_upload_excel_testcase_name_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_testcase_name_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_testcase_name_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue
                        if chr(c) == report_column_head[sheetname]['重要级别']:
                            write_upload_excel_importance_level_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_importance_level_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_importance_level_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue
                        if chr(c) == report_column_head[sheetname]['前置条件']:
                            write_upload_excel_premise_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_premise_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_premise_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue

                        if chr(c) == report_column_head[sheetname]['操作步骤']:
                            write_upload_excel_step_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_step_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_step_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue

                        if chr(c) == report_column_head[sheetname]['预期输出']:
                            write_upload_excel_expected_col = [k for k, v in report_column.items() if v == chr(c)]
                            if len(write_upload_excel_expected_col) != 0:
                                write_col = update_excel_sheet_col_dirt[write_upload_excel_expected_col[0]]
                            print(sheetname, write_col, write_row, '写入的值=', cellvalue)
                            upload_excel_write_dirt[str(write_col) + str(
                                write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                            update_excel_sheet.range(write_col + str(write_row)).value = cellvalue


                        # #判断要写入哪列
                        # write_upload_excel_module_value_col = [k for k, v in report_column.items() if v == module_value_col]
                        # write_upload_excel_function_value_col = [k for k, v in report_column.items() if v == function_value_col]
                        # write_upload_excel_testcase_version_col = [k for k, v in report_column.items() if v == testcase_version_col]
                        # write_upload_excel_testcase_id_col = [k for k, v in report_column.items() if v == testcase_id_col]
                        # write_upload_excel_testcase_name_col = [k for k, v in report_column.items() if v == testcase_name_col]
                        # write_upload_excel_importance_level_col = [k for k, v in report_column.items() if v == importance_level_col]
                        # write_upload_excel_premise_col = [k for k, v in report_column.items() if v == premise_col]
                        # write_upload_excel_step_col = [k for k, v in report_column.items() if v == step_col]
                        # write_upload_excel_expected_col = [k for k, v in report_column.items() if v == expected_col]
                        # write_col = ''
                        # if len(write_upload_excel_module_value_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_module_value_col[0]]
                        # if len(write_upload_excel_function_value_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_function_value_col[0]]
                        # if len(write_upload_excel_testcase_version_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_testcase_version_col[0]]
                        # if len(write_upload_excel_testcase_id_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_testcase_id_col[0]]
                        # if len(write_upload_excel_testcase_name_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_testcase_name_col[0]]
                        # if len(write_upload_excel_importance_level_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_importance_level_col[0]]
                        # if len(write_upload_excel_premise_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_premise_col[0]]
                        # if len(write_upload_excel_step_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_step_col[0]]
                        # if len(write_upload_excel_expected_col) != 0:
                        #     write_col = update_excel_sheet_col_dirt[write_upload_excel_expected_col[0]]

                        #得到行和列，开始写入upload excel
                        # write_col = update_excel_sheet_col_dirt[write_col_key]
                        # print("write_col=",write_col)
                        # print("写入的列号：", write_col)
                        # print(sheetname,write_col,write_row,'写入的值=',cellvalue)
                        # upload_excel_write_dirt[str(write_col) + str(write_row)] = cellvalue  # 保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                        # update_excel_sheet.range(write_col + str(write_row)).value = cellvalue
                        self.uploadwb.save('upload.xlsx')
                    else:
                        print(sheetname,'行号:',newtestcase_row_list[i],'≤merge_area.last_cell.row:',combine_upper,'但不在区间内！有问题！')


                    # #######写入upload表格的位置确认
                    # #行号
                    # write_row_string = [k for k, v in new_excel_row_dirt[sheetname].items() if v == newtestcase_row_list[i]]  #TODO 取到的是数组，如['L34']，还要切片
                    # write_row = write_row_string[0][1:]
                    # print('写入的行号：',write_row)
                    # #列号
                    # # 各个要素在《测试报告》中对应的列号
                    # module_value_col = report_column_head[sheetname]['模块']
                    # function_value_col = report_column_head[sheetname]['功能点']
                    # testcase_version_col = report_column_head[sheetname]['测试用例版本号']
                    # testcase_id_col = report_column_head[sheetname]['测试用例版本号']
                    # testcase_name_col = report_column_head[sheetname]['用例名称']
                    # importance_level_col = report_column_head[sheetname]['重要级别']
                    # premise_col = report_column_head[sheetname]['前置条件']
                    # step_col = report_column_head[sheetname]['操作步骤']
                    # expected_col = report_column_head[sheetname]['预期输出']
                    #
                    # write_col_string =[k for k, v in report_column.items() if v == module_value_col]
                    # write_col = write_col_string[0][1:]
                    # print("写入的列好：",write_col)
                    # upload_excel_write_dirt[write_col+write_row]=cellvalue  #保存新表格中，每个单元格对应的值，如{'A1':'关闭弹窗'，'A2':'下拉刷新'}
                    # update_excel_sheet.range(write_col+str(write_row)).value = cellvalue
                    # self.uploadwb.save('upload.xlsx')





    #这个函数为调试的，到时候可以删除，但是删除前需要再看下哪些是有用的信息
    def check_excelreport_form(self):
        # 显示表格sheet名称
        print(self.wb.sheets)
        #获取表格sheet数量
        print(self.wb.sheets.count)

        #引用名为“Sheet1”的表单
        # ws =wb.sheets('Sheet1')
        ws = self.wb.sheets(7)
        print(ws.name)
        # ws.activate()
        # ws = wb.sheets.active

        #读取行和列
        info = ws.used_range
        print(info)
        nrowss = info.last_cell.row
        ncols = info.last_cell.column
        print("本sheet行列数：",nrowss,ncols)
        n = ws.used_range.shape
        print("shape范围：",n)
        #TODO 这边要先找下“前提”“步骤”“预期”所在的列，同A列一起判断，是否真的整行为无效行，目前单单只通过A列判断还是会误判

        premise_col = ''
        step_col = ''
        expected_col = ''
        '''
        for i in range(ord('A'), ord('O')):
            cellvalue = ws.range(chr(i) + str(1)).value  # 这边需要用到正则表达式，不然可能找不到。因为大家编辑excel的时候可能将空格带入
            # print("单元格位置",chr(i) + str(1),cellvalue)
            match_premise = re.findall("前置条件", str(cellvalue))
            match_step = re.findall("操作步骤", str(cellvalue))
            match_expected = re.findall("预期输出", str(cellvalue))

            # 页面中要有“前提”“操作”“预期”才能算是测试用例
            if len(match_premise) != 0:
                premise_col = chr(i)
            if len(match_step) != 0:
                step_col = chr(i)
            if len(match_expected) != 0:
                expected_col = chr(i)
        
            '''


        #验证sheet最后面的行是否为空行
        # ww = ws.range('A'+str(nrows)).expand('right').value
        nrows = nrowss
        # print('A列',nrows,'的值：', ws.range('A' + str(nrows)).expand('right').value)
        # print('H列', nrows, '的值：', ws.range('H' + str(nrows)).expand('right').value)
        # print('I列', nrows, '的值：', ws.range('J' + str(nrows)).expand('right').value)

        while (ws.range('A'+str(nrows)).expand('right').value) == None : #获取A列不为空的行数，从获得的行数倒数筛查整行是否为空----为了获得最大行数的准确性，需要增加获取“前提”“操作”“预期”3列的列号，然后加上这3列对应的每行值的判断。#TODO 放在后续优化

            # print('A列最后一行的值',ws.range('A'+str(nrows)).expand('right').value,'当前行号',nrows)
                nrows -= 1
        print(nrows) #输出不为空的行总数


        #这里需要做一个异常判断，如果最后一行第一格有内容，后面都没有内容。也会被while循环误判算一行，所以要扣掉;
        #如果有必要的话，需要做多行判断的话，这边需要加一个for循环
        premise_value = ws.range('H' + str(nrows)).value
        step_value = ws.range('I' + str(nrows)).value
        expected_value = ws.range('J' + str(nrows)).value
        if premise_value != None or step_value != None or expected_value != None:
            pass
        else:
            nrows -= 1

        while (ws.range('A' + str(nrows)).expand(
                'right').value) == None:  # 获取A列不为空的行数，从获得的行数倒数筛查整行是否为空----为了获得最大行数的准确性，需要增加获取“前提”“操作”“预期”3列的列号，然后加上这3列对应的每行值的判断。#TODO 放在后续优化

            # print('A列最后一行的值',ws.range('A'+str(nrows)).expand('right').value,'当前行号',nrows)
            nrows -= 1
        print(nrows)  # 输出不为空的行总数




        hh = ws.range('A'+str(nrows)).merge_cells  #因为这边只是判断A列的，考虑到A列存在合并的情况，所有要再判断下，A列最后一行是否是合并，因为如果合并的话，合并的第二个第三个等获取到的值都是空的，所有要获取合并区域，最大行数要加上合并的区域
        print(hh)
        if hh:  #判断
            hebing_len = ws.range('A'+str(nrows)).merge_area
            print("当前sheet总行数：",hebing_len.last_cell.row)
            #做一个判断，如果得到的非空白行总数>实际有用行总数的10，就加一条提示语，如果表格底下有很多空白无用行，提示大家删除--#TODO 后续优化
        else:
            print("当前sheet总行数：",nrows)
        print('A1=',ws.range('A1').value)
        print('A4667行=', ws.range('A4667').expand('right').value)
        #TODO 获取“前提”的列号
        print(ws.range('F1').value)
        for i in range(ord('A'),ord('O')):
            # print(chr(i))
            # aa = chr(i) + str(1)
            # print(aa)
            # print(type(aa))
            # print(ws.range(aa).value)
            if ws.range(chr(i) + str(1)).value == "操作步骤":   #这边需要用到正则表达式，不然可能找不到。因为大家便捷excel的时候可能将空格带入
                print('找到了')
                print(chr(i))
        #获取“操作步骤”的列号
        #获取“预期结果”的列号


        # 以下内容先不要删除
        print(ws.range('E2920').value)
        # x = ws.range('C1').expand('down').value  #获取从C1开始，C列整列的值，注意，如果中间有空值，会停止向下获取直接停止结束，不会报错
        he = ws.range('E2920').merge_area  #找出合并单元格的区域,用he.last_cell.row获取行和列值
        print(he.last_cell.row)
        print(he.last_cell.column)
        hebing = ws.range('E2920').merge_cells  #如果返回True就是有合并单元格
        print(hebing)

        #找到合并的首个单元格的值
        row =he.last_cell.row
        while (ws.range('E' + str(row)).value == None):
            row -= 1
        print(row)





        #读取整行的值









# '''
#以下为调试信息
a = ExcelReportAnalysis()
# a.check_excelreport_form()
# print(a.testcase_sheet_area())
print(a.testcase_form_check())
# print(a.testcase_column_head())
# a.testcase_to_upload()
a.finish_analysis()
# '''

# old = {"模块":'A',"功能点":'B',"用例编号":'C'}
# new = {"模块":'A',"功能点":'C',"用例编号":'D'}
#
# for k,v in old.items():
#     if v == 'A':
#         print(k)
#
# print(str(k for k,v in old.items() if v == 'A'))#这个怎么取值？
# x=[k for k,v in old.items() if v == 'A']
# print("x:",x)
#
# def get_keys(d, value):
#     return [k for k,v in d.items() if v == value]
#
# print(get_keys({'a':'001', 'b':'002'}, '001')) # => ['a']


# list = [2,4,5,6,8,9,12,14,18,33,56]
# interval = (2,15)
# l = [x for x in list if x in interval]
# print(l)
# print(x for x in list if x in interval)  #intreval是一个区间库
# for x in list :
#     if x in interval:
#         print(x)

# s='L22334'
# print(s[1:])