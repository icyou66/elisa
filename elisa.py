import time
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Font, Alignment


class Elisa:
    def __init__(self, excel, pprint):
        self.excel = excel
        self.pprint = pprint

        # 条件格式
        fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')  # 浅红色填充
        font = Font(color='9c0006')  # 深红色文本
        self.cell_rule = CellIsRule(operator='greaterThan', formula=['0.2'], fill=fill, font=font)

    def start(self):
        for sheet_name in self.excel.sheetnames:
            self.pprint(f"已打开工作表：{sheet_name}", color="purple")
            time.sleep(1)

            # sheet即为工作表对象
            sheet = self.excel[sheet_name]
            self.handle_elisa(sheet)

            # 添加版权信息
            sheet.merge_cells("A33:A40")
            sheet['A33'] = "感谢使用ELISA数据处理小工具\n心里默默感谢一下王桂迪大神喔~\n"
            sheet['A33'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            self.pprint(f"{sheet_name}处理数据完毕！\n", color="green")

    def handle_elisa(self, sheet):
        """
        对工作表制作elisa数据单
        :param sheet: 当前工作表
        :return:
        """
        self.pprint("处理ELISA数据中...")
        # 计算阴对和阳对的平均值
        sheet['C33'] = '=AVERAGE(C25:C26)'
        sheet['C34'] = '=AVERAGE(C27:C28)'

        # 对每个单元格添加ELISA公式
        for col in range(4, 15):  # D列到N列（4到14）
            for row in range(33, 41):  # 第33行到第34行
                # 每个单元格中的公式对于的元数据规律为：列-8
                # ELISA公式：(x-阴)/(阳-阴)
                this = sheet.cell(row=row, column=col)
                this.value = f"=({chr(64 + col)}{row - 8}-C33)/(C34-C33)"

                # 添加条件格式
                sheet.conditional_formatting.add(this.coordinate, self.cell_rule)

        # 还有C37~C40单元格需要填充数据
        for row in range(37, 41):
            this = sheet.cell(row=row, column=3)
            this.value = f"=(C{row - 8}-C33)/(C34-C33)"

            # 添加条件公式
            sheet.conditional_formatting.add(this.coordinate, self.cell_rule)

        self.pprint("正在添加公式和条件格式...")
        self.pprint("正在添加条件格式...")
        time.sleep(1)
