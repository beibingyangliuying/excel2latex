import xlwings as xlw


class ExcelContent:
    def __init__(self, scope: xlw.Range):
        self.values = []
        self.format = []
        self.horizontal_alignments = []
        # self.row_borders = []
        # self.column_borders = []
        # TODO：需要完善关于borders的支持

        for i in range(scope.shape[0]):
            self.values.append([])
            self.format.append([])
            self.horizontal_alignments.append([])

            for j in range(scope.shape[1]):
                cell = scope[i, j]
                self.values[i].append(str.strip(cell.api.Text))
                self.format[i].append(cell.font)
                self.horizontal_alignments[i].append(cell.api.HorizontalAlignment)

    def __str__(self):
        return str(self.values)
