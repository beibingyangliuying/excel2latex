import xlwings as xlw
from Constants import style_code


class ExcelContent:
    def __init__(self, scope: xlw.Range):
        self.values = []
        self.format = []
        self.shape = tuple(scope.shape)

        for i in range(scope.shape[0]):
            self.values.append([])
            self.format.append([])

            for j in range(scope.shape[1]):
                cell = scope[i, j]
                font = cell.api.DisplayFormat.Font
                self.values[i].append(str.strip(cell.api.Text))
                self.format[i].append(
                    CellFont(xlw.utils.int_to_rgb(font.Color), font.Bold, font.Italic,
                             False if font.Underline == style_code['not_underline'] else True))

    def __str__(self):
        return str(self.values)


class CellFont:
    def __init__(self, color: tuple[int, int, int], bold: bool, italic: bool, underline: bool):
        self.color = color
        self.bold = bold
        self.italic = italic
        self.underline = underline
