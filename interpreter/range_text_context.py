import xlwings as xlw

from constants import style_code
from abc import ABC, abstractmethod
from typing import Any
from decorators import singleton


class RangeTextContext:
    def __init__(self, scope: xlw.Range):
        self.shape: tuple[int, int] = tuple(scope.shape)
        self.cell_text_contexts: list[list[CellTextContext, ...], ...] = []

        for i in range(scope.shape[0]):
            self.cell_text_contexts.append([])

            for j in range(scope.shape[1]):
                cell = scope[i, j]
                font = cell.api.DisplayFormat.Font

                self.cell_text_contexts[i].append(
                    CellTextContext(xlw.utils.int_to_rgb(font.Color),
                                    font.Bold,
                                    font.Italic,
                                    False if font.Underline == style_code['not_underline'] else True,
                                    str.strip(cell.api.Text)))

    def __getitem__(self, item):
        if isinstance(item, tuple):
            i = item[0]
            j = item[1]
        elif isinstance(item, int):
            i = item // self.shape[1] - 1
            j = item - i * self.shape[1] - 1
        else:
            raise IndexError('Unsupported index: {0}'.format(item))

        return self.cell_text_contexts[i][j]

    def __str__(self):
        return str(self.cell_text_contexts)


class CellTextContext:
    def __init__(self, color: tuple[int, int, int], bold: bool, italic: bool, underline: bool, text: str):
        self.color = color
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.text = text

    def __str__(self):
        return '{color:{0},bold:{1},italic:{2},underline:{3},text:{4}}'.format(self.color, self.bold, self.italic,
                                                                               self.underline, self.text)


class AbstractExpression(ABC):
    @abstractmethod
    def interpret(self, context: CellTextContext) -> Any:
        pass


class CommandExpression(AbstractExpression):
    @abstractmethod
    def interpret(self, context: CellTextContext) -> list[str, ...]:
        pass


class LiteralExpression(AbstractExpression):
    @abstractmethod
    def interpret(self, context: CellTextContext) -> str:
        pass


@singleton
class ColorExpression(CommandExpression):
    default = (0, 0, 0)
    pattern = r'\textcolor[rgb]{0}'

    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if context.color == self.default:
            return []
        else:
            temp = '{' + ','.join((str(number) for number in context.color)) + '}'
            return [self.pattern.format(temp)]


@singleton
class BoldExpression(CommandExpression):
    default = r'\textbf'

    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if not context.bold:
            return []
        else:
            return [self.default]


@singleton
class ItalicExpression(CommandExpression):
    default = r'\textit'

    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if not context.italic:
            return []
        else:
            return [self.default]


@singleton
class UnderlineExpression(CommandExpression):
    default = r'\underline'

    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if not context.underline:
            return []
        else:
            return [self.default]


@singleton
class TextExpression(LiteralExpression):
    translate_table = str.maketrans({'$': r'\$',
                                     '^': r'\^',
                                     '_': r'\_',
                                     '\\': r'\textbackslash{}',
                                     '%': r'\%',
                                     '\n': r'\\'})

    def interpret(self, context: CellTextContext) -> str:
        result = context.text.translate(self.translate_table)
        if r'\\' in result:
            result = r'\makecell{' + result + '}'

        return result


class CommandChainExpression(CommandExpression):
    def __init__(self, *args: CommandExpression):
        """
        :param args: CommandExpression序列，顺序很重要！
        """
        self.commands = args

    def interpret(self, context: CellTextContext) -> list[str, ...]:
        result = []

        for expression in self.commands:
            temp = expression.interpret(context)
            if not temp:
                continue
            else:
                result.extend(temp)

        return result


class ParameterExpression(LiteralExpression):
    def __init__(self, command: CommandExpression, entity: LiteralExpression):
        self.command = command
        self.entity = entity

    def interpret(self, context: CellTextContext) -> str:
        commands = self.command.interpret(context)
        parameter = self.entity.interpret(context)

        if not commands:  # 说明没有命令需要执行
            return parameter

        result = []
        bracket = '{'
        for command in commands:
            result.append(command)
            result.append(bracket)
        else:
            result.append(parameter)

        result.extend(['}'] * len(commands))
        return ''.join(result)
