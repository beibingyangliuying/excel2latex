from abc import ABC, abstractmethod
from typing import Any


class CellTextContext:
    def __init__(self, color: tuple[int, int, int], bold: bool, italic: bool, underline: bool, text: str):
        self.color = color
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.text = text


class AbstractExpression(ABC):
    @abstractmethod
    def interpret(self, context: CellTextContext) -> Any:
        pass


class CommandExpression(AbstractExpression):
    @abstractmethod
    def interpret(self, context: CellTextContext) -> list[str, ...]:
        pass


class EntityExpression(AbstractExpression):
    @abstractmethod
    def interpret(self, context: CellTextContext) -> str:
        pass


class ColorExpression(CommandExpression):
    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if context.color == (0, 0, 0):
            return []
        else:
            temp = '{' + ','.join((str(number) for number in context.color)) + '}'
            return [r'\textcolor[rgb]{}'.format(temp)]


class BoldExpression(CommandExpression):
    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if not context.bold:
            return []
        else:
            return [r'\textbf']


class ItalicExpression(CommandExpression):
    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if not context.italic:
            return []
        else:
            return [r'\textit']


class UnderlineExpression(CommandExpression):
    def interpret(self, context: CellTextContext) -> list[str, ...]:
        if not context.underline:
            return []
        else:
            return [r'\underline']


class TextExpression(EntityExpression):
    table = str.maketrans({'$': r'\$',
                           '^': r'\^',
                           '_': r'\_',
                           '\\': r'\textbackslash{}',
                           '%': r'\%',
                           '\n': r'\\'})

    def interpret(self, context: CellTextContext) -> str:
        return context.text.translate(self.table)


class ConcatenateExpression(CommandExpression):
    def __init__(self, *args: CommandExpression):
        """
        :param args: CommandExpression序列，顺序很重要！
        """
        self.expressions = args

    def interpret(self, context: CellTextContext) -> list[str, ...]:
        result = []

        for expression in self.expressions:
            temp = expression.interpret(context)
            if not temp:
                continue
            else:
                result.extend(temp)

        return result
