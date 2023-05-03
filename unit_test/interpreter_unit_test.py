import unittest


class CellTextContextTestCase(unittest.TestCase):
    def test_text_expression(self):
        from interpreter.range_text_context import TextExpression
        from interpreter.range_text_context import CellTextContext

        context = CellTextContext((0, 0, 0), False, False, False, 'abcabcabc')
        text_expression = TextExpression()
        self.assertEqual(text_expression.interpret(context), 'abcabcabc')

        context.text = 'abc$$abc'
        self.assertEqual(text_expression.interpret(context), r'abc\$\$abc', '转义$')

        context.text = 'abc^^abc'
        self.assertEqual(text_expression.interpret(context), r'abc\^\^abc', '转义^')

        context.text = 'abc__abc'
        self.assertEqual(text_expression.interpret(context), r'abc\_\_abc', '转义_')

        context.text = 'abc%%abc'
        self.assertEqual(text_expression.interpret(context), r'abc\%\%abc', '转义%')

        context.text = 'abc\nabc'
        self.assertEqual(text_expression.interpret(context), r'\makecell{abc\\abc}')

        context.text = r'abc\\abc'
        self.assertEqual(text_expression.interpret(context), r'abc\textbackslash{}\textbackslash{}abc', '转义\\')

        context.text = '$^_\\%\n'
        self.assertEqual(text_expression.interpret(context), r'\makecell{\$\^\_\textbackslash{}\%\\}', '综合转义')

    def test_color_expression(self):
        from interpreter.range_text_context import ColorExpression
        from interpreter.range_text_context import CellTextContext
        color_expression = ColorExpression()

        context = CellTextContext((1, 0, 0), False, False, False, 'what?')
        self.assertEqual(color_expression.interpret(context), [r'\textcolor[rgb]{1,0,0}'])

        context.color = (0, 0, 0)
        self.assertEqual(color_expression.interpret(context), [])

    def test_concatenate_expression(self):
        from interpreter.range_text_context import CommandChainExpression
        from interpreter.range_text_context import UnderlineExpression
        from interpreter.range_text_context import ItalicExpression
        from interpreter.range_text_context import BoldExpression
        from interpreter.range_text_context import ColorExpression
        from interpreter.range_text_context import CellTextContext

        context = CellTextContext((1, 0, 0,), True, True, True, 'what?')
        color_expression = ColorExpression()
        bold_expression = BoldExpression()
        italic_expression = ItalicExpression()
        underline_expression = UnderlineExpression()

        concatenate_expression = CommandChainExpression(color_expression, bold_expression, italic_expression,
                                                        underline_expression)
        self.assertEqual(concatenate_expression.interpret(context),
                         [r'\textcolor[rgb]{1,0,0}', r'\textbf', r'\textit', r'\underline'])

        context.bold = False
        self.assertEqual(concatenate_expression.interpret(context),
                         [r'\textcolor[rgb]{1,0,0}', r'\textit', r'\underline'])

        context.italic = False
        self.assertEqual(concatenate_expression.interpret(context), [r'\textcolor[rgb]{1,0,0}', r'\underline'])

        context.underline = False
        self.assertEqual(concatenate_expression.interpret(context), [r'\textcolor[rgb]{1,0,0}'])

        context.color = (0, 0, 0)
        self.assertEqual(concatenate_expression.interpret(context), [])

    def test_single_parameter_expression(self):
        from interpreter.range_text_context import ParameterExpression
        from interpreter.range_text_context import CommandChainExpression
        from interpreter.range_text_context import TextExpression
        from interpreter.range_text_context import UnderlineExpression
        from interpreter.range_text_context import ItalicExpression
        from interpreter.range_text_context import BoldExpression
        from interpreter.range_text_context import ColorExpression
        from interpreter.range_text_context import CellTextContext

        context = CellTextContext((1, 0, 0,), True, True, True, 'what?')
        concatenate_expression = CommandChainExpression(ColorExpression(), BoldExpression(), ItalicExpression(),
                                                        UnderlineExpression())
        text_expression = TextExpression()
        single_parameter_expression = ParameterExpression(concatenate_expression, text_expression)

        self.assertEqual(single_parameter_expression.interpret(context),
                         r'\textcolor[rgb]{1,0,0}{\textbf{\textit{\underline{what?}}}}')

        context.color = (0, 0, 0)
        self.assertEqual(single_parameter_expression.interpret(context), r'\textbf{\textit{\underline{what?}}}')

        context.italic = False
        self.assertEqual(single_parameter_expression.interpret(context), r'\textbf{\underline{what?}}')

        context.text = 'abc\nabc'
        self.assertEqual(single_parameter_expression.interpret(context), r'\textbf{\underline{\makecell{abc\\abc}}}')

    def test_interpret_cell_data(self):
        from interpreter.range_text_context import RangeTextContext
        from interpreter.backend import ExcelReader
        from interpreter.range_text_context import ParameterExpression
        from interpreter.range_text_context import CommandChainExpression
        from interpreter.range_text_context import TextExpression
        from interpreter.range_text_context import UnderlineExpression
        from interpreter.range_text_context import ItalicExpression
        from interpreter.range_text_context import BoldExpression
        from interpreter.range_text_context import ColorExpression

        single_parameter_expression = ParameterExpression(
            CommandChainExpression(ColorExpression(), BoldExpression(), ItalicExpression(), UnderlineExpression()),
            TextExpression())

        with ExcelReader(r'test.xlsx', 'simplest with format', 'A1:D5') as reader:
            range_context = RangeTextContext(reader.range)
            for i in range(range_context.shape[0]):
                for j in range(range_context.shape[1]):
                    cell_context = range_context[i, j]
                    print(single_parameter_expression.interpret(cell_context))

        self.assertTrue(True)


if __name__ == '__main__':
    unittest.main()
