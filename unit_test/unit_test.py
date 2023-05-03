import unittest


class ContentTestCase(unittest.TestCase):
    """
    测试Content类能否正确读取工作簿的数据
    """

    def test_simplest(self):
        from interpreter.excel_context import ExcelContext
        from content import ExcelContent

        with ExcelContext(r'test.xlsx', 'simplest', 'A1:D5') as context:
            content = ExcelContent(context.range)
            self.assertTrue(content.values)
            self.assertTrue(content.format)
            print(content)

    def test_simplest_with_borders(self):
        from interpreter.excel_context import ExcelContext
        from content import ExcelContent

        with ExcelContext(r'test.xlsx', 'simplest with borders', 'A1:D5') as context:
            content = ExcelContent(context.range)
            self.assertTrue(content.values)
            self.assertTrue(content.format)
            print(content)

    def test_simplest_with_format(self):
        from interpreter.excel_context import ExcelContext
        from content import ExcelContent

        with ExcelContext(r'test.xlsx', 'simplest with format', 'A1:D5') as context:
            content = ExcelContent(context.range)
            self.assertTrue(content.values)
            self.assertTrue(content.format)
            print(content)


class CellTextContextTestCase(unittest.TestCase):
    def test_text_expression(self):
        from interpreter.cell_text_context import TextExpression, CellTextContext

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
        from interpreter.cell_text_context import CellTextContext, ColorExpression
        color_expression = ColorExpression()

        context = CellTextContext((1, 0, 0), False, False, False, 'what?')
        self.assertEqual(color_expression.interpret(context), [r'\textcolor[rgb]{1,0,0}'])

        context.color = (0, 0, 0)
        self.assertEqual(color_expression.interpret(context), [])

    def test_concatenate_expression(self):
        from interpreter.cell_text_context import CellTextContext, CommandChainExpression, BoldExpression, \
            ItalicExpression, UnderlineExpression, ColorExpression

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
        from interpreter.cell_text_context import CellTextContext, CommandChainExpression, BoldExpression, \
            ItalicExpression, UnderlineExpression, ColorExpression, ParameterExpression, TextExpression

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
        from content import ExcelContent
        from interpreter.excel_context import ExcelContext
        from interpreter.cell_text_context import CellTextContext, CommandChainExpression, BoldExpression, \
            ItalicExpression, UnderlineExpression, ColorExpression, ParameterExpression, TextExpression
        single_parameter_expression = ParameterExpression(
            CommandChainExpression(ColorExpression(), BoldExpression(), ItalicExpression(), UnderlineExpression()),
            TextExpression())

        with ExcelContext(r'test.xlsx', 'simplest with format', 'A1:D5') as context:
            excel_content = ExcelContent(context.range)
            for i in range(excel_content.shape[0]):
                for j in range(excel_content.shape[1]):
                    text = excel_content.values[i][j]
                    font = excel_content.format[i][j]
                    cell_context = CellTextContext(font.color, font.bold, font.italic, font.underline, text)
                    print(single_parameter_expression.interpret(cell_context))

        self.assertTrue(True)


if __name__ == '__main__':
    unittest.main()
