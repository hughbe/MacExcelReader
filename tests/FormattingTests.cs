namespace MacExcelReader.Tests;

public class FormattingTests
{
    [Theory]
    [InlineData("A1", 0, 0)]
    [InlineData("b19", 18, 1)]
    [InlineData("$C$5", 4, 2)]
    [InlineData("IV16384", 16383, 255)]
    [InlineData(" AA10 ", 9, 26)]
    public void Address_Parses(string text, int row, int column)
    {
        Assert.Equal(new ExcelCellAddress(row, column), ExcelCellAddress.Parse(text));
    }

    [Theory]
    [InlineData("")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A0")]
    [InlineData("IW1")]
    [InlineData("A16385")]
    [InlineData("ABC1")]
    [InlineData("A1B")]
    [InlineData(null)]
    public void Address_RejectsInvalid(string? text)
    {
        Assert.False(ExcelCellAddress.TryParse(text, out _));
    }

    [Fact]
    public void Address_ToStringRoundTrips()
    {
        for (int column = 0; column < 256; column += 17)
        {
            var address = new ExcelCellAddress(column * 31, column);
            Assert.Equal(address, ExcelCellAddress.Parse(address.ToString()));
        }
    }

    [Theory]
    [InlineData(1234.567, 0, "1234.567")]
    [InlineData(1234.567, 1, "1235")]
    [InlineData(1234.567, 2, "1234.57")]
    [InlineData(1234.567, 3, "1,235")]
    [InlineData(1234.567, 4, "1,234.57")]
    [InlineData(1234.567, 5, "$1,235 ")]
    [InlineData(-1234.567, 5, "($1,235)")]
    [InlineData(-1234.567, 6, "($1,234.57)")]
    [InlineData(0.256, 7, "26%")]
    [InlineData(0.256, 8, "25.60%")]
    [InlineData(1234.567, 9, "1.23E+03")]
    [InlineData(29220, 10, "1/1/84")]
    [InlineData(29220, 11, "1-Jan-84")]
    [InlineData(29220, 12, "1-Jan")]
    [InlineData(29220, 13, "Jan-84")]
    [InlineData(0.75, 14, "6:00 PM")]
    [InlineData(0.75, 16, "18:00")]
    [InlineData(29220.5, 18, "1/1/84 12:00")]
    public void Format_BuiltInFormats(double value, int format, string expected)
    {
        Assert.Equal(expected, ExcelValueFormatter.Format(value, format));
    }

    [Theory]
    [InlineData(0.5, 15, "12:00:00 PM")]
    [InlineData(0.75, 17, "18:00:00")]
    [InlineData(-1, 10, "-1/2/04")]
    [InlineData(3_000_000, 10, "3000000")]
    [InlineData(1E+30, 0, "1E+30")]
    [InlineData(1.25, 19, "1.25")]
    public void Format_HandlesAdditionalAndFallbackFormats(double value, int format, string expected)
    {
        Assert.Equal(expected, ExcelValueFormatter.Format(value, format));
    }

    // A format shows the number Excel's decimal conversion makes, at most 14 significant digits (docs/Display.md), not the
    // double's binary expansion; percent signs move its point rather than multiplying, so they cannot overflow.
    [Theory]
    [InlineData(0.30000000000000004, "0.000000000000000000", "0.300000000000000000")]
    [InlineData(1.0 / 3, "0.000000000000000000", "0.333333333333330000")]
    [InlineData(123456789012345678, "0", "123456789012350000")]
    [InlineData(123456789012345678, "0%", "12345678901235000000%")]
    [InlineData(12345678901234.5, "0.00", "12345678901234.00")]
    [InlineData(0.999, "0.00", "1.00")]
    [InlineData(9.996, "0.00E+00", "1.00E+01")]
    [InlineData(0.004, "0.00", "0.00")]
    [InlineData(0.005, "0.00", "0.01")]
    public void Format_ShowsFourteenSignificantDigits(double value, string format, string expected)
    {
        Assert.Equal(expected, ExcelNumberFormat.Parse(format).Format(value));
    }

    [Theory]
    [InlineData(1234.5678, "0E+00.0", "1E+003.0")]
    [InlineData(1234.5678, "0.00E+00", "1.23E+03")]
    [InlineData(1234.5678, "0E+0", "1E+3")]
    [InlineData(1234.5678, "0E-0", "1E3")]
    [InlineData(0.00012345, "0E+0", "1E-4")]
    [InlineData(0.00012345, "0.0E-00", "1.2E-04")]
    [InlineData(1.5E+123, "0E+0", "2E+123")]
    [InlineData(1.5E+123, "0E+00000", "2E+00123")]
    [InlineData(1234.5678, "0E+00 0", "1E+00 3")]
    [InlineData(1234.5678, "00E+00", "12E+02")]
    public void Format_FillsTheExponentsRunsUpToADecimalPoint(double value, string format, string expected)
    {
        // TEXT in Excel 1.00 under the harness: the exponent, at least as many digits as there are placeholders after
        // the E, fills the runs after it up to a '.' as an integer would; runs after the '.' show only their 0s
        // (docs/Display.md).
        Assert.Equal(expected, ExcelNumberFormat.Parse(format).Format(value));
    }

    [Theory]
    [InlineData(0, "0.0.0", "0.0.0")]
    [InlineData(0, "0.00.0", "0.00.0")]
    [InlineData(999.995, "0.000.0:", "999.995.0:")]
    [InlineData(999.995, "0.000.00E+00", "9.999.95E+02")]
    [InlineData(-0.5, "#.##0.000,0$", "-.50.000,0$")]
    [InlineData(10000000000, ".000$#.##", "10000000000.000$.")]
    [InlineData(-1E-05, "0.0% #.##", "-0.0% 0.1")]
    [InlineData(-1, "0.0\" kg\"#,##0.000", "-1.0 kg,0.000")]
    [InlineData(0.001234, "#,##0.000#.##/0.00", "0.0012.34/0.00")]
    [InlineData(29220.75, "0E-0000.0%", "3E00004.0%")]
    [InlineData(12.345, "##0.0E+0#,##0/#.##", "123.5E-0,0/1.")]
    [InlineData(-1, "", "-")]
    [InlineData(1, "h:mmh:mm $A/P", "0:0012:00 $A")]
    [InlineData(999.995, "AM/PMh:mm/h:mm:ss", "PM23:52/23:52:48")]
    public void Format_FollowsExcelsDisplayRoutineTokenByToken(double value, string format, string excel)
    {
        // TEXT in Excel 1.00 for random formats the library showed differently before it followed 19.0d: a second
        // '.' prints and the decimals go on, each decimal run pads on its own, integer digits no run took come out at
        // the point, percent signs do not scale a scientific number, an empty section still shows the sign, and only
        // the hour just before AM/PM is 12-hour (docs/Display.md).
        Assert.Equal(excel, ExcelNumberFormat.Parse(format).Format(value));
    }

    [Fact]
    public void Parse_PutsTheExponentCountWhereExcelsCompilerDoes()
    {
        // Padding a thousands format to four integer positions inserts a run before a 0 run; Excel then writes the
        // exponent's count two bytes early, on the E (spec §7).
        var format = ExcelNumberFormat.Parse("0,00E-00");
        Assert.Equal("1C040000070401060302021E000602", Convert.ToHexString(format.Body).Substring(0, 30));
    }

    [Fact]
    public void Format_PercentOfTheLargestNumberIsDigits()
    {
        string text = ExcelNumberFormat.Parse("0%").Format(double.MaxValue);
        Assert.StartsWith("17976931348623000", text);
        Assert.Equal(312, text.Length);
    }

    [Fact]
    public void FormatCell_HandlesAllBasicKinds()
    {
        Assert.Throws<ArgumentNullException>(() => ExcelValueFormatter.FormatCell(null!));
        Assert.Equal(string.Empty, ExcelValueFormatter.FormatCell(new ExcelCell()));

        var cell = new ExcelCell();
        cell.SetText("text");
        Assert.Equal("text", ExcelValueFormatter.FormatCell(cell));

        var formula = new ExcelFormula([0x1D, 1]);
        formula.SetResult(false);
        cell.SetFormula(formula);
        Assert.Equal("FALSE", ExcelValueFormatter.FormatCell(cell));
    }

    [Fact]
    public void Dates_Use1904System()
    {
        Assert.Equal(new DateTime(1904, 1, 1), ExcelValueFormatter.ToDateTime(0));
        Assert.Equal(new DateTime(1984, 1, 1), ExcelValueFormatter.ToDateTime(29220));
        Assert.Equal(29220, ExcelValueFormatter.ToSerial(new DateTime(1984, 1, 1)));
    }

    [Fact]
    public void BuiltInFormatNames()
    {
        Assert.Equal("General", ExcelNumberFormat.BuiltIn[0].ToText());
        Assert.Equal("m/d/yy h:mm", ExcelNumberFormat.BuiltIn[18].ToText());
        Assert.Equal(ExcelNumberFormat.BuiltInCount, ExcelNumberFormat.BuiltIn.Count);
        Assert.True(ExcelValueFormatter.IsDateFormat(10));
        Assert.False(ExcelValueFormatter.IsDateFormat(9));
    }

    [Fact]
    public void FormatCell_UsesCellFormatAndKinds()
    {
        var document = TestData.Load("Expenses.xls");
        Assert.Equal("1/1/84", ExcelValueFormatter.FormatCell(document.Cells["A11"]));
        Assert.Equal("$1,000 ", ExcelValueFormatter.FormatCell(document.Cells["C11"]));
        Assert.Equal("overhead", ExcelValueFormatter.FormatCell(document.Cells["B11"]));

        var formula = new ExcelFormula([0x1D, 1]);
        formula.SetResult(true);
        var cell = new ExcelCell();
        cell.SetFormula(formula);
        Assert.Equal("TRUE", ExcelValueFormatter.FormatCell(cell));
        formula.SetResult(ExcelValue.FromError(ExcelError.DivideByZero));
        Assert.Equal("#DIV/0!", ExcelValueFormatter.FormatCell(cell));
    }

    [Fact]
    public void Password_HashIsStable()
    {
        Assert.Equal(0, ExcelPassword.Hash(string.Empty));
        Assert.Equal(ExcelPassword.Hash("secret"), ExcelPassword.Hash("secret"));
        Assert.NotEqual(ExcelPassword.Hash("secret"), ExcelPassword.Hash("Secret"));

        // Worked by hand from the algorithm in the spec (§5.2): 'a' (0x61) gives 0x61, rotated to 0xC2,
        // XOR the length 1 gives 0xC3, XOR 0xCE4B gives 0xCE88.
        Assert.Equal(0xCE88, ExcelPassword.Hash("a"));
    }

    [Fact]
    public void MacRoman_DecodesSpecialCharacters()
    {
        Assert.Equal("é", MacRoman.GetString([0x8E]));
        Assert.Equal("™", MacRoman.GetString([0xAA]));
        Assert.Equal("", MacRoman.GetString([0xF0]));
        Assert.Equal(new byte[] { 0x8E, 0x41 }, MacRoman.GetBytes("éA"));
        Assert.Equal(new byte[] { (byte)'?' }, MacRoman.GetBytes("中"));
    }

    [Fact]
    public void MacRoman_RoundTripsAllBytes()
    {
        var bytes = Enumerable.Range(0, 256).Select(i => (byte)i).ToArray();
        Assert.Equal(bytes, MacRoman.GetBytes(MacRoman.GetString(bytes)));
    }

    [Fact]
    public void SupportingTypes_ExposeAndValidateState()
    {
        var external = new ExcelExternalDocument("Book");
        Assert.Equal("Book", external.ToString());
        Assert.Throws<ArgumentNullException>(() => new ExcelExternalDocument(null!));

        var inner = new InvalidOperationException("inner");
        var exception = new ExcelFormatException("message", inner);
        Assert.Equal("message", exception.Message);
        Assert.Same(inner, exception.InnerException);

        var state = new ExcelSheetState(new byte[ExcelSheetState.Size])
        {
            ReferenceStyle = ExcelReferenceStyle.A1,
            LastVisibleColumn = 255,
            RecalculationState = ExcelRecalculationState.Pending,
            StandardFont = 3,
        };
        Assert.Equal(ExcelReferenceStyle.A1, state.ReferenceStyle);
        Assert.Equal(255, state.LastVisibleColumn);
        Assert.Equal(ExcelRecalculationState.Pending, state.RecalculationState);
        Assert.Equal(3, state.StandardFont);
            Assert.Throws<ArgumentException>(() => new ExcelSheetState(new byte[ExcelSheetState.Size - 1]));

            Assert.True(ExcelFunctions.Get(53)?.IsMacroOnly ?? false);
            Assert.Throws<ArgumentOutOfRangeException>(() => FormulaDecompiler.ColumnName(-1));
        }
    
    [Theory]
    [InlineData(0, "General")]
    [InlineData(4, "#,##0.00")]
    [InlineData(5, "$#,##0 ;($#,##0)")]
    [InlineData(9, "0.00E+00")]
    [InlineData(14, "h:mm AM/PM")]
    [InlineData(18, "m/d/yy h:mm")]
    public void BuiltInFormats_DecompileToTheirNames(int index, string expected)
    {
        // Excel 1.00's FRMT 256 bodies, turned back into text as the Format Number dialog does (spec §7).
        Assert.Equal(expected, ExcelNumberFormat.BuiltIn[index].ToText());
    }

    [Theory]
    [InlineData("WEST GERMANY", "00-00-000", 1234.5, "00-01-235")]
    [InlineData("INSTRUMENT DATA", "### ##0.00", 1234.5, "1 234.50")]
    [InlineData("SURVEY", "0.0\" N\";0.0\" S\";0", -12.25, "12.3 S")]
    [InlineData("INVOICE LIST", "$#,##0.00 ;($#,##0.00);", 0, "")]
    [InlineData("Format Scaling", "#,##0,", 1234567, "1,235")]
    [InlineData("Colour Formats", "[blue]#,##0;[red]-#,##0", -1234.5, "-1,235")]
    public void UserFormats_DecompileAndFormat(string sample, string text, double value, string expected)
    {
        var format = Assert.Single(TestData.Load(sample).NumberFormats, f => f.ToText() == text);
        Assert.Equal(expected, format.Format(value));
    }

    [Fact]
    public void TextSection_FormatsText()
    {
        // Excel 1.5's @ shows text through @, and 0;-0;0;"t" replaces it with t (docs/Display.md).
        Assert.Equal("abc", Assert.Single(TestData.Load("Colour Formats").NumberFormats, f => f.ToText() == "@").Format("abc"));
        Assert.Equal("t", Assert.Single(TestData.Load("Format Text Section").NumberFormats).Format("abc"));
        Assert.Equal("abc", ExcelNumberFormat.BuiltIn[2].Format("abc"));
    }

    [Theory]
    [MemberData(nameof(TestData.AllSheets), MemberType = typeof(TestData))]
    public void EveryCell_FormatsWithItsDocumentsFormat(string sample)
    {
        var document = ExcelDocument.Read(File.ReadAllBytes(TestData.SampleFile(sample)));
        foreach (var entry in document.Cells.NonEmptyCells)
        {
            var cell = entry.GetCell();
            Assert.NotNull(ExcelValueFormatter.FormatCell(cell, document));
        }

        // An empty format is one Excel stores (FORMAT.NUMBER("") in `TEXT Formats`), so only check there is text.
        Assert.All(document.NumberFormats, f => Assert.NotNull(f.ToText()));
    }

    [Theory]
    [InlineData(1234567, 5, "1E+06")]          // too many integer digits: one mantissa digit up to 6 characters
    [InlineData(1234567, 9, "1234567")]
    [InlineData(123.456, 5, "123.5")]          // rounded to the digits that fit, less one for the point
    [InlineData(0.5, 3, "0.5")]
    [InlineData(0.000012345, 8, "1.23E-05")]   // a tiny number that does not fit: 8 - 6 + 1 mantissa digits
    [InlineData(0.000012345, 12, "0.000012345")]
    [InlineData(-2.5, 3, "-2.5")]              // the minus sign is not counted
    public void General_FitsTheColumn(double value, int width, string expected)
    {
        // Excel 1.00's General (19.0c): docs/Display.md.
        Assert.Equal(expected, ExcelValueFormatter.FormatGeneral(value, width));
    }

    [Fact]
    public void DateFormats_AreDecidedAsExcelDoes()
    {
        // 07.18 counts day, month and year tokens only; 34.04 adds the built-in formats 10–18 (docs/Display.md).
        Assert.True(ExcelNumberFormat.BuiltIn[10].IsDate);    // m/d/yy
        Assert.False(ExcelNumberFormat.BuiltIn[16].IsDate);   // h:mm
        Assert.True(ExcelValueFormatter.IsDateFormat(16));
        Assert.False(ExcelValueFormatter.IsDateFormat(9));
        var document = TestData.Load("INVOICE TEMPLATE");
        int dayMonthYear = document.NumberFormats.FindIndex(f => f.ToText() == "dd mmm yy") + ExcelNumberFormat.BuiltInCount;
        Assert.True(ExcelValueFormatter.IsDateFormat(dayMonthYear, document));
        Assert.False(ExcelValueFormatter.IsDateFormat(dayMonthYear));
    }

    [Fact]
    public void Formats_FitTheirColumn()
    {
        // Excel 1.00's display (19.0d, 19.0e; docs/Display.md), counting the column in digit widths.
        var general = ExcelNumberFormat.BuiltIn[0];
        Assert.Equal("1.235E+12", general.Format(1234567891234, 9));      // General switches to an exponent
        Assert.Equal("-123", general.Format(-123.456, 5));                 // the minus sign takes a character
        Assert.Equal("######", ExcelNumberFormat.BuiltIn[2].Format(1234567.891, 6));   // 0.00 too wide
        Assert.Equal("######", ExcelNumberFormat.BuiltIn[10].Format(100000, 6));       // a date Excel cannot show
        Assert.Equal(string.Empty, general.Format(1, 0));                  // a hidden column

        var fill = Assert.Single(TestData.Load("Fill Format").NumberFormats, f => f.ToText() == "*-0");
        Assert.Equal("12", fill.Format(12));
        Assert.Equal("-------12", fill.Format(12, 10));                    // fills what is left, less a pixel
    }

    [Fact]
    public void FormatCell_UsesTheColumnWidth()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["A1"].SetNumber(123456789);
        document.State.SetColumnWidth(0, 40);                              // room for (40 - 4) / 7 = 5 digits
        Assert.Equal("1E+08", ExcelValueFormatter.FormatCell(document.Cells["A1"], document, 0));
        Assert.Equal("123456789", ExcelValueFormatter.FormatCell(document.Cells["A1"], document));
        document.Cells["A1"].NumberFormat = 3;                             // #,##0
        Assert.Equal("####", ExcelValueFormatter.FormatCell(document.Cells["A1"], document, 0)); // (40 - 4) / 9
    }

    [Theory]
    [InlineData(20, 1.5)]
    [InlineData(30, 0.125)]
    [InlineData(100, 2)]
    public void Format_MoreThan15Decimals_Formats(int decimals, double value)
    {
        var format = ExcelNumberFormat.Parse("0." + new string('0', decimals));
        string text = format.Format(value);
        Assert.Equal(decimals, text.Length - text.IndexOf('.') - 1);
        Assert.Equal(value, double.Parse(text, System.Globalization.CultureInfo.InvariantCulture));
        _ = ExcelNumberFormat.Parse("0." + new string('0', decimals) + "E+00").Format(value);
    }

    [Fact]
    public void Format_TruncatedToken_IsAFormatError()
    {
        // A number header (03) needs three operand bytes; this section ends after one.
        var format = new ExcelNumberFormat([0, 0, 0, 2, 2, 2, 1], [0x03, 0x01]);
        Assert.Throws<ExcelFormatException>(() => format.Format(1));
    }

    [Fact]
    public void Format_SameFormatTwice_GivesTheSameText()
    {
        // The tokens are cached after the first call.
        var format = ExcelNumberFormat.Parse("#,##0.00;(#,##0.00);\"zero\"");
        Assert.Equal("1,234.50", format.Format(1234.5));
        Assert.Equal("1,234.50", format.Format(1234.5));
        Assert.Equal("(2.00)", format.Format(-2));
        Assert.Equal("zero", format.Format(0));
        var other = format with { Body = ExcelNumberFormat.Parse("0").Body, Descriptor = ExcelNumberFormat.Parse("0").Descriptor };
        Assert.Equal("1235", other.Format(1234.5));
    }
}
