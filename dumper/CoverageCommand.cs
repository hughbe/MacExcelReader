using System.ComponentModel;
using System.Globalization;
using System.Text;
using Spectre.Console;
using Spectre.Console.Cli;
using MacExcelReader;

/// <summary>
/// Writes macro sheets that exercise every part of the formula encoding, for checking in Excel itself.
/// </summary>
sealed class CoverageSettings : CommandSettings
{
    [CommandArgument(0, "<output>")]
    [Description("The macro sheet to write (its data fork). A MacBinary copy and a token listing are written next to it.")]
    public required string Output { get; init; }

    [CommandOption("--experimental")]
    [Description("Also write <output>.experimental: hand-built tokens the parser cannot produce (intersection, array constants). These are described from code alone and could upset Excel.")]
    public bool Experimental { get; init; }

    [CommandOption("--excel <VERSION>")]
    [Description("Write the sheet for Excel 1.0 (1.00–1.04, the default) or 1.5, which adds functions and encodes a few things differently.")]
    public string Excel { get; init; } = "1.0";

    [CommandOption("--split <COUNT>")]
    [Description("Write several smaller sheets, \"<output> 01\", \"<output> 02\"…, of at most COUNT formulas each (Excel on a small Mac runs out of memory on the whole sheet).")]
    public int Split { get; init; }
}

sealed class CoverageCommand : Command<CoverageSettings>
{
    // Argument candidates, each with the value types it satisfies.
    private static readonly (string Text, ExcelValueTypes Types)[] s_candidates =
    [
        ("A1", ExcelValueTypes.Reference | ExcelValueTypes.Number | ExcelValueTypes.Text | ExcelValueTypes.Logical),
        ("2", ExcelValueTypes.Number),
        ("\"x\"", ExcelValueTypes.Text),
        ("TRUE", ExcelValueTypes.Logical),
        ("$A$1:$B$2", ExcelValueTypes.Reference | ExcelValueTypes.Array),
        ("Total", ExcelValueTypes.Reference | ExcelValueTypes.Number),
        ("0.5", ExcelValueTypes.Number),
        ("#N/A", ExcelValueTypes.Error),
        ("B$3", ExcelValueTypes.Reference | ExcelValueTypes.Number | ExcelValueTypes.Text),
        ("FALSE", ExcelValueTypes.Logical),
        ("\"\"", ExcelValueTypes.Text),
        ("$C4", ExcelValueTypes.Reference | ExcelValueTypes.Number),
    ];

    public override int Execute(CommandContext context, CoverageSettings settings, CancellationToken cancellationToken)
    {
        var version = settings.Excel == "1.5" ? ExcelVersion.Excel150 : ExcelVersion.Excel100;
        var document = CreateDocument();
        var listing = new StringBuilder();
        int row = 0;
        int formulas = 0;
        int partFormulas = 0;
        string heading = string.Empty;
        var parts = new List<(ExcelDocument Document, StringBuilder Listing, int Formulas)>();
        var skipped = new List<string>();

        void Header()
        {
            document.Cells[0, 0].SetText("Intended formula (text)");
            document.Cells[0, 1].SetText("Formula as Excel shows it");
            document.Cells[0, 0].Bold = true;
            document.Cells[0, 1].Bold = true;
            row = 1;
        }

        void WriteHeading()
        {
            row++;
            document.Cells[row, 0].SetText(heading);
            document.Cells[row, 0].Bold = true;
            listing.AppendLine().AppendLine($"## {heading}");
            row++;
        }

        // Starts a new sheet when splitting and the current one is full.
        void MakeRoom(int needed)
        {
            if (settings.Split <= 0 || partFormulas == 0 || partFormulas + needed <= settings.Split)
            {
                return;
            }

            parts.Add((document, listing, partFormulas));
            document = CreateDocument();
            listing = new StringBuilder();
            partFormulas = 0;
            Header();
            WriteHeading();
        }

        void Heading(string text)
        {
            heading = text;
            if (settings.Split > 0 && partFormulas > 0 && partFormulas >= settings.Split)
            {
                MakeRoom(int.MaxValue);
                return;
            }

            WriteHeading();
        }

        void Add(string text)
        {
            MakeRoom(1);
            ExcelFormula formula;
            try
            {
                formula = ExcelFormula.Parse(text, document, version);
            }
            catch (ExcelFormulaSyntaxException exception)
            {
                skipped.Add($"{text}: {exception.Message}");
                return;
            }

            formula.SetResult(0);
            document.Cells[row, 0].SetText(text);
            document.Cells[row, 1].SetFormula(formula);
            listing.AppendLine($"A{row + 1,-5} {text}");
            listing.AppendLine($"       {Convert.ToHexString(formula.Expression)}" + (formula.Fixups.IsEmpty ? "" : $"  fix-ups {Convert.ToHexString(formula.Fixups)}"));
            row++;
            formulas++;
            partFormulas++;
        }

        Header();

        Heading("Cell references");
        foreach (var text in new[]
        {
            "=A1", "=$A$1", "=$A1", "=A$1", "=B2", "=IV1", "=A16384", "=IV16384", "=$IV$16384", "=Z100",
        })
        {
            Add(text);
        }

        Heading("Areas");
        foreach (var text in new[]
        {
            "=A1:B2", "=$A$1:$B$2", "=$A1:B$2", "=A$1:$B2", "=A1:A1", "=A1:IV16384", "=$A$1:$IV$16384",
            "=SUM(A1:B2)", "=SUM($A$1:$A$5,C1:C3)", "=AREAS((A1,B2:C3))", "=AREAS((A1:B2,C3,$D$4))",
        })
        {
            Add(text);
        }

        Heading("Ranges built from function results");
        foreach (var text in new[]
        {
            "=INDEX(A1:C3,1,1):C3", "=A1:INDEX(A1:C3,2,2)", "=INDEX(A1:C3,1,1):INDEX(A1:C3,3,3)",
            "=SUM(OFFSET(A1,1,1):B5)", "=ROWS(A1:OFFSET(A1,2,0))",
        })
        {
            Add(text);
        }

        Heading("Names, macros and external references");
        foreach (var text in new[]
        {
            "=Total", "=Area", "=SUM(Area)", "=Total*2", "=Macro1()", "=Macro1(1)", "=Macro1(1,\"a\",A1)",
            "=Budget!A1", "=Budget!$A$1:$B$2", "=Budget!Rate", "=SUM(Budget!A1:A3)", "=Budget!A1+Budget!Rate",
            "='My Budget'!A1", "='My Budget'!Q1", "=!A1", "=!A1:B2", "=!Database", "=ROWS(!Database)",
        })
        {
            Add(text);
        }

        Heading("Number constants");
        foreach (var text in new[]
        {
            "=0", "=1", "=255", "=256", "=65535", "=65536", "=123456789", "=0.5", "=1.25", "=-1", "=-0.5",
            "=-65535", "=1E+100", "=1.5E-10", "=-1E+300", "=-(-1)",
        })
        {
            Add(text);
        }

        Heading("Text, logical and error constants");
        foreach (var text in new[]
        {
            "=\"\"", "=\"text\"", "=\"say \"\"hi\"\"\"", "=\"a,b:c\"", "=TRUE", "=FALSE",
            "=#NULL!", "=#DIV/0!", "=#VALUE!", "=#REF!", "=#NAME?", "=#NUM!", "=#N/A",
        })
        {
            Add(text);
        }

        Heading("Operators");
        foreach (var text in new[]
        {
            "=1+2", "=1-2", "=2*3", "=6/3", "=2^3", "=\"a\"&\"b\"", "=1<2", "=1<=2", "=1=2", "=1>=2", "=1>2", "=1<>2",
            "=A1+B1", "=A1-B1", "=A1*B1", "=A1/B1", "=A1^B1", "=A1&B1", "=A1<B1", "=A1<=B1", "=A1=B1", "=A1>=B1", "=A1>B1", "=A1<>B1",
            "=+A1", "=-A1", "=+1", "=A1%", "=50%", "=-A1%", "=(A1)", "=((A1))", "=(1+2)*3", "=1+2*3", "=2^-1",
            "=-2^2", "=1+2*3^4&\"x\"=\"y\"", "=(1<2)=(2<3)", "=A1:B2", "=\"a\"&1+2",
        })
        {
            Add(text);
        }

        Heading("IF and CHOOSE");
        foreach (var text in new[]
        {
            "=IF(A1,1)", "=IF(A1,1,2)", "=IF(A1>0,SUM(A1:A3),-1)", "=IF(A1,IF(B1,1,2),IF(C1,3,4))", "=IF(A1,,2)",
            "=IF(A1,\"yes\",\"no\")", "=IF(A1,B1:B3,C1:C3)", "=IF(IF(A1,B1,C1),1,2)",
            "=CHOOSE(A1,1)", "=CHOOSE(A1,1,2)", "=CHOOSE(A1,1,2,3)", "=CHOOSE(2,\"a\",B1,C1+1)",
            "=CHOOSE(A1,IF(B1,1,2),3)", "=IF(A1,CHOOSE(B1,1,2),3)",
            "=CHOOSE(A1,1,2,3,4,5,6,7,8,9,10,11,12,13)",
        })
        {
            Add(text);
        }

        Heading("SUM and volatile functions");
        foreach (var text in new[]
        {
            "=SUM(A1:A3)", "=SUM(1)", "=SUM(A1)", "=SUM(A1,A2)", "=SUM(1,2,3)", "=1+SUM(A1:A3)",
            "=RAND()", "=NOW()", "=1+RAND()", "=IF(A1,NOW())", "=INDEX(A1:C3,2,2)", "=ROWS(A1:B2)",
            "=COLUMNS(A1:B2)", "=AREAS(A1)",
        })
        {
            Add(text);
        }

        Heading("Array constants and calls to another sheet");
        foreach (var text in new[]
        {
            "=SUM({1,2,3})", "=SUM({1,2;3,4})", "=INDEX({\"a\",\"b\";\"c\",\"d\"},2,1)", "=Table",
            "='My Budget'!Q1+1", "=Budget!Rate*2",
        })
        {
            Add(text);
        }

        Heading("Array formulas (each covers two cells: this row and the next)");
        foreach (var text in new[] { "{=A1:A2*2}", "{=TRANSPOSE(A1:B1)}", "{=SUM(A1:A2*B1:B2)}" })
        {
            MakeRoom(2);
            partFormulas += 2;
            document.Cells[row, 0].SetText(text);
            document.SetArrayFormula(new ExcelRange((ushort)row, (ushort)(row + 1), 1, 1), text, version);
            listing.AppendLine($"A{row + 1,-5} {text}");
            row += 2;
            formulas += 2;
        }

        Heading("Missing arguments");
        foreach (var text in new[]
        {
            "=OFFSET(A1,,1)", "=OFFSET(A1,1,)", "=MID(\"abc\",2,1)", "=INDEX(A1:C3,,2)",
        })
        {
            Add(text);
        }

        // Every function with its fewest arguments, one more, its most arguments, and its dialog-box form.
        Heading("Every function");
        for (int number = 0; number <= ExcelFunctions.MaxNumber; number++)
        {
            var function = ExcelFunctions.Get(number, version);
            // Functions 176, 203 and 204 are named "7", "8" and "9" in Excel's list: placeholders that cannot be typed.
            if (function is null || char.IsDigit(function.Name[0]))
            {
                continue;
            }

            var counts = new SortedSet<int> { function.MinArguments, Math.Min(function.MinArguments + 1, function.MaxArguments), function.MaxArguments };
            foreach (int count in counts)
            {
                Add(Call(function.Name, function, count, number));
            }

            if (function.HasDialogForm)
            {
                Add(Call(function.Name + "?", function, function.MinArguments, number));
                if (function.MaxArguments > function.MinArguments)
                {
                    Add(Call(function.Name + "?", function, function.MaxArguments, number));
                }
            }
        }

        parts.Add((document, listing, partFormulas));
        if (settings.Split <= 0)
        {
            WriteAndCheck(settings.Output, parts[0].Document, parts[0].Listing.ToString());
        }
        else
        {
            for (int i = 0; i < parts.Count; i++)
            {
                WriteAndCheck($"{settings.Output} {i + 1:D2}", parts[i].Document, parts[i].Listing.ToString());
            }
        }

        AnsiConsole.MarkupLine($"{formulas} formulas in {parts.Count} sheet(s).");
        foreach (var skip in skipped)
        {
            AnsiConsole.MarkupLine($"[yellow]Skipped[/] {Markup.Escape(skip)}");
        }

        if (settings.Experimental)
        {
            var experimental = CreateExperimental(out string experimentalListing);
            string path = settings.Output + ".experimental";
            WriteOutputs(path, experimental.ToArray(), "XLPG", experimentalListing);
            AnsiConsole.MarkupLine($"Wrote [bold]{Markup.Escape(path)}[/] with hand-built tokens.");
        }

        return 0;
    }

    // Writes a sheet and makes sure every formula reads back to its intended text.
    private static void WriteAndCheck(string path, ExcelDocument document, string listing)
    {
        var bytes = document.ToArray();
        WriteOutputs(path, bytes, "XLPG", listing);
        var reread = ExcelDocument.Read(bytes);
        int checkedFormulas = 0;
        int mismatches = 0;
        for (int r = 0; r < reread.Cells.RowCount; r++)
        {
            if (reread.Cells[r, 1].Formula is { } formula && reread.Cells[r, 0].Type == ExcelCellType.Text)
            {
                checkedFormulas++;
                string expected = reread.Cells[r, 0].TextValue;
                string actual = formula.ToText(reread);
                if (!Equivalent(expected, actual))
                {
                    mismatches++;
                    AnsiConsole.MarkupLine($"[yellow]{Markup.Escape(Path.GetFileName(path))} row {r + 1}[/]: {Markup.Escape(expected)} decompiles as {Markup.Escape(actual)}");
                }
            }
        }

        AnsiConsole.MarkupLine($"Wrote [bold]{Markup.Escape(path)}[/]: {reread.Cells.RowCount} rows, {bytes.Length:N0} bytes" + (mismatches == 0 ? ", every formula reads back as intended." : $", [yellow]{mismatches} of {checkedFormulas} differ[/]."));
    }

    // The decompiler writes numbers in their shortest form (1E+100, 1.5E-10) and folds --1 to 1.
    private static bool Equivalent(string expected, string actual)
        => string.Equals(expected.Replace("E+", "E"), actual.Replace("E+", "E"), StringComparison.OrdinalIgnoreCase);

    private static string Call(string name, ExcelFunction function, int count, int seed)
    {
        var arguments = new string[count];
        for (int i = 0; i < count; i++)
        {
            var accepts = function.GetArgumentType(i);
            var fitting = s_candidates.Where(c => (c.Types & accepts) != 0).ToArray();
            if (fitting.Length == 0)
            {
                fitting = [s_candidates[1]];
            }

            arguments[i] = fitting[(seed + i) % fitting.Length].Text;
        }

        return "=" + name + "(" + string.Join(",", arguments) + ")";
    }

    private static ExcelDocument CreateDocument()
    {
        var document = ExcelDocument.CreateMacroSheet();
        document.State.SetColumnWidth(0, 220);
        document.State.SetColumnWidth(1, 220);

        // External documents: a named one with two names, one whose name needs quotes, and the active sheet.
        document.ExternalDocuments.Add(new ExcelExternalDocument("Budget") { Names = { "Rate", "Total" } });
        document.ExternalDocuments.Add(new ExcelExternalDocument("My Budget") { Names = { "Q1" } });
        document.ExternalDocuments.Add(new ExcelExternalDocument("\0") { Names = { "Database" } });

        // Defined names, in Excel's case-insensitive alphabetical order (tokens refer to them by position).
        // Macro1 is a function macro in column D; Command1 a command macro with a shortcut key.
        var definitions = new (string Name, string Definition, byte Flags, char Key)[]
        {
            ("Area", "=$A$1:$B$3", 0, '\0'),
            ("Check", "=$E$1", ExcelDefinedName.MacroFlag | ExcelDefinedName.CommandFlag, 'k'),
            ("Command1", "=$D$5", ExcelDefinedName.MacroFlag | ExcelDefinedName.CommandFlag, 'j'),
            ("Gone", "=#REF!", 0, '\0'),
            ("Macro1", "=$D$1", ExcelDefinedName.MacroFlag, '\0'),
            ("Table", "={1,2;3,4}", 0, '\0'),
            ("Total", "=$B$2", 0, '\0'),
            ("TwoAreas", "=$A$1:$A$3", 0, '\0'),
        };

        // Command macros get bit 0x40 in the shortcut flags byte, as Excel's own do; without it Excel leaves them out
        // of Macro > Run. Names must be in case-insensitive alphabetical order, since tokens refer to them by position.
        foreach (var (name, _, flags, key) in definitions)
        {
            document.Names.Add(new ExcelDefinedName(name, [], flags)
            {
                ShortcutKey = (byte)key,
                ShortcutExtra = (flags & ExcelDefinedName.CommandFlag) != 0 ? (byte)0x40 : (byte)0,
            });
        }

        document.Names.Sort((a, b) => string.CompareOrdinal(a.Name.ToUpperInvariant(), b.Name.ToUpperInvariant()));
        foreach (var (name, definition, _, _) in definitions)
        {
            document.Names.Single(n => n.Name == name).SetDefinition(definition, document);
        }

        // The macros the names point at.
        SetFormula(document, 0, 3, "=ARGUMENT(\"x\")");
        SetFormula(document, 1, 3, "=RETURN(D1)");
        SetFormula(document, 4, 3, "=BEEP()");
        SetFormula(document, 5, 3, "=RETURN()");

        // The Check command macro (Option-Command-K, or Macro > Run): compares Excel's own text for every formula in
        // column B (GET.FORMULA) with the intended text in column A, counts mismatches in F1, lists their rows in
        // column G and puts Excel's text for each in column C.
        string[] check =
        [
            "=SET.VALUE(F1,0)",
            "=SET.VALUE(F2,2)",
            "=SET.VALUE(F2,F2+1)",
            "=IF(F2>250,GOTO(E12))",
            "=IF(INDEX(A1:A250,F2,1)=\"\",GOTO(E3))",
            "=IF(GET.FORMULA(INDEX(B1:B250,F2,1))=\"\",GOTO(E3))",
            "=IF(GET.FORMULA(INDEX(B1:B250,F2,1))=INDEX(A1:A250,F2,1),GOTO(E3))",
            "=SET.VALUE(F1,F1+1)",
            "=SET.VALUE(INDEX(G1:G250,F1,1),F2)",
            "=SET.VALUE(INDEX(C1:C250,F2,1),GET.FORMULA(INDEX(B1:B250,F2,1)))",
            "=GOTO(E3)",
            "=ALERT(F1&\" mismatches; rows in column G, Excel's text in column C\",2)",
            "=RETURN()",
        ];
        for (int i = 0; i < check.Length; i++)
        {
            SetFormula(document, i, 4, check[i]);
        }

        // Give the check's working cells (F1, F2 and the list in column G) values now, so the grid already covers
        // them and the macro never has to grow it.
        document.Cells[0, 5].SetValue(0);
        document.Cells[1, 5].SetValue(0);
        document.Cells[0, 6].SetText("Mismatched rows");
        return document;
    }

    private static void SetFormula(ExcelDocument document, int row, int column, string text)
    {
        var formula = ExcelFormula.Parse(text, document);
        formula.SetResult(0);
        document.Cells[row, column].SetFormula(formula);
    }

    // Tokens the parser cannot produce, built from the spec (§9.1, §8.5). No sample contains them.
    private static ExcelDocument CreateExperimental(out string listing)
    {
        var document = ExcelDocument.CreateMacroSheet();
        document.State.SetColumnWidth(0, 220);
        document.State.SetColumnWidth(1, 220);
        var text = new StringBuilder();
        int row = 0;

        void Add(string label, byte[] expression, byte[]? fixups = null)
        {
            var formula = new ExcelFormula(expression, fixups ?? []);
            formula.SetResult(0);
            document.Cells[row, 0].SetText(label);
            document.Cells[row, 1].SetFormula(formula);
            text.AppendLine($"A{row + 1,-5} {label}").AppendLine($"       {Convert.ToHexString(expression)}" + (fixups is null ? "" : $"  fix-ups {Convert.ToHexString(fixups)}"));
            row++;
        }

        // Intersection (0x0F): A1:B2 B1:C3, as SUM's argument and alone.
        byte[] a1b2 = [0x25, 0xC0, 0x00, 0xC0, 0x01, 0x00, 0x01];
        byte[] b1c3 = [0x25, 0xC0, 0x00, 0xC0, 0x02, 0x01, 0x02];
        Add("=SUM(A1:B2 B1:C3) — intersection token 0F", [.. a1b2, .. b1c3, 0x0F, 0x42, 0x01, 0x04]);
        Add("=A1:B2 B1:C3 — intersection token 0F", [.. a1b2, .. b1c3, 0x0F]);

        // Array constant (0x20): {1,2;3,4}. Element type 1 = number, 2 = text (BIFF's codes; (assumed)).
        var block = new List<byte> { 2, 0, 2 };
        foreach (double value in new[] { 1.0, 2.0, 3.0, 4.0 })
        {
            block.Add(1);
            var bytes = new byte[8];
            System.Buffers.Binary.BinaryPrimitives.WriteDoubleBigEndian(bytes, value);
            block.AddRange(bytes);
        }

        Add("={1,2;3,4} — array constant token 20, number elements", [0x60, 0, 0, 0, 0, 0, 0], [.. block]);
        Add("=SUM({1,2;3,4}) — array constant as an argument", [0x60, 0, 0, 0, 0, 0, 0, 0x42, 0x01, 0x04], [.. block]);

        var textBlock = new List<byte> { 2, 0, 1 };
        textBlock.Add(2);
        textBlock.Add(1);
        textBlock.Add((byte)'a');
        textBlock.Add(1);
        var one = new byte[8];
        System.Buffers.Binary.BinaryPrimitives.WriteDoubleBigEndian(one, 1);
        textBlock.AddRange(one);
        Add("={\"a\",1} — array constant with a text element (type 2)", [0x60, 0, 0, 0, 0, 0, 0], [.. textBlock]);

        listing = text.ToString();
        return document;
    }

    private static void WriteOutputs(string path, byte[] data, string type, string listing)
    {
        File.WriteAllBytes(path, data);
        File.WriteAllBytes(path + ".bin", MacBinary(Path.GetFileName(path), type, "XCEL", data));
        File.WriteAllText(path + ".tokens.txt", listing);
    }

    /// <summary>
    /// Wraps a data fork in a MacBinary II file, so emulators and transfer tools keep the Finder type and creator.
    /// </summary>
    private static byte[] MacBinary(string name, string type, string creator, byte[] data)
    {
        var nameBytes = Encoding.ASCII.GetBytes(name.Length > 63 ? name[..63] : name);
        var header = new byte[128];
        header[1] = (byte)nameBytes.Length;
        nameBytes.CopyTo(header, 2);
        Encoding.ASCII.GetBytes(type).CopyTo(header, 65);
        Encoding.ASCII.GetBytes(creator).CopyTo(header, 69);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(83), data.Length);
        uint now = (uint)(DateTime.UtcNow - new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(91), now);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(95), now);
        header[122] = 129;
        header[123] = 129;
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(124), Crc16(header.AsSpan(0, 124)));

        var result = new byte[128 + (data.Length + 127) / 128 * 128];
        header.CopyTo(result, 0);
        data.CopyTo(result, 128);
        return result;
    }

    // CRC-16/XMODEM, as MacBinary II uses.
    private static ushort Crc16(ReadOnlySpan<byte> bytes)
    {
        ushort crc = 0;
        foreach (byte b in bytes)
        {
            crc ^= (ushort)(b << 8);
            for (int i = 0; i < 8; i++)
            {
                crc = (crc & 0x8000) != 0 ? (ushort)((crc << 1) ^ 0x1021) : (ushort)(crc << 1);
            }
        }

        return crc;
    }
}
