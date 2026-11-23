using System.Buffers.Binary;
using System.ComponentModel;
using Spectre.Console;
using Spectre.Console.Cli;
using MacExcelReader;

// Fills every field the spec says Excel ignores, rebuilds or leaves over from memory with junk, to test that
// claim in a real Excel: a file that opens, draws and re-saves the same with junk in all of them shows Open does
// not depend on any. Categories can be chosen to find the field at fault when one does not.

sealed class JunkSettings : CommandSettings
{
    [CommandArgument(0, "<input>")]
    [Description("A worksheet or macro sheet (data fork).")]
    public required string Input { get; init; }

    [CommandArgument(1, "<output>")]
    [Description("Where to write the junk-filled copy.")]
    public required string Output { get; init; }

    [CommandOption("--only <CATEGORIES>")]
    [Description("Comma-separated categories to fill (default: all): " + JunkFiller.CategoryList)]
    public string? Only { get; init; }

    [CommandOption("--value <BYTE>")]
    [Description("The junk byte, in hex (default A5).")]
    public string Value { get; init; } = "A5";
}

sealed class JunkCommand : Command<JunkSettings>
{
    public override int Execute(CommandContext context, JunkSettings settings, CancellationToken cancellationToken)
    {
        // Start from the library's own clean save, so every leftover holds a known value before it is filled.
        byte[] data = ExcelDocument.Load(settings.Input).ToArray();
        var categories = settings.Only is null ? JunkFiller.All : settings.Only.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        var filler = new JunkFiller(data, Convert.ToByte(settings.Value, 16), categories);
        filler.Fill();
        File.WriteAllBytes(settings.Output, data);
        foreach (var (category, count) in filler.Counts.OrderBy(c => c.Key))
        {
            AnsiConsole.MarkupLine($"  {category}: {count} byte(s)");
        }

        // The library must still read it, and see the same document.
        var junk = ExcelDocument.Read(data);
        AnsiConsole.MarkupLine($"Wrote {Program.Escape(settings.Output)} ({junk.Cells.CountNonEmpty()} cells)");
        return 0;
    }
}

sealed class JunkFiller(byte[] data, byte junk, IReadOnlyCollection<string> categories)
{
    public const string CategoryList = "header, segcount, segments, sections, state, cache, counts, links, rows, cells, formulas, formula-scratch, formula-text, recalc-link, margin-marks, state-live, print-cache, tokens, shared, names, windows, print";

    // rows (the row index, §8) is a size hint Excel allocates from, so junk there runs it out of memory; it is left
    // out by default.
    // segcount (header word 12) and sections (the section table) are left out by default: the library itself relies
    // on them, so each is tested in Excel on its own.
    public static readonly string[] All = [.. CategoryList.Split(", ").Where(c => c is not ("segcount" or "sections" or "recalc-link" or "margin-marks" or "state-live" or "print-cache" or "rows") && !c.StartsWith("formula-", StringComparison.Ordinal))];

    public Dictionary<string, int> Counts { get; } = [];

    public void Fill()
    {
        var layout = ExcelFileLayout.Read(data);
        int state = layout.StateOffset;

        // Header words 6–11 are never read; word 12, the segment count, and the segment table are skipped (§4).
        Junk("header", 6, 6);
        Junk("segcount", 12, 2);

        // Each segment's interpreter stamp and the 10 bytes of the next heap block's header (§4.1).
        foreach (var (offset, length) in layout.CodeSegments)
        {
            if (length >= 12)
            {
                Junk("segments", offset + 2, 2);
                Junk("segments", offset + length - 10, 10);
            }
        }

        // Open reads the sections in order and skips the section table (§5).
        Junk("sections", layout.SectionTableOffset, 32);

        // State block (§6): the unused word and the pointers; the cached font metrics and print-title values,
        // which Excel works out again; and the counts Save recomputes.
        // Open loads the block into its globals, then 10.00 resets 0x026, 0x02A, 0x02E, 0x032, 0x036, 0x03A, 0x03E,
        // 0x25A, 0x25E, 0x256 and 0x266. 0x028, 0x030 and 0x03C survive from the file and are used ("state-live").
        Junk("state", state + 0x008, 2);
        Junk("state", state + 0x026, 2);
        Junk("state-live", state + 0x028, 2);
        Junk("state", state + 0x02A, 4 + 2);
        Junk("state-live", state + 0x030, 2);
        Junk("state", state + 0x032, 10);
        Junk("state-live", state + 0x03C, 2);
        Junk("state", state + 0x256, 4);
        Junk("state", state + 0x25A, 8);
        Junk("state", state + 0x266, 4);
        // Font metrics, recomputed on Open (10.00 calls 09.1c); print-title values, derived from the names but only
        // when a name or column width changes, so Open keeps the file's (10.0b, which clears them, is New only).
        Junk("cache", state + 0x248, 0x256 - 0x248);
        Junk("print-cache", state + 0x262, 4);
        Junk("print-cache", state + 0x26A, 0x274 - 0x26A);
        Junk("counts", state + 0x024, 2);
        Junk("counts", state + 0x03E, 2);

        int rows = Word(state + 0x014) + Word(state + 0x020);
        int columns = Word(state + 0x274) + Word(state + 0x020);

        // Section 1: each external document's 6 unused bytes, then the row index of size hints (§8).
        int p = layout.GetSectionOffset(ExcelSection.ExternalDocuments);
        if (Word(p) != 0)
        {
            int count = Word(p + 2);
            p += 4;
            for (int i = 0; i < count; i++)
            {
                Junk("links", p, 6);
                p += 6;
                p += 1 + data[p];
            }
        }
        else
        {
            p += 2;
        }

        for (int r = 0; r < rows; r++)
        {
            // A small number: Open makes an empty table of that size for each row above 2 (§8).
            if (Has("rows"))
            {
                BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(p + 2 * r), 3);
                Count("rows", 2);
            }
        }

        FillCells(layout, rows, columns);
        FillNames(layout);
        FillWindows(layout);
        FillPrint(layout);
    }

    private void FillCells(ExcelFileLayout layout, int rows, int columns)
    {
        int layoutState = layout.StateOffset;
        int table = layout.GetSectionOffset(ExcelSection.CellTable);
        int cellData = layout.GetSectionOffset(ExcelSection.CellData);
        // Open clears word 2's transient marks (bits 12–6) in rows 0 to the last row − 1 only (46.06), so the margin
        // row keeps whatever the file has: it is filled only when asked for.
        int marginRow = Word(layoutState + 0x020) != 0 ? rows - 1 : -1;
        for (int i = 0; i < rows * columns; i++)
        {
            int entry = table + 6 * i;
            ushort word0 = Word(entry);
            ushort word1 = Word(entry + 2);
            int type = (word1 >> 9) & 7;

            // Word 2 bits 12–6 are transient marks Open clears (§9).
            if (i / columns == marginRow ? Has("margin-marks") : Has("cells"))
            {
                BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(entry + 4), (ushort)(Word(entry + 4) | (0x1FC0 & (junk * 0x41))));
                Count(i / columns == marginRow ? "margin-marks" : "cells", 1);
            }

            switch (type)
            {
                case 0:
                    // An empty cell's word 0 is a pointer, and it has no data offset.
                    Junk("cells", entry, 2);
                    SetHighNibble(entry + 2);
                    break;
                case 1:
                case 4:
                    SetHighNibble(entry + 2);
                    break;
                case 5:
                case 6:
                    FillFormula(cellData + (word0 << 4) + (word1 >> 12));
                    break;
            }
        }
    }

    // Bits 15–12 of word 1, meaningless for cells without data.
    private void SetHighNibble(int at)
    {
        if (Has("cells"))
        {
            data[at] = (byte)((data[at] & 0x0F) | (junk & 0xF0));
            Count("cells", 1);
        }
    }

    // A formula record (§12.3): bytes 8–10, a text result's pointer, the expression's stale operands and the
    // shared record's leftovers.
    private void FillFormula(int record)
    {
        // Bytes 8–10 are not junk: the formula's link in the list of formulas waiting to be recalculated, which Open
        // rebuilds from them (spec §12.3). Filling them marks every formula as waiting.
        Junk("recalc-link", record + 8, 2);
        Junk("formula-scratch", record + 10, 1);
        bool text = data[record] == 0xFF && data[record + 1] == 0xFF && data[record + 2] == 0;
        if (text)
        {
            Junk("formula-text", record + 4, 4);
        }

        int length = data[record + 14];
        int expression = record + 15;
        FillTokens(expression, length, definition: false);
        int fixups = expression + length + (text ? 1 + data[expression + length] : 0);
        byte first = data[expression];
        if ((first & 0x9F) is 0x01 or 0x02 && Word(fixups) == 0x4000)
        {
            // Anchor: marker, 0, then the shared header (array) or table record. Bytes 6–8 of an array's header are
            // left over; a table's 3-byte link at 6 must stay 0, so only its pointer at 12–14 is filled.
            int shared = fixups + 3;
            if ((first & 0x1F) == 0x01)
            {
                Junk("shared", shared + 6, 3);
                FillTokens(shared + 10, data[shared + 9], definition: false);
            }
            else if (Has("shared"))
            {
                Junk("shared", shared + 12, 2);
                data[shared + 14] = (byte)((data[shared + 14] & 0xC0) | (junk & 0x3F));
                Count("shared", 1);
            }
        }
    }

    // The operand bytes FormulaTokens.ClearStaleBytes clears (§15.2).
    private void FillTokens(int at, int length, bool definition)
    {
        if (!Has("tokens"))
        {
            return;
        }

        var expression = data.AsSpan(at, length);
        foreach (var token in FormulaTokens.Enumerate(expression))
        {
            int o = at + token.Offset + 1;
            int operands = token.Length - 1;
            switch (token.Code)
            {
                case 0x23:
                    Junk("tokens", o + (definition ? 2 : 0), operands - (definition ? 2 : 0));
                    break;
                case 0x1A:
                    // In a definition bytes 5–6 are 00 00 and byte 7 the document index: fill only the first four.
                    Junk("tokens", o, definition ? 4 : 7);
                    break;
                case 0x20:
                    Junk("tokens", o + 3, 3);
                    break;
                case 0x1B:
                case 0x2A:
                    Junk("tokens", o, operands);
                    break;
                case 0x26:
                    Junk("tokens", o, 3);
                    break;
                case 0x27:
                    Junk("tokens", o + 2, 1);
                    break;
                case 0x01 or 0x02 when token.Offset == 0:
                    Junk("tokens", o, operands);
                    break;
                case 0x19 when data[o] is 0x01 or 0x10:
                    Junk("tokens", o + 1, 1);
                    break;
            }
        }
    }

    // Section 4 (§11): the flag bits Open discards, a macro kind's low nibble, and the definitions' stale operands.
    private void FillNames(ExcelFileLayout layout)
    {
        int p = layout.GetSectionOffset(ExcelSection.DefinedNames);
        while (data[p] != 0)
        {
            int nameLength = data[p];
            int definitionLength = data[p + 1];
            byte flags = data[p + 2];
            if (Has("names"))
            {
                data[p + 2] = (byte)(flags | 0x83);
                Count("names", 1);
            }

            int definition = p + 3;
            if ((flags & 0x40) != 0)
            {
                if (Has("names"))
                {
                    data[p + 4] = (byte)((data[p + 4] & 0xF0) | (junk & 0x0F));
                    Count("names", 1);
                }

                definition += 2;
            }

            FillTokens(definition, definitionLength, definition: true);
            p = definition + definitionLength + 1 + nameLength;
        }
    }

    // Section 6 (§13): each window's pointers and padding.
    private void FillWindows(ExcelFileLayout layout)
    {
        int p = layout.GetSectionOffset(ExcelSection.Windows);
        int stateSize = Word(2) >= 18 ? 129 : 41;
        while (data[p] == 0xFF)
        {
            int window = p + 9;
            Junk("windows", window + 23, 3);
            Junk("windows", window + 29, 3);
            Junk("windows", window + 32, 1);
            if (stateSize == 129)
            {
                Junk("windows", window + 122, 1);
            }

            p = window + stateSize;
            p += 2 + 6 * Word(p);
        }
    }

    // Section 7 (§14): the page-setup record's three string pointers.
    private void FillPrint(ExcelFileLayout layout)
    {
        int p = layout.GetSectionOffset(ExcelSection.PrintSettings);
        while (data[p] != 0)
        {
            p += 1 + data[p];
        }

        p++;
        while (data[p] != 0)
        {
            p += 1 + 2 * data[p];
        }

        p++;
        if (data[p] == 0xFF)
        {
            int record = p + 1;
            Junk("print", record + 94, 3);
            Junk("print", record + 98, 3);
            Junk("print", record + 118, 3);
        }
    }

    // "formulas" stands for its three parts.
    private bool Has(string category)
        => categories.Contains(category) || (category.StartsWith("formula-", StringComparison.Ordinal) && categories.Contains("formulas"));

    // recalc-link (formula bytes 8–9) is a real field, so it is never filled unless asked for.

    private void Junk(string category, int at, int length)
    {
        if (!Has(category) || length <= 0)
        {
            return;
        }

        data.AsSpan(at, length).Fill(junk);
        Count(category, length);
    }

    private void Count(string category, int bytes) => Counts[category] = Counts.GetValueOrDefault(category) + bytes;

    private ushort Word(int at) => BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(at));
}
