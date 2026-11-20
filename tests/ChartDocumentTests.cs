using MacExcelReader.Charts;

namespace MacExcelReader.Tests;

public class ChartDocumentTests
{
    [Theory]
    [InlineData(1.0, "4110000000000000")]
    [InlineData(0.0, "0000000000000000")]
    [InlineData(-2.5, "C125000000000000")]
    [InlineData(123.456, "4312345600000000")]
    public void MicrosoftDecimal_EncodesKnownValues(double value, string hex)
    {
        var encoded = MicrosoftDecimal.FromDouble(value);
        Assert.Equal(hex, Convert.ToHexString(encoded.ToBytes()));
        Assert.Equal(value, encoded.ToDouble());
    }

    [Theory]
    [InlineData(1e-30)]
    [InlineData(6.02214076e23)]
    [InlineData(-0.000123)]
    public void MicrosoftDecimal_RoundTrips(double value)
    {
        Assert.Equal(value, MicrosoftDecimal.FromDouble(value).ToDouble(), 12);
    }

    [Fact]
    public void Create_WriteAndRead_RoundTrips()
    {
        var document = ChartDocument.Create(ChartType.ColumnOrBar);
        var chart = document.MainChart;
        chart.IsStacked = true;
        chart.LegendPosition = LegendPosition.Vertical;
        chart.ValueScale.AutoMaximum = false;
        chart.ValueScale.Maximum = 250;
        chart.Texts.Add(ChartText.Create(ChartTextKind.ChartTitle, "Sales"));
        var label = ChartText.Create(ChartTextKind.DataLabel, string.Empty);
        label.IsAutomaticText = true;
        label.Text = null;
        chart.Texts.Add(label);

        var first = ChartSeries.Create([1, 2.5, 4], "Q1", "=R1C1:R1C3");
        first.InFirstList = true;
        document.Series.Add(first);
        document.Series.Add(ChartSeries.Create([10, 20], "Q2", "=R2C1:R2C2"));

        var bytes = document.ToArray();
        var reread = ChartDocument.Read(bytes);

        Assert.Equal(ChartType.ColumnOrBar, reread.MainChart.Type);
        Assert.True(reread.MainChart.IsStacked);
        Assert.Equal(LegendPosition.Vertical, reread.MainChart.LegendPosition);
        Assert.Equal(250, reread.MainChart.ValueScale.Maximum);
        Assert.Equal("Sales", reread.MainChart.Texts[0].Text);
        Assert.Null(reread.MainChart.Texts[1].Text);
        Assert.Equal(2, reread.Series.Count);
        Assert.Equal([1, 2.5, 4], reread.Series[0].SecondValues.Select(v => v.ToDouble()));
        Assert.Equal("Q2", reread.Series[1].Name);
        Assert.Equal("=R2C1:R2C2", reread.Series[1].FirstReference);
        Assert.Equal(bytes, reread.ToArray());
    }

    [Fact]
    public void Create_Scatter_HasXScale()
    {
        var bytes = ChartDocument.Create(ChartType.Scatter).ToArray();
        var reread = ChartDocument.Read(bytes);
        Assert.NotNull(reread.MainChart.XScale);
        Assert.True(reread.MainChart.XScale!.AutoMinimum);
    }

    [Fact]
    public void Read_WorksheetBytes_Throws()
    {
        Assert.Throws<ExcelFormatException>(() => ChartDocument.Read(new byte[200]));
    }
}
