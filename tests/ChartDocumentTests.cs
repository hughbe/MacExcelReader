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
    public void MicrosoftDecimal_HandlesBoundsAndValidation()
    {
        Assert.Throws<ArgumentException>(() => MicrosoftDecimal.FromBytes(new byte[7]));
        Assert.Throws<ArgumentOutOfRangeException>(() => MicrosoftDecimal.FromDouble(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => MicrosoftDecimal.FromDouble(double.PositiveInfinity));
        Assert.Equal(MicrosoftDecimal.Zero, MicrosoftDecimal.FromDouble(-0));
        Assert.True(MicrosoftDecimal.FromDouble(2) == MicrosoftDecimal.FromDouble(2));
        Assert.True(MicrosoftDecimal.FromDouble(2) != MicrosoftDecimal.FromDouble(3));
        Assert.True(MicrosoftDecimal.FromDouble(2).Equals((object)MicrosoftDecimal.FromDouble(2)));
        Assert.False(MicrosoftDecimal.FromDouble(2).Equals(2));
        Assert.Equal(MicrosoftDecimal.FromDouble(2).GetHashCode(), MicrosoftDecimal.FromDouble(2).GetHashCode());
        Assert.Equal("2", MicrosoftDecimal.FromDouble(2).ToString());
        Assert.Equal(9.9999999999999E+62, MicrosoftDecimal.FromDouble(1E+100).ToDouble());
        Assert.Equal(1E-64, MicrosoftDecimal.FromDouble(1E-100).ToDouble());
    }

    [Fact]
    public void ChartOptions_RoundTrip()
    {
        var document = ChartDocument.Create(ChartType.ColumnOrBar);
        var chart = document.MainChart;
        chart.IsBar = true;
        chart.IsStacked = true;
        chart.IsPercent = true;
        chart.VaryByCategory = true;
        chart.AutomaticSeriesDistribution = true;
        chart.FirstOverlaySeries = 7;
        chart.Overlap = -25;
        chart.ClusterSpacing = 80;
        chart.FirstSliceAngle = 45;
        chart.ValueAxisCrossesAtCategory = 9;
        chart.ValueAxisCrossesBetweenCategories = false;
        chart.ValueAxisCrossesAtMaximumCategory = true;
        chart.CategoriesInReverseOrder = true;
        chart.CategoriesBetweenTickLabels = 3;
        var axis = chart.CategoryAxis;
        axis.MinorTickMarks = TickMarkType.Cross;
        axis.MajorGridlines.Pattern = ChartPattern.LightGray;
        axis.MinorGridlines.Pattern = ChartPattern.DarkGray;
        var scale = chart.ValueScale;
        scale.Minimum = -10;
        scale.Maximum = 20;
        scale.MajorUnit = 5;
        scale.MinorUnit = 1;
        scale.CrossesAt = 2;
        scale.AutoMinimum = false;
        scale.AutoMaximum = false;
        scale.AutoMajorUnit = false;
        scale.AutoMinorUnit = false;
        scale.AutoCrossesAt = false;
        scale.Logarithmic = true;
        scale.ReverseOrder = true;

        var reread = ChartDocument.Read(document.ToArray()).MainChart;
        Assert.True(reread.IsBar && reread.IsStacked && reread.IsPercent);
        Assert.True(reread.VaryByCategory && reread.AutomaticSeriesDistribution);
        Assert.Equal(7, reread.FirstOverlaySeries);
        Assert.Equal(-25, reread.Overlap);
        Assert.Equal(80, reread.ClusterSpacing);
        Assert.Equal(45, reread.FirstSliceAngle);
        Assert.Equal(9, reread.ValueAxisCrossesAtCategory);
        Assert.False(reread.ValueAxisCrossesBetweenCategories);
        Assert.True(reread.ValueAxisCrossesAtMaximumCategory);
        Assert.True(reread.CategoriesInReverseOrder);
        Assert.Equal(3, reread.CategoriesBetweenTickLabels);
        Assert.Equal(TickMarkType.Cross, reread.CategoryAxis.MinorTickMarks);
        Assert.Equal(ChartPattern.LightGray, reread.CategoryAxis.MajorGridlines.Pattern);
        Assert.Equal(ChartPattern.DarkGray, reread.CategoryAxis.MinorGridlines.Pattern);
        Assert.Equal(-10, reread.ValueScale.Minimum);
        Assert.Equal(20, reread.ValueScale.Maximum);
        Assert.Equal(5, reread.ValueScale.MajorUnit);
        Assert.Equal(1, reread.ValueScale.MinorUnit);
        Assert.Equal(2, reread.ValueScale.CrossesAt);
        Assert.True(reread.ValueScale.Logarithmic && reread.ValueScale.ReverseOrder);
        Assert.Equal(4, reread.CategoryAxis.Font.Length);
    }

    [Fact]
    public void ChartTextOptions_RoundTrip()
    {
        var document = ChartDocument.Create(ChartType.Line);
        var text = ChartText.Create(ChartTextKind.DataLabel, "Sales");
        text.SeriesNumber = 2;
        text.PointNumber = 255;
        text.IsVertical = true;
        text.HorizontalAlignment = 2;
        text.VerticalAlignment = 1;
        text.ShowKey = true;
        text.ShowValue = true;
        document.MainChart.Texts.Add(text);
        var reread = ChartDocument.Read(document.ToArray()).MainChart.Texts.Single();

        Assert.Equal(ChartTextKind.DataLabel, reread.Kind);
        Assert.Equal(2, reread.SeriesNumber);
        Assert.Equal(255, reread.PointNumber);
        Assert.True(reread.IsVertical);
        Assert.Equal(2, reread.HorizontalAlignment);
        Assert.Equal(1, reread.VerticalAlignment);
        Assert.True(reread.ShowKey && reread.ShowValue);
        Assert.Equal("DataLabel: Sales", reread.ToString());
        Assert.Throws<ArgumentException>(() => new ChartText(new byte[45]));
    }

    [Fact]
    public void ChartDocument_PreservesOptionalCollectionsAndStreams()
    {
        var document = ChartDocument.Create(ChartType.ColumnOrBar);
        document.CodeSegment = [1, 2, 3];
        document.PageSetup = Enumerable.Range(0, 664).Select(i => (byte)i).ToArray();
        document.Charts.Insert(0, Chart.Create(ChartType.Line));
        document.MainChart.SeriesWords.AddRange([10, 20]);
        document.MainChart.PointFormats.Add(new byte[12]);

        var textSeries = ChartSeries.Create("Categories", [], "Data", "R2C1:R2C2", "R1C1:R1C2");
        textSeries.SeriesType = 3;
        textSeries.FirstTexts.AddRange(["One", "Two"]);
        textSeries.SecondValues.AddRange([MicrosoftDecimal.FromDouble(3), MicrosoftDecimal.FromDouble(4)]);
        document.Series.Add(textSeries);

        var path = Path.GetTempFileName();
        try
        {
            document.Save(path);
            var reread = ChartDocument.Load(path);
            Assert.NotNull(reread.OverlayChart);
            Assert.Equal([1, 2, 3], reread.CodeSegment);
            Assert.Equal(document.PageSetup, reread.PageSetup);
            Assert.Equal([10, 20], reread.MainChart.SeriesWords);
            Assert.Single(reread.MainChart.PointFormats);
            Assert.Equal(["One", "Two"], reread.Series[0].FirstTexts);
            Assert.Equal([3, 4], reread.Series[0].SecondValues.Select(v => v.ToDouble()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ChartDocument_RejectsUnsupportedVersionsAndChartCounts()
    {
        var versionOne = ChartDocument.Create(ChartType.Line).ToArray();
        versionOne[3] = 1;
        Assert.Throws<ExcelFormatException>(() => ChartDocument.Read(versionOne));

        var noCharts = ChartDocument.Create(ChartType.Line);
        noCharts.Charts.Clear();
        Assert.Throws<InvalidOperationException>(() => noCharts.ToArray());
        noCharts.Charts.AddRange([Chart.Create(ChartType.Line), Chart.Create(ChartType.Line), Chart.Create(ChartType.Line)]);
        Assert.Throws<InvalidOperationException>(() => noCharts.ToArray());
    }

    [Fact]
    public void ChartValueTypes_ValidateConstructorsAndFlags()
    {
        Assert.Throws<ArgumentNullException>(() => new Chart(null!));
        Assert.Throws<ArgumentNullException>(() => new ChartSeries(null!));
        Assert.Throws<ArgumentNullException>(() => new ChartText(null!));
        Assert.Throws<ArgumentException>(() => new Chart(new byte[Chart.RecordSize - 1]));
        Assert.Throws<ArgumentException>(() => new ChartSeries(new byte[ChartSeries.HeaderSize - 1]));

        var chart = Chart.Create(ChartType.Line);
        Assert.Null(chart.XScale);
        chart.XAxisScale = new byte[Chart.ScaleSize];
        Assert.NotNull(chart.XScale);
        var text = ChartText.Create(ChartTextKind.ChartTitle, "x");
        text.IsAutomaticText = true;
        text.IsAutomaticText = false;
    }

    [Fact]
    public void ChartDocument_RejectsInvalidFixedSizeBuffers()
    {
        var chartHeader = ChartDocument.Create(ChartType.Line);
        chartHeader.ChartSetHeader = new byte[7];
        Assert.Throws<InvalidOperationException>(() => chartHeader.ToArray());

        var window = ChartDocument.Create(ChartType.Line);
        window.Window = new byte[11];
        Assert.Throws<InvalidOperationException>(() => window.ToArray());

        var pageSetup = ChartDocument.Create(ChartType.Line);
        pageSetup.PageSetup = new byte[663];
        Assert.Throws<InvalidOperationException>(() => pageSetup.ToArray());

        var scale = ChartDocument.Create(ChartType.Scatter);
        scale.MainChart.XAxisScale = new byte[Chart.ScaleSize - 1];
        Assert.Throws<InvalidOperationException>(() => scale.ToArray());

        var pointFormat = ChartDocument.Create(ChartType.Line);
        pointFormat.MainChart.PointFormats.Add(new byte[11]);
        Assert.Throws<InvalidOperationException>(() => pointFormat.ToArray());
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

        var first = ChartSeries.Create("Q1", [1, 2.5, 4], "Data", "R1C1:R1C3");
        first.InFirstList = true;
        document.Series.Add(first);
        document.Series.Add(ChartSeries.Create("Q2", [10, 20], "Data", "R2C1:R2C2"));

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
        Assert.Equal("R2C1:R2C2", reread.Series[1].SecondReference);
        Assert.Equal("Data", reread.Series[1].SourceDocument);
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
