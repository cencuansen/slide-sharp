using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideSharp.Models;
using SlideSharp.Utils;
using System.Collections.Generic;
using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideSharp
{
    public class PieCharts
    {
        internal readonly Graphs _graphs;

        internal PieCharts(Graphs graphs)
        {
            this._graphs = graphs;
        }

        public Position2D Position2D { get; set; } = new Position2D(0, 0);

        public Size2D Size2D { get; set; } = new Size2D(100, 100);

        /// <summary>
        /// 标签数值格式
        /// </summary>
        public string NumberingFormat { get; set; } = "0.00%";

        /// <summary>
        /// 数据标签位置
        /// </summary>
        public DataLabelPositionValues DataLabelPosition { get; set; } = DataLabelPositionValues.OutsideEnd;

        /// <summary>
        /// 是否显示图例
        /// </summary>
        public bool ShowLegendKey { get; set; } = true;

        /// <summary>
        /// 是否显示图例项文字
        /// </summary>
        public bool ShowValue { get; set; } = true;

        /// <summary>
        /// 是否显示分类名
        /// </summary>
        public bool ShowCategoryName { get; set; } = true;

        /// <summary>
        /// 是否显示系列名
        /// </summary>
        public bool ShowSeriesName { get; set; } = false;

        /// <summary>
        /// 是否显示百分百
        /// </summary>
        public bool ShowPercent { get; set; } = true;

        /// <summary>
        /// 是否显示气泡
        /// </summary>
        public bool ShowBubbleSize { get; set; } = false;

        /// <summary>
        /// 是否显示引导线
        /// </summary>
        public bool ShowLeaderLines { get; set; } = true;

        /// <summary>
        /// 图例位置
        /// </summary>
        public LegendPositionValues LegendPosition { get; set; } = LegendPositionValues.Bottom;

        /// <summary>
        /// 数据标签文字大小
        /// </summary>
        public int LabelFontSize { get; set; } = 10;

        /// <summary>
        /// 标签文字是否加粗
        /// </summary>
        public bool LabelTextBold { get; set; } = false;

        /// <summary>
        /// 标签分隔符
        /// </summary>
        public string LabelSeparator { get; set; } = ",";

        /// <summary>
        /// 标签文字颜色：黑色（000000）~ 白色（ffffff）
        /// </summary>
        public string LabelFontColor { get; set; } = "000000";

        /// <summary>
        /// 标题字体大小
        /// </summary>
        public int TitleFontSize { get; set; } = 20;

        /// <summary>
        /// 标题文字
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// 新增饼图
        /// </summary>
        /// <param name="datas"></param>
        public void Create(IDictionary<string, decimal> datas)
        {
            if (datas == null || datas.Count == 0)
            {
                return;
            }

            var _slides = _graphs._slides;

            SlidePart slidePart = _slides.SlidePart;
            CommonSlideData comSlddata = slidePart.Slide.CommonSlideData;
            ShapeTree shapeTree = comSlddata.ShapeTree;
            P.GraphicFrame graphicFrame = new();

            P.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties = new();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new() { Id = 1U, Name = "Graph_1" };
            P.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties = new();
            ApplicationNonVisualDrawingProperties applicationnonVisualDrawingProperties = new();
            nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
            nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
            nonVisualGraphicFrameProperties.Append(applicationnonVisualDrawingProperties);

            P.Transform transform = new();
            D.Offset offset = new() { X = SlideUtils.Pixel2EMU(Position2D.X), Y = SlideUtils.Pixel2EMU(Position2D.Y) };
            D.Extents extents = new() { Cx = SlideUtils.Pixel2EMU(Size2D.Width), Cy = SlideUtils.Pixel2EMU(Size2D.Height) };
            transform.Append(offset);
            transform.Append(extents);

            var chartReferenceId = $"rId{_slides.GraphId++}"; //SlideUtils.GenerateId("rId", false);
            graphicFrame.Append(nonVisualGraphicFrameProperties);
            graphicFrame.Append(transform);
            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference { Id = chartReferenceId }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));
            shapeTree.Append(graphicFrame);

            #region 饼图数据信息 - 数据的项和对应的值
            int index = 0;
            StringReference stringReference = new();
            stringReference.AppendChild(new C.Formula());
            var stringCache = stringReference.AppendChild(new C.StringCache());
            stringCache.AppendChild(new PointCount { Val = (uint)datas.Keys.Count });
            datas.Keys.ToList().ForEach(key =>
            {
                // 项
                stringCache.Append(new StringPoint(new NumericValue(key)) { Index = new UInt32Value((uint)index++) });
            });
            index = 0;
            NumberReference numberReference = new();
            numberReference.AppendChild(new C.Formula());
            NumberingCache numberingCache = numberReference.AppendChild(new C.NumberingCache());
            numberingCache.AppendChild(new FormatCode("General"));
            numberingCache.AppendChild(new PointCount { Val = (uint)datas.Keys.Count });
            datas.Keys.ToList().ForEach(key =>
            {
                // 项值
                numberingCache.Append(new NumericPoint(new NumericValue($"{datas[key]}")) { Index = new UInt32Value((uint)index++) });
            });
            #endregion

            #region 饼图主体区
            var pie = new PieChart(
                new PieChartSeries(
                    new Index() { Val = 0U },
                    new Order() { Val = 0U },
                    new DataLabels(
                        new TextProperties(new BodyProperties(), new ListStyle(), new Paragraph(new ParagraphProperties(new DefaultRunProperties(new SolidFill(new RgbColorModelHex { Val = LabelFontColor })) { Bold = LabelTextBold, FontSize = LabelFontSize * 100 }))),
                        new NumberingFormat { FormatCode = NumberingFormat, SourceLinked = false },
                        new DataLabelPosition { Val = DataLabelPosition },
                        new ShowLegendKey() { Val = ShowLegendKey },
                        new ShowValue() { Val = ShowValue },
                        new ShowCategoryName() { Val = ShowCategoryName },
                        new ShowSeriesName() { Val = ShowSeriesName },
                        new ShowPercent() { Val = ShowPercent },
                        new ShowBubbleSize() { Val = ShowBubbleSize },
                        new Separator(LabelSeparator),
                        new ShowLeaderLines() { Val = ShowLeaderLines }),
                    new CategoryAxisData(stringReference),
                    new Values(numberReference)));
            #endregion

            #region 图例
            Legend legend = new();
            LegendPosition legendPosition = new() { Val = LegendPosition };
            Layout legendLayout = new();
            Overlay legendOverLay = new();
            ChartShapeProperties chartShapeProperties = new(new NoFill(), new Outline(new NoFill()), new EffectList());
            TextProperties textProperties = new(new BodyProperties(), new ListStyle(), new Paragraph(new ParagraphProperties(new DefaultRunProperties(new SolidFill(new RgbColorModelHex { Val = LabelFontColor })) { Bold = LabelTextBold, FontSize = LabelFontSize * 100 })));
            legend.AppendChild(legendPosition);
            legend.AppendChild(legendLayout);
            legend.AppendChild(legendOverLay);
            legend.AppendChild(chartShapeProperties);
            legend.AppendChild(textProperties);
            #endregion

            ChartPart newChartPart = slidePart.AddNewPart<ChartPart>(chartReferenceId);
            newChartPart.ChartSpace = new ChartSpace(new C.Chart(new Title(ChartTitleXml()), new C.PlotArea(new Layout(), pie), legend));
            slidePart.Slide.Save();
        }

        private string ChartTitleXml()
        {
            string xml = @$"<c:title xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:tx><c:rich><a:bodyPr rot=""0"" spcFirstLastPara=""0"" vertOverflow=""ellipsis"" vert=""horz"" wrap=""square"" anchor=""ctr"" anchorCtr=""1"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /><a:lstStyle xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /><a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:pPr defTabSz=""914400""><a:defRPr lang=""zh-CN"" sz=""1400"" b=""0"" i=""0"" u=""none"" strike=""noStrike"" kern=""1200"" spc=""0"" baseline=""0""><a:solidFill><a:schemeClr val=""tx1""><a:lumMod val=""65000"" /><a:lumOff val=""35000"" /></a:schemeClr></a:solidFill><a:latin typeface=""+mn-lt"" /><a:ea typeface=""+mn-ea"" /><a:cs typeface=""+mn-cs"" /></a:defRPr></a:pPr><a:r><a:rPr lang=""zh-CN"" altLang=""en-US"" /><a:t>{Title}</a:t></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val=""0"" /><c:spPr><a:noFill xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /><a:ln xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:noFill /></a:ln><a:effectLst xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /></c:spPr></c:title>";
            return SlideUtils.ParseXml(xml);
        }
    }
}
