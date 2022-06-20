using DocumentFormat.OpenXml.Presentation;
using SlideSharp.Constants;
using SlideSharp.Utils;
using System.Linq;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace SlideSharp
{
    /// <summary>
    /// 段落，文本框
    /// </summary>
    public class Paragraphs
    {
        internal readonly Slides _slides;

        internal Paragraphs(Slides slides)
        {
            _slides = slides;
        }

        /// <summary>
        /// 坐标X
        /// </summary>
        public long X { get; set; }

        /// <summary>
        /// 坐标Y
        /// </summary>
        public long Y { get; set; }

        /// <summary>
        /// 段落宽度
        /// </summary>
        public long Width { get; set; } = 100;

        /// <summary>
        /// 段落高度
        /// </summary>
        public long Height { get; set; } = 30;

        public sbyte PitchFamily { get; set; } = 100;

        private Runs _runs;
        public Runs Runs
        {
            get
            {
                if ( null == _runs)
                {
                    _runs = new Runs(this);
                }
                return _runs;
            }
        }

        /// <summary>
        /// 文字是否加粗
        /// </summary>
        public bool Bold { get; set; } = false;

        /// <summary>
        /// 是否更改
        /// </summary>
        public bool Dirty { get; set; } = false;

        /// <summary>
        /// 文字方向是否为从右向左
        /// </summary>
        public bool RightToLeftColumns { get; set; } = false;

        /// <summary>
        /// 字体
        /// </summary>
        public string TypeFace { get; set; } = Typeface.MicrosoftYaHei;

        public string Language { get; set; } = "zh-CN";
        /// <summary>
        /// 索引器
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        public Paragraphs this[long x, long y, long width = 100, long height = 100]
        {
            get
            {
                X = x;
                Y = y;
                Width = width;
                Height = height;
                return this;
            }
        }

        /// <summary>
        /// 文字大小
        /// </summary>
        public int FontSize { get; set; } = 14;

        /// <summary>
        /// 生成文本框
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        private Drawing.Paragraph NewParagraph(string content)
        {
            // 段落、文本框
            return new Drawing.Paragraph(
                new Drawing.Run(
                    new Drawing.RunProperties(
                        new Drawing.LatinFont() { Typeface = TypeFace, PitchFamily = PitchFamily, CharacterSet = 0 },
                        new Drawing.EastAsianFont() { Typeface = TypeFace, PitchFamily = PitchFamily, CharacterSet = 0 },
                        new Drawing.ComplexScriptFont() { Typeface = TypeFace, PitchFamily = PitchFamily, CharacterSet = 0 })
                    { FontSize = FontSize * 100, Bold = Bold, Dirty = Dirty },
                    new Drawing.Text() { Text = content },
                    new Drawing.EndParagraphRunProperties(
                        new Drawing.LatinFont() { Typeface = TypeFace, PitchFamily = PitchFamily, CharacterSet = 0 },
                        new Drawing.EastAsianFont() { Typeface = TypeFace, PitchFamily = PitchFamily, CharacterSet = 0 },
                        new Drawing.ComplexScriptFont() { Typeface = TypeFace, PitchFamily = PitchFamily, CharacterSet = 0 })
                    { FontSize = FontSize * 100, Bold = Bold, Dirty = Dirty }
                    )
                );
        }

        public Paragraphs SetFontSize(int fontSize)
        {
            this.FontSize = fontSize;
            return this;
        }

        /// <summary>
        ///  新增段落（文本框）
        /// </summary>
        /// <param name="content"></param>
        public void AddParagraph(string content)
        {
            var shape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = 1, Name = "text" },
                    new NonVisualShapeDrawingProperties() { TextBox = true },
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset() { X = SlideUtils.Pixel2EMU(X), Y = SlideUtils.Pixel2EMU(Y) },
                    new Drawing.Extents() { Cx = SlideUtils.Pixel2EMU(Width), Cy = SlideUtils.Pixel2EMU(Height) }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                new Drawing.NoFill()),
                new TextBody(
                new Drawing.BodyProperties(new Drawing.ShapeAutoFit())
                {
                    Wrap = Drawing.TextWrappingValues.Square,
                    RightToLeftColumns = RightToLeftColumns
                },
                new Drawing.ListStyle(),
                NewParagraph(content)));

            _slides.SlidePart.Slide.CommonSlideData.ShapeTree.Append(shape);
        }

        /// <summary>
        /// 删除包含指定文字的文本框
        /// </summary>
        /// <param name="keyword"></param>
        public void Remove(string keyword)
        {
            var paragraphs = _slides.SlidePart.Slide.Descendants<Drawing.Paragraph>();
            paragraphs.ToList().ForEach(paragraph =>
            {
                if (paragraph.InnerText.Contains(keyword))
                {
                    paragraph.Parent.Parent.Remove();
                }
            });
        }
    }
}
