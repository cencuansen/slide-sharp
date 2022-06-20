using DocumentFormat.OpenXml.Presentation;
using SlideSharp.Constants;
using SlideSharp.Utils;
using System.Linq;
using System.Text.RegularExpressions;
using D = DocumentFormat.OpenXml.Drawing;

namespace SlideSharp
{
    /// <summary>
    /// Master和Layout的关系：一个Master下包含一个或多个Layout
    /// </summary>
    public class MasterSlides
    {
        private readonly PresentationPackage _ppt;

        //internal MasterSlides(DataBag pptDate)
        //{
        //    PptDate = pptDate;
        //}

        internal MasterSlides(PresentationPackage ppt)
        {
            _ppt = ppt;
        }

        private int index = 0;

        public MasterSlides this[int index]
        {
            get
            {
                this.index = index;
                return this;
            }
        }

        /// <summary>
        /// 替换模板字符串
        /// </summary>
        /// <param name="datas"></param>
        public void Replace(object datas)
        {
            if (datas == null)
            {
                return;
            }

            // 模板，至少有一个
            var SlideMasterPart1 = _ppt.Document.PresentationPart.SlideMasterParts.ToList()[index];
            var SlideMaster = SlideMasterPart1.SlideMaster;
            var shapes = SlideMaster.CommonSlideData.Elements<ShapeTree>().FirstOrDefault().Elements<Shape>().ToList();

            foreach (var shape in shapes)
            {
                // 幻灯片中全部段落（文本框中的文字）
                var paragraphs = shape.Descendants<D.Paragraph>();
                foreach (var paragraph in paragraphs)
                {
                    var runs = paragraph.Descendants<D.Run>().Where(x => Regex.IsMatch(x.InnerText, Consts.Patten)).ToList();
                    foreach (var run in runs)
                    {
                        var oldString = run.InnerText;
                        var matchedText = Regex.Match(oldString, Consts.Patten).Groups[1].Value;
                        var newString = new Regex(Consts.Patten).Replace(oldString, datas.GetPropertyValue(matchedText)?.ToString() ?? string.Empty);
                        // 赋值更改
                        run.Text = new D.Text(newString);
                    }
                }
            }
        }
    }
}
