using System.Collections.Generic;
using System.Linq;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace SlideSharp
{
    public class Runs
    {
        internal readonly Paragraphs _paragraphs;

        internal Runs(Paragraphs paragraphs)
        {
            this._paragraphs = paragraphs;
        }

        /// <summary>
        /// 去除包含指定文本的文本行：Run
        /// </summary>
        /// <param name="keyword"></param>
        public void Remove(string keyword)
        {
            _paragraphs._slides.SlidePart.Slide.Descendants<Drawing.Run>()
                .Where(run => run.InnerText.Contains(keyword)).ToList()
                .ForEach(matchedParagraph => matchedParagraph.Parent.Remove());
        }
    }
}
