using DocumentFormat.OpenXml.Drawing.Charts;
using SlideSharp.Enums;
using SlideSharp.Models;

namespace SlideSharp
{
    /// <summary>
    /// 图表
    /// </summary>
    public class Graphs
    {
        //private readonly DataBag PptData;

        internal readonly Slides _slides;

        internal Graphs(Slides slides)
        {
            _slides = slides;
        }

        private PieCharts _pieCharts;
        public PieCharts PieCharts
        {
            get 
            { 
                if(null == _pieCharts)
                {
                    _pieCharts = new PieCharts(this);
                }
                return _pieCharts; 
            }
        }
    }
}
