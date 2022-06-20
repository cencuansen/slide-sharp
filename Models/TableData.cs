using SlideSharp.Enums;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SlideSharp.Models
{
    public class TableData
    {
        public int Idx { get; set; } = 2;

        /// <summary>
        /// 名称
        /// </summary>
        public string Title { get; set; } = "Table1";
        /// <summary>
        /// 位置x，单位：像素
        /// </summary>
        public long X { get; set; } = 0;
        /// <summary>
        /// 位置y，单位：像素
        /// </summary>
        public long Y { get; set; } = 0;
        /// <summary>
        /// 首行
        /// </summary>
        internal bool FirstRow { get; set; } = false;
        /// <summary>
        /// 条纹
        /// </summary>
        internal bool BandRow { get; set; } = false;
        /// <summary>
        /// 各行高，单位：像素
        /// </summary>
        [Required, MinLength(1)]
        public IEnumerable<long> RowsHeight { get; set; } = new List<long>();
        /// <summary>
        /// 各列宽，单位：像素
        /// </summary>
        [Required, MinLength(1)]
        public IEnumerable<long> ColumnsWidth { get; set; } = new List<long>();
        /// <summary>
        /// 水平位置
        /// </summary>
        public HorizontalPosition Horizontal { get; set; } = HorizontalPosition.Default;
        /// <summary>
        /// 垂直位置
        /// </summary>
        public VerticalPosition Vertical { get; set; } = VerticalPosition.Default;

        /// <summary>
        /// 四周margin
        /// </summary>
        public long MarginLeft { get; set; }
        public long MarginRight { get; set; }
        public long MarginTop { get; set; }
        public long MarginBottom { get; set; }
    }
}
