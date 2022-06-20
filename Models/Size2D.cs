using System;
using System.Collections.Generic;
using System.Text;

namespace SlideSharp.Models
{
    public class Size2D
    {
        public Size2D(long width = 100, long height = 100)
        {
            Width = width;
            Height = height;
        }

        /// <summary>
        /// 单位：像素
        /// </summary>
        public long Width { get; set; } = 100;

        /// <summary>
        /// 单位：像素
        /// </summary>
        public long Height { get; set; } = 100;
    }
}
