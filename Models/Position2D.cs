using System;
using System.Collections.Generic;
using System.Text;

namespace SlideSharp.Models
{
    public class Position2D
    {
        public Position2D(long x = 0, long y = 0)
        {
            X = x;
            Y = y;
        }

        public long X { get; set; } = 0;
        public long Y { get; set; } = 0;
    }
}
