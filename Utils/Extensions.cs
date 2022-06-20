using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SlideSharp.Utils
{
    static class Extensions
    {
        /// <summary>
        /// 集合数据进行字符串join处理
        /// </summary>
        /// <param name="list"></param>
        /// <param name="split"></param>
        /// <returns></returns>
        internal static string StringJoin(this IEnumerable<string> list, string split = "")
        {
            if(list.ToList().Count() == 0)
            {
                return string.Empty;
            }
            return string.Join(split, list);
        }
    }
}
