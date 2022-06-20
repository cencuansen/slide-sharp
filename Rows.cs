using DocumentFormat.OpenXml.Drawing;
using System.Linq;

namespace SlideSharp
{
    /// <summary>
    /// 表格的行
    /// </summary>
    public class Rows
    {
        internal readonly Tables _tables;

        internal Rows(Tables tables)
        {
            _tables = tables;
        }

        /// <summary>
        /// 表格总行数
        /// </summary>
        public int Count => _tables.Table.Descendants<TableRow>().Count();

        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="rowIndex"></param>
        public void Remove(int rowIndex)
        {
            var rows = _tables.Table.Elements<TableRow>();
            var row = rows.ElementAtOrDefault(rowIndex);
            row?.Remove();
        }
    }
}
