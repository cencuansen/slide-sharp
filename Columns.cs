using DocumentFormat.OpenXml.Drawing;
using System.Collections.Generic;
using System.Linq;

namespace SlideSharp
{
    public class Columns
    {
        internal readonly Tables _tables;

        internal Columns(Tables tables)
        {
            _tables = tables;
        }

        /// <summary>
        /// 删除列
        /// </summary>
        /// <param name="columnIndex"></param>
        public void Remove(int columnIndex)
        {
            // 删除网格
            _tables.Table.Descendants<GridColumn>().ElementAtOrDefault(columnIndex).Remove();

            // 删除数据
            var deleteCells = new List<TableCell>();
            var rows = _tables.Table.Elements<TableRow>();
            foreach (var row in rows)
            {
                var cell = row.Elements<TableCell>()?.ElementAtOrDefault(columnIndex);
                if (cell != null)
                {
                    deleteCells.Add(cell);
                }
            }

            foreach (var cell in deleteCells)
            {
                cell.Remove();
            }
        }
    }
}
