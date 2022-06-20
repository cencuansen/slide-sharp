using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using SlideSharp.Utils;
using System.Collections.Generic;
using System.Linq;
using Presentation = DocumentFormat.OpenXml.Presentation;

namespace SlideSharp
{
    /// <summary>
    /// 表格的单元格
    /// </summary>
    public class Cells
    {
         internal readonly Tables _tables;

        /// <summary>
        /// 网格
        /// </summary>
        internal TableGrid TableGrid => GetTableGrid(_tables.Table);
        /// <summary>
        /// 网格列
        /// </summary>
        internal GridColumn GridColumn => GetGridColumn(TableGrid, ColumnIndex);
        /// <summary>
        /// 行节点
        /// </summary>
        internal TableRow TableRow => GetTableRow(_tables.Table, RowIndex);
        /// <summary>
        /// 单元格节点
        /// </summary>
        internal TableCell TableCell => GetTableCell(TableRow, ColumnIndex);
        /// <summary>
        /// 行索引值
        /// </summary>
        internal int RowIndex { get; set; }
        /// <summary>
        /// 列索引值
        /// </summary>
        internal int ColumnIndex { get; set; }
        /// <summary>
        /// 结束行
        /// </summary>
        internal int EndRowIndex { get; set; }
        /// <summary>
        /// 结束列
        /// </summary>
        internal int EndColumnIndex { get; set; }
        /// <summary>
        /// 单元格宽
        /// </summary>
        public long Width => SlideUtils.EMU2Pixel(GridColumn.Width);
        /// <summary>
        /// 单元格高
        /// </summary>
        public long Height => SlideUtils.EMU2Pixel(TableRow.Height);
        /// <summary>
        /// 设置或读取单元格值
        /// </summary>
        public object Value
        {
            get => GetValue();
            set => SetValue(value);
        }

        /// <summary>
        /// 字体大小
        /// </summary>
        public int FontSize
        {
            set
            {
                for (var i = RowIndex; i <= EndRowIndex; i++)
                {
                    var row = GetTableRow(_tables.Table, i);
                    for (var j = ColumnIndex; j <= EndColumnIndex; j++)
                    {
                        var cell = GetTableCell(row, j);
                        SetFontSize(cell, value);
                    }
                }
            }
        }

        /// <summary>
        /// 文本是否加粗
        /// </summary>
        public bool Bold
        {
            set
            {
                for (var i = RowIndex; i <= EndRowIndex; i++)
                {
                    var row = GetTableRow(_tables.Table, i);
                    for (var j = ColumnIndex; j <= EndColumnIndex; j++)
                    {
                        var cell = GetTableCell(row, j);
                        SetBold(cell, value);
                    }
                }
            }
        }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public string Color
        {
            set
            {
                for (var i = RowIndex; i <= EndRowIndex; i++)
                {
                    var row = GetTableRow(_tables.Table, i);
                    for (var j = ColumnIndex; j <= EndColumnIndex; j++)
                    {
                        var cell = GetTableCell(row, j);
                        SetFontColor(cell, value);
                    }
                }
            }
        }

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public TextAlignmentTypeValues Alignment
        {
            set
            {
                for (var i = RowIndex; i <= EndRowIndex; i++)
                {
                    var row = GetTableRow(_tables.Table, i);
                    for (var j = ColumnIndex; j <= EndColumnIndex; j++)
                    {
                        var cell = GetTableCell(row, j);
                        SetAlignment(cell, value);
                    }
                }
            }
        }

        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        public TextAnchoringTypeValues Vertical
        {
            set
            {
                for (var i = RowIndex; i <= EndRowIndex; i++)
                {
                    var row = GetTableRow(_tables.Table, i);
                    for (var j = ColumnIndex; j <= EndColumnIndex; j++)
                    {
                        var cell = GetTableCell(row, j);
                        SetAnchoring(cell, value);
                    }
                }
            }
        }

        /// <summary>
        /// 背景颜色，如："FF0000"
        /// </summary>
        public HexBinaryValue BackgroundColor
        {
            set
            {
                for (var i = RowIndex; i <= EndRowIndex; i++)
                {
                    var row = GetTableRow(_tables.Table, i);
                    for (var j = ColumnIndex; j <= EndColumnIndex; j++)
                    {
                        var cell = GetTableCell(row, j);
                        SetBackgroundColor(cell, value);
                    }
                }
            }
        }

        /// <summary>
        /// 索引器
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public Cells this[int rowIndex, int columnIndex]
        {
            get
            {
                RowIndex = rowIndex;
                ColumnIndex = columnIndex;
                EndRowIndex = rowIndex;
                EndColumnIndex = columnIndex;
                GetRealIndex();
                return this;
            }
        }

        /// <summary>
        /// 索引器
        /// </summary>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        public Cells this[int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex]
        {
            get
            {
                RowIndex = startRowIndex;
                ColumnIndex = startColumnIndex;
                EndRowIndex = endRowIndex;
                EndColumnIndex = endColumnIndex;
                GetRealIndex();
                return this;
            }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="tables"></param>
        internal Cells(Tables tables)
        {
            this._tables = tables;
        }

        /// <summary>
        /// 取单元格值
        /// </summary>
        /// <returns></returns>
        private string GetValue()
        {
            var textBody = TableCell.GetFirstChild<TextBody>();
            return textBody?.InnerText ?? string.Empty;
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="value"></param>
        private void SetValue(object value)
        {
            if (value is string && SlideUtils.IsPicture(value as string))
            {
                SetPicture(value as string);
                return;
            }

            var p = TableCell.Descendants<Paragraph>().FirstOrDefault();
            var eprp = TableCell.Descendants<EndParagraphRunProperties>().FirstOrDefault();

            // 通过OuterXml实现元素节点的深拷贝
            var eprpString = eprp.OuterXml;
            var copyedRp = eprpString.Replace("a:endParaRPr", "a:rPr");

            var pp = p.GetFirstChild<ParagraphProperties>();
            p.RemoveAllChildren<Run>();
            var run = new Run(new RunProperties(copyedRp), new Text(value?.ToString()));
            p.InsertAfter(run, pp);
        }

        /// <summary>
        /// 合并的单元格重定位
        /// </summary>
        private void GetRealIndex()
        {
            // 合并的单元格数据设置
            while (TableCell.HorizontalMerge?.Value ?? false)
            {
                ColumnIndex--;
            }

            while (TableCell.VerticalMerge?.Value ?? false)
            {
                RowIndex--;
            }
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="pictureUrl"></param>
        public void SetPicture(string pictureUrl)
        {
            (long x, long y, long width, long height) = GetCellPositionAndSize(RowIndex, ColumnIndex);
            var pic = new Pictures(_tables._slides);
            pic.Create(pictureUrl, x, y, width, height);
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="path"></param>
        /// <param name="margin"></param>
        public void SetPicture(string path, int margin)
        {
            (long x, long y, long width, long height) = GetCellPositionAndSize(RowIndex, ColumnIndex);
            var pic = new Pictures(_tables._slides);
            pic.Create(path, x + margin, y + margin, width - margin * 2, height - margin * 2);
        }

        /// <summary>
        /// 设置单元格背景颜色
        /// </summary>
        /// <param name="cell">目标单元格</param>
        /// <param name="hexColorString">格式如：FF0000</param>
        private void SetBackgroundColor(TableCell cell, HexBinaryValue hexColorString)
        {
            var tcp = cell.TableCellProperties ?? cell.AppendChild(new TableCellProperties());
            tcp.RemoveAllChildren<NoFill>(); // NoFill和SolidFill不能同时存在于同一个父元素下
            tcp.RemoveAllChildren<SolidFill>();
            tcp.AppendChild(new SolidFill(new RgbColorModelHex() { Val = hexColorString.Value.ToUpper() }));
        }

        /// <summary>
        /// 设置单元格字体颜色
        /// </summary>
        /// <param name="cell">目标单元格</param>
        /// <param name="hexColorString">格式如：FF0000</param>
        private void SetFontColor(TableCell cell, HexBinaryValue hexColorString)
        {
            var paragraph = cell.Descendants<Paragraph>().FirstOrDefault();
            // 获取存放文本的节点
            var run = paragraph.Elements<Run>().FirstOrDefault();
            if (run == null)
            {
                paragraph.Append(new Run(new RunProperties(new SolidFill { RgbColorModelHex = new RgbColorModelHex { Val = hexColorString } })));
            }
            else
            {
                var runProperties = run.Elements<RunProperties>().FirstOrDefault();
                if (runProperties == null)
                {
                    run.Append(new RunProperties(new SolidFill { RgbColorModelHex = new RgbColorModelHex { Val = hexColorString } }));
                }
                else
                {
                    runProperties.RemoveAllChildren<SolidFill>();
                    runProperties.AppendChild(new SolidFill { RgbColorModelHex = new RgbColorModelHex { Val = hexColorString } });
                }
            }
        }

        /// <summary>
        /// 设置文本水平位置
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="alignmentType"></param>
        public void SetAlignment(TableCell cell, TextAlignmentTypeValues alignmentType)
        {
            var paragraphs = GetParagraphs(cell).ToList();
            foreach (var p in paragraphs)
            {
                var paragraphProperties = p.Elements<ParagraphProperties>();
                if (paragraphProperties.Count() == 0)
                {
                    paragraphProperties = new List<ParagraphProperties> { p.AppendChild(new ParagraphProperties()) };
                }

                foreach (var property in paragraphProperties)
                {
                    property.Alignment = alignmentType;
                }
            }
        }

        /// <summary>
        /// 设置文本垂直位置
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="AnchoringType"></param>
        private void SetAnchoring(TableCell cell, TextAnchoringTypeValues AnchoringType)
        {
            var tableCellProperties = cell.Elements<TableCellProperties>();
            if (tableCellProperties.Count() == 0)
            {
                tableCellProperties = new List<TableCellProperties> { cell.AppendChild(new TableCellProperties()) };
            }

            foreach (var property in tableCellProperties)
            {
                property.Anchor = AnchoringType;
            }
        }

        /// <summary>
        /// 设置单元格文本加粗
        /// </summary>
        /// <param name="tableCell"></param>
        /// <param name="isBold"></param>
        private void SetBold(TableCell tableCell, bool isBold)
        {
            var eprps = tableCell.Descendants<EndParagraphRunProperties>();
            foreach (var property in eprps)
            {
                property.Bold = isBold;
            }
            var runProps = tableCell.Descendants<RunProperties>();
            foreach (var property in runProps)
            {
                property.Bold = isBold;
            }
        }

        /// <summary>
        /// 设置单元格文本字体大小
        /// </summary>
        private void SetFontSize(TableCell tableCell, int size)
        {
            var eprps = tableCell.Descendants<EndParagraphRunProperties>();
            foreach (var property in eprps)
            {
                property.FontSize = size * 100;
            }
            var runProps = tableCell.Descendants<RunProperties>();
            foreach (var property in runProps)
            {
                property.FontSize = size * 100;
            }
        }

        /// <summary>
        /// 计算单元格位置和长宽
        /// </summary>
        /// <returns></returns>
        public (long x, long y, long width, long height) GetCellPositionAndSize(int rowIndex, int columnIndex)
        {
            // 表格和幻灯片间的距离
            var transform = _tables.GraphicFrame.Descendants<Presentation.Transform>().FirstOrDefault();
            var tableX = SlideUtils.EMU2Pixel(transform.Offset.X);
            var tableY = SlideUtils.EMU2Pixel(transform.Offset.Y);

            // 计算图片位置
            long height2Top = tableY;
            long width2Left = tableX;
            long cellWidth = SlideUtils.EMU2Pixel(_tables.Table.Descendants<TableGrid>().First().Descendants<GridColumn>().ElementAt(columnIndex).Width);
            long cellHeight = SlideUtils.EMU2Pixel(_tables.Table.Elements<TableRow>().ElementAt(rowIndex).Height);

            // y轴距离
            for (int i = 0; i < rowIndex; i++)
            {
                height2Top += SlideUtils.EMU2Pixel(_tables.Table.Elements<TableRow>().ElementAt(i).Height);
            }

            // x轴距离
            for (int i = 0; i < columnIndex; i++)
            {
                width2Left += SlideUtils.EMU2Pixel(_tables.Table.Descendants<TableGrid>().First().Descendants<GridColumn>().ElementAt(i).Width);
            }

            return (width2Left, height2Top, cellWidth, cellHeight);
        }

        /// <summary>
        /// 获取网格
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private TableGrid GetTableGrid(Table table)
        {
            return table.Descendants<TableGrid>().FirstOrDefault();
        }

        /// <summary>
        /// 获取网格指定序号列
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private GridColumn GetGridColumn(TableGrid grid, int index)
        {
            return grid.Descendants<GridColumn>().ElementAt(index);
        }

        /// <summary>
        /// 由 Table 获取 TableRow
        /// </summary>
        /// <param name="table"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public TableRow GetTableRow(Table table, int index)
        {
            return table.Descendants<TableRow>().ElementAt(index);
        }

        /// <summary>
        /// 由 TableRow 获取 TableCell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private TableCell GetTableCell(TableRow row, int index)
        {
            return row.Descendants<TableCell>().ElementAt(index);
        }

        /// <summary>
        /// 获取单元格段落
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private IEnumerable<Paragraph> GetParagraphs(TableCell cell)
        {
            return cell.Descendants<Paragraph>();
        }
    }
}
