using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using SlideSharp.Enums;
using SlideSharp.Models;
using SlideSharp.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideSharp
{
    /// <summary>
    /// 表格
    /// </summary>
    public class Tables
    {
        internal readonly Slides _slides;

        internal IEnumerable<P.GraphicFrame> GraphicFrames => _slides.SlidePart.Slide.Descendants<P.GraphicFrame>();

        /// <summary>
        /// 表格节点
        /// </summary>
        internal P.GraphicFrame GraphicFrame
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Title))
                {
                    throw new Exception("请提供表格名称");
                }

                var graphicFrame = GraphicFrames?.FirstOrDefault(gf =>
                {
                    var innerTitle = gf.Descendants<P.NonVisualDrawingProperties>()?.FirstOrDefault()?.Title ?? string.Empty;
                    if (innerTitle == Title)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                });

                return graphicFrame;
            }
        }

        /// <summary>
        /// 表格节点
        /// </summary>
        internal Table Table => GraphicFrame?.Descendants<Table>().FirstOrDefault();

        /// <summary>
        /// 表格个数
        /// </summary>
        public int Count => GraphicFrames?.Count() ?? 0;

        /// <summary>
        /// 第一个表格
        /// </summary>
        public Tables First
        {
            get
            {
                var first = GraphicFrames.FirstOrDefault();
                if (first == null)
                {
                    throw new IndexOutOfRangeException("无表格");
                }
                var firstProp = first?.Descendants<P.NonVisualDrawingProperties>()?.FirstOrDefault();
                if (string.IsNullOrWhiteSpace(firstProp.Title))
                {
                    firstProp.Title = "default";
                    firstProp.Name = "default";
                }
                Title = firstProp.Title;
                return this;
            }
        }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; private set; }

        /// <summary>
        /// 行
        /// </summary>
        private Rows _rows;
        public Rows Rows
        {
            get
            {
                if(null == _rows)
                {
                    _rows = new Rows(this);
                }
                return _rows;
            }
        }

        /// <summary>
        /// 列
        /// </summary>
        private Columns _columns;
        public Columns Columns
        {
            get
            {
                if (null == _columns)
                {
                    _columns = new Columns(this);
                }
                return _columns;
            }
        }

        /// <summary>
        /// 单元格
        /// </summary>
        private Cells _cells;
        public Cells Cells
        {
            get
            {
                if (null == _cells)
                {
                    _cells = new Cells(this);
                }
                return _cells;
            }
        }

        /// <summary>
        /// 索引器
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public Tables this[string name]
        {
            get
            {
                Title = name;
                return this;
            }
        }

        internal Tables(Slides slides)
        {
            _slides = slides;
        }

        /// <summary>
        /// 添加表格
        /// </summary>
        /// <param name="tableData"></param>
        public Tables Create(TableData tableData)
        {
            if (string.IsNullOrWhiteSpace(tableData.Title))
            {
                throw new ArgumentNullException("表格Title不能省略");
            }

            var tableWidth = tableData.ColumnsWidth.Sum();
            var slideWidth = _slides.Width;
            var tableHeight = tableData.RowsHeight.Sum();
            var slideHeight = _slides.Height;

            // 计算水平位置
            if (tableData.Horizontal == HorizontalPosition.Center)
            {
                tableData.X = slideWidth / 2 - tableWidth / 2 + tableData.MarginLeft - tableData.MarginRight;
            }
            else if (tableData.Horizontal == HorizontalPosition.Left)
            {
                tableData.X = 0 + tableData.MarginLeft - tableData.MarginRight;
            }
            else if (tableData.Horizontal == HorizontalPosition.Right)
            {
                tableData.X = slideWidth - tableWidth + tableData.MarginLeft - tableData.MarginRight;
            }
            // 计算垂直位置
            if (tableData.Vertical == VerticalPosition.Center)
            {
                tableData.Y = slideHeight / 2 - tableHeight / 2 + tableData.MarginTop - tableData.MarginBottom;
            }
            else if (tableData.Vertical == VerticalPosition.Top)
            {
                tableData.Y = 0 + tableData.MarginTop - tableData.MarginBottom;
            }
            else if (tableData.Vertical == VerticalPosition.Bottom)
            {
                tableData.Y = slideHeight - tableHeight + tableData.MarginTop - tableData.MarginBottom;
            }

            _slides.TableId = Math.Max(tableData.Idx, _slides.TableId);
            Title = $"{tableData.Title}{_slides.TableId}";

            // 影响列
            var gridCol = Enumerable.Range(0, tableData.ColumnsWidth.Count()).Select(columnIndex => $@"<a:gridCol w = ""{SlideUtils.Pixel2EMU(tableData.ColumnsWidth.ToList()[columnIndex])}"" />").StringJoin();
            var tr = Enumerable.Range(0, tableData.RowsHeight.Count()).Select(rowIndex => $@"<a:tr h = ""{SlideUtils.Pixel2EMU(tableData.RowsHeight.ToArray()[rowIndex])}"">
                    {Enumerable.Range(0, tableData.ColumnsWidth.Count()).Select(columnIndex => $@"<a:tc>
                        <a:txBody>
                          <a:bodyPr />
                          <a:p>
                            <a:pPr>
                              <a:buNone />
                            </a:pPr>
                            <a:endParaRPr lang = ""zh-CN"" altLang=""en-US"" />
                          </a:p>
                        </a:txBody>
                        <a:tcPr>
                          <a:lnL w = ""3175"" cmpd=""sng"">
                            <a:solidFill>
                              <a:schemeClr val = ""tx1"" />
                            </a:solidFill>
                            <a:prstDash val = ""solid"" />
                          </a:lnL>
                          <a:lnR w = ""3175"" cmpd=""sng"">
                            <a:solidFill>
                              <a:schemeClr val = ""tx1"" />
                            </a:solidFill>
                            <a:prstDash val = ""solid"" />
                          </a:lnR>
                          <a:lnT w = ""3175"" cmpd=""sng"">
                            <a:solidFill>
                              <a:schemeClr val = ""tx1"" />
                            </a:solidFill>
                            <a:prstDash val = ""solid"" />
                          </a:lnT>
                          <a:lnB w = ""3175"" cmpd=""sng"">
                            <a:solidFill>
                              <a:schemeClr val = ""tx1"" />
                            </a:solidFill>
                            <a:prstDash val = ""solid"" />
                          </a:lnB>
                          <a:noFill />
                        </a:tcPr>
                      </a:tc>").StringJoin()} 
                    </a:tr>").StringJoin();

            var graphicFrameXml = $@"<p:graphicFrame>
                  <p:nvGraphicFramePr>
                    <p:cNvPr id = ""{_slides.TableId++}"" name=""{Title}"" title=""{Title}""/>
                    <p:cNvGraphicFramePr />
                    <p:nvPr />
                  </p:nvGraphicFramePr>
                  <p:xfrm>
                    <a:off x = ""{SlideUtils.Pixel2EMU(tableData.X)}"" y=""{SlideUtils.Pixel2EMU(tableData.Y)}"" />
                    <a:ext cx = ""{SlideUtils.Pixel2EMU(tableData.ColumnsWidth.Sum())}"" cy=""{SlideUtils.Pixel2EMU(tableData.RowsHeight.Sum())}"" />
                  </p:xfrm>
                  <a:graphic>
                    <a:graphicData uri = ""http://schemas.openxmlformats.org/drawingml/2006/table"" >
                      <a:tbl>
                        <a:tblPr firstRow = ""{Convert.ToInt32(tableData.FirstRow)}"" bandRow=""{Convert.ToInt32(tableData.BandRow)}"">
                          <a:tableStyleId>{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}</a:tableStyleId>
                        </a:tblPr>
                        <a:tblGrid>
                          {gridCol}
                        </a:tblGrid>
                        {tr}
                      </a:tbl>
                    </a:graphicData>
                  </a:graphic>
                </p:graphicFrame>";

            var parsedXml = SlideUtils.ParseXml(graphicFrameXml);
            var gf = new P.GraphicFrame(parsedXml);
            SlidePart slidePart = _slides.SlidePart;
            var graphicFrame = slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(gf);
            slidePart.Slide.Save();
            return this;
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endRow"></param>
        /// <param name="endColumn"></param>
        public void Merge(int startRow, int startColumn, int endRow, int endColumn)
        {
            for (var x = startRow; x <= endRow; x++)
            {
                for (var y = startColumn; y <= endColumn; y++)
                {
                    var ce = Cells[x, y].TableCell;

                    // 首行全部设置 rowspan
                    if (startRow == x)
                    {
                        ce.RowSpan = endRow - startRow + 1;
                    }
                    else
                    {
                        ce.VerticalMerge = true;
                    }

                    // 首列全部设置 gridspan
                    if (startColumn == y)
                    {
                        ce.GridSpan = endColumn - startColumn + 1;
                    }
                    else
                    {
                        ce.HorizontalMerge = true;
                    }
                }
            }
        }
    }
}
