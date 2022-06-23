using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideSharp.Constants;
using SlideSharp.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace SlideSharp
{
    /// <summary>
    /// 幻灯片
    /// </summary>
    public class Slides
    {
        internal readonly PresentationPackage _ppt;

        internal int GraphId = 2;
        internal int TableId = 2;
        internal int PictureId = 2;

        /// <summary>
        /// 幻灯片id集合
        /// </summary>
        public SlideIdList SlideIdList
        {
            get
            {
                SlideIdList slideIdList = _ppt.Document.PresentationPart.Presentation.SlideIdList;
                // 空白PPT时
                if (null == slideIdList)
                {
                    slideIdList = new SlideIdList();
                }
                return slideIdList;
            }
        }

        /// <summary>
        /// SlideId
        /// </summary>
        public SlideId SlideId
        {
            get
            {
                return SlideIdList.ElementAt(Index) as SlideId;
            }
        }

        /// <summary>
        /// 引用ID
        /// </summary>
        public string RelationshipId
        {
            get
            {
                return SlideId.RelationshipId;
            }
        }

        /// <summary>
        /// 幻灯片部分
        /// </summary>
        public SlidePart SlidePart
        {
            get
            {
                return (SlidePart)_ppt.Document.PresentationPart.GetPartById(RelationshipId);
            }
        }    

        /// <summary>
        /// 幻灯片个数
        /// </summary>
        public int Count => SlideIdList.Count();

        /// <summary>
        /// 幻灯片宽度
        /// </summary>
        public long Width => SlideUtils.EMU2Pixel(_ppt.Document.PresentationPart.Presentation.SlideSize.Cx);

        /// <summary>
        /// 幻灯片高度
        /// </summary>
        public long Height => SlideUtils.EMU2Pixel(_ppt.Document.PresentationPart.Presentation.SlideSize.Cy);

        /// <summary>
        /// 索引值
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// 笔记
        /// </summary>
        private Notes _notes;
        public Notes Notes
        {
            get
            {
                if (null == _notes)
                {
                    _notes = new Notes(this);
                }
                return _notes;
            }
        }

        /// <summary>
        /// 表格
        /// </summary>
        private Tables _tables;
        public Tables Tables
        {
            get
            {
                if (null == _tables)
                {
                    _tables = new Tables(this);
                }
                return _tables;
            }
        }

        /// <summary>
        /// 图片
        /// </summary>
        private Pictures _pictures;
        public Pictures Pictures
        {
            get
            {
                if (null == _pictures)
                {
                    _pictures = new Pictures(this);
                }
                return _pictures;
            }
        }

        /// <summary>
        /// 图表
        /// </summary>
        private Graphs _graphs;
        public Graphs Graphs
        {
            get
            {
                if (null == _graphs)
                {
                    _graphs = new Graphs(this);
                }
                return _graphs;
            }
        }

        /// <summary>
        /// 段落
        /// </summary>
        private Paragraphs _paragraphs;
        public Paragraphs Paragraphs
        {
            get
            {
                if (null == _paragraphs)
                {
                    _paragraphs = new Paragraphs(this);
                }
                return _paragraphs;
            }
        }

        /// <summary>
        /// 索引器
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Slides this[int index]
        {
            get
            {
                Index = index;
                return this;
            }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="ppt"></param>
        internal Slides(PresentationPackage ppt)
        {
            _ppt = ppt;
        }

        /// <summary>
        /// 替换指定幻灯片中的模板字符串
        /// </summary>
        /// <param name="datas"></param>
        public void Replace(object datas)
        {
            if (datas == null)
            {
                return;
            }

            // 幻灯片中全部段落（文本框中的文字）
            var paragraphs = SlidePart.Slide.Descendants<Paragraph>().ToList();
            foreach (var paragraph in paragraphs)
            {
                var paraInnerText = paragraph.InnerText;
                var matchGroups = Regex.Matches(paraInnerText, Consts.Patten);

                if (matchGroups.Count == 0)
                {
                    // 无匹配，此时整个Paragraph都可忽略
                    continue;
                }

                var allRuns = paragraph.Descendants<Run>().ToList();
                var removeRuns = new List<Run>();
                int index = 0;
                for (int i = 0; i < allRuns.Count; i++)
                {
                    var isMatched = false;
                    var innerTexts = string.Empty;
                    if (index == i)
                    {
                        innerTexts = allRuns[i].InnerText;
                    }
                    else
                    {
                        // 这里说明一个Patten匹配涉及到了多个Run
                        innerTexts = allRuns.Skip(index).Take(i - index + 1).Select(x => x.InnerText).StringJoin("");
                    }

                    isMatched = Regex.IsMatch(innerTexts, Consts.Patten);

                    if (isMatched)
                    {
                        var matched = Regex.Match(innerTexts, Consts.Patten);
                        var matchedPatten = matched.Value;
                        var matchedText = matched.Groups[1].Value;
                        var inputData = datas.GetPropertyValue(matchedText)?.ToString();
                        allRuns[i].Text = new D.Text(innerTexts.Replace(matchedPatten, inputData));

                        if (index == i)
                        {

                        }
                        else
                        {
                            var innerRemoveRuns = allRuns.Skip(index).Take(i - index).ToList();
                            removeRuns.AddRange(innerRemoveRuns);
                        }

                        index = i + 1;
                    }
                    else
                    {
                        // 说明当前所遍历到的Run还没到Pattern范围内，那么可以忽略掉这些Run
                        if (!innerTexts.Contains(Consts.Patten.First()))
                        {
                            index = i + 1;
                        }
                    }
                }
                removeRuns.ForEach(x =>
                {
                    paragraph.RemoveChild(x);
                });
            }
            SlidePart.Slide.Save();
        }

        /// <summary>
        /// 新建空白幻灯片
        /// </summary>
        /// <param name="targetIndex"></param>
        public Slides Create(int targetIndex)
        {
            Index = targetIndex;

            SlidePart slidePart = _ppt.Document.PresentationPart.AddNewPart<SlidePart>();
            // 获取第一个母板(SlideMasterPart)
            var slideMasterPart = _ppt.Document.PresentationPart.SlideMasterParts.First();
            // 获取母板中存有的布局集合信息(SlideLayoutIdList)
            var firstSlideLayoutId = slideMasterPart.SlideMaster.Descendants<SlideLayoutId>()?.FirstOrDefault();
            // 获取母版下指定id的布局
            var slideLayoutPart = slideMasterPart.GetPartById(firstSlideLayoutId.RelationshipId);
            slidePart.AddPart(slideLayoutPart, firstSlideLayoutId.RelationshipId);

            slidePart.Slide = new Slide();
            slidePart.Slide.Append(new CommonSlideData(CommonSlideDataXml()));
            Insert(slidePart, targetIndex);
            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private string CommonSlideDataXml()
        {
            string xml = @$"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=""1"" name="""" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>";
            return SlideUtils.ParseXml(xml);
        }

        /// <summary>
        /// 复制当前幻灯片，搭配 <see cref="Insert(SlidePart, int)"/> 来使用
        /// </summary>
        /// <returns></returns>
        public SlidePart Clone()
        {
            // 为新幻灯片创建 slide part
            var newSlidePart = _ppt.Document.PresentationPart.AddNewPart<SlidePart>();
            using (var templateStream = SlidePart.GetStream(FileMode.Open))
            {
                newSlidePart.FeedData(templateStream);
            }

            // 复制 Layout
            if (SlidePart.SlideLayoutPart != null)
            {
                newSlidePart.AddPart(SlidePart.SlideLayoutPart);
            }

            // 复制 Image
            if (SlidePart.ImageParts != null && SlidePart.ImageParts.Count() > 0)
            {
                foreach (ImagePart image in SlidePart.ImageParts)
                {
                    ImagePart imageClone = newSlidePart.AddImagePart(image.ContentType, SlidePart.GetIdOfPart(image));
                    using (var imageStream = image.GetStream())
                    {
                        imageClone.FeedData(imageStream);
                    }
                }
            }

            // 复制 Notes
            if (SlidePart.NotesSlidePart != null)
            {
                newSlidePart.AddPart(SlidePart.NotesSlidePart);
            }

            // 去除 CustomerData
            var customerDataList = newSlidePart.Slide.Descendants<CustomerDataList>().ToList();
            customerDataList.ForEach(data => data.RemoveAllChildren());

            // 去除 CustomerData
            var graphicFrames = newSlidePart.Slide.CommonSlideData.ShapeTree.Descendants<P.GraphicFrame>()?.ToList();
            graphicFrames?.ForEach(gf =>
            {
                gf.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.RemoveAllChildren();
            });

            return newSlidePart;
        }

        /// <summary>
        /// 指定位置插入幻灯片
        /// </summary>
        /// <param name="slidePart"></param>
        /// <param name="targetIndex"></param>
        public void Insert(SlidePart slidePart, int targetIndex)
        {
            // 插入新幻灯片
            SlideId newSlideId;
            if (!SlideIdList.Any())
            {
                newSlideId = SlideIdList.AppendChild(new SlideId());
            }
            else if (targetIndex <= 0)
            {
                newSlideId = SlideIdList.InsertBefore(new SlideId(), SlideIdList.First());
            }
            else if (targetIndex >= Count)
            {
                newSlideId = SlideIdList.InsertAfter(new SlideId(), SlideIdList.Last());
            }
            else
            {
                newSlideId = SlideIdList.InsertAfter(new SlideId(), (SlideId)SlideIdList.ToList()[targetIndex - 1]);
            }

            // 根据现有幻灯片最大id值计算新幻灯片的id值
            newSlideId.Id = NewSlideId();
            // 将 [幻灯片(Slide)] 和 [幻灯片id列表(SlideIdList)] 进行关联
            newSlideId.RelationshipId = _ppt.Document.PresentationPart.GetIdOfPart(slidePart);
            // 保存新的 slide part.
            slidePart.Slide.Save(slidePart);
        }

        /// <summary>
        /// 复制
        /// </summary>
        /// <param name="targetIndex"></param>
        public void Clone(int targetIndex)
        {
            Insert(Clone(), targetIndex);
        }

        /// <summary>
        /// 删除，下标从0开始
        /// </summary>
        public void Remove()
        {
            // 获得指定序号位置的幻灯片id
            var slideId = (SlideId)SlideIdList.ToList()[Index];

            // 幻灯片的关联id
            string OepnOepnOepnSlideRelId = slideId.RelationshipId;

            //// 根据id获取幻灯片实体
            SlidePart slidePart = _ppt.Document.PresentationPart.GetPartById(OepnOepnOepnSlideRelId) as SlidePart;

            // 删除幻灯片实体
            _ppt.Document.PresentationPart.DeletePart(slidePart);

            // 从幻灯片id列表中删除
            SlideIdList.RemoveChild(slideId);
        }

        /// <summary>
        /// 生成id
        /// </summary>
        /// <returns></returns>
        private uint NewSlideId()
        {
            // idlist中的id最小值是256
            uint maxSlideId = 256;

            if (SlideIdList.ChildElements == null || SlideIdList.ChildElements.Count() == 0)
            {
                return maxSlideId;
            }

            foreach (SlideId slideId in SlideIdList.ChildElements)
            {
                if (slideId.Id != null)
                {
                    maxSlideId = Math.Max(maxSlideId, slideId.Id);
                }
            }

            maxSlideId++;
            return maxSlideId;
        }
    }
}
