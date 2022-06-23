using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideSharp.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace SlideSharp
{
    /// <summary>
    /// 图片
    /// </summary>
    public class Pictures
    {
        internal readonly Slides _slides;

        /// <summary>
        /// 图片宽
        /// </summary>
        internal long Width { get; set; } = 100;

        /// <summary>
        /// 图片高
        /// </summary>
        internal long Height { get; set; } = 100;

        /// <summary>
        /// 图片水平方向上的位置
        /// </summary>
        internal long PositionX { get; set; } = 0;

        /// <summary>
        /// 图片垂直方向上的位置
        /// </summary>
        internal long PositionY { get; set; } = 0;

        internal string Name { get; set; } = string.Empty;

        internal Pictures(Slides slides)
        {
            _slides = slides;
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="path"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void Create(string path, long x, long y, long width, long height)
        {
            // 生成流
            Stream stream = SlideUtils.GetStream(path);
            if (stream == null)
            {
                return;
            }
            (long originWidth, long originHeight) = GetSize(stream, out var outStream);
            (long positionX, long positionY, long outWidth, long outHeight) = CalculatePositionAndSize(originWidth, originHeight, width, height, x, y);
            Create(outStream, positionX, positionY, outWidth, outHeight);
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void Create(Stream stream, long x, long y, long width, long height)
        {
            try
            {
                var pic = new Picture(PictureXml(x, y, width, height));

                Slide slide = _slides.SlidePart.Slide;
                slide.CommonSlideData!.ShapeTree!.Append(pic);
                ImagePart imagePart = slide.SlidePart!.AddImagePart(ImagePartType.Jpeg);
                imagePart.FeedData(stream);

                // pic节点关联图片数据
                pic.BlipFill!.Blip!.Embed = slide.SlidePart.GetIdOfPart(imagePart);

                stream.Dispose();
            }
            catch
            {
            }
        }

        /// <summary>
        /// 生成图片节点
        /// </summary>
        /// <returns></returns>
        private string PictureXml(long x, long y, long width, long height)
        {
            var xml = $@"<p:pic><p:nvPicPr><p:cNvPr id=""{_slides.PictureId++}"" name=""image{_slides.PictureId++}""/><p:cNvPicPr><p:picLocks noChangeAspect=""1"" /></p:cNvPicPr ><p:nvPr /></p:nvPicPr><p:blipFill><a:blip r:embed=""rId{_slides.PictureId++}"" /><a:stretch><a:fillRect /></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=""{SlideUtils.Pixel2EMU(x)}"" y=""{SlideUtils.Pixel2EMU(y)}"" /><a:ext cx=""{SlideUtils.Pixel2EMU(width)}"" cy=""{SlideUtils.Pixel2EMU(height)}"" /></a:xfrm><a:prstGeom prst=""rect""><a:avLst /></a:prstGeom ><a:ln w=""12700"" cmpd=""sng""><a:solidFill><a:schemeClr val=""bg1""><a:lumMod val=""85000"" /></a:schemeClr></a:solidFill><a:prstDash val=""solid"" /></a:ln></p:spPr></p:pic>";
            return SlideUtils.ParseXml(xml);
        }

        /// <summary>
        /// 设置图片宽高
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="newStream"></param>
        private (long positionX, long positionY, long width, long height) SetPictureSize(Stream stream, out Stream newStream)
        {
            (long originWidth, long originHeight) = GetSize(stream, out var outStream);
            newStream = outStream;
            return CalculatePositionAndSize(originWidth, originHeight, Width, Height, PositionX, PositionY);
        }

        /// <summary>
        /// 根据图片流获取图片尺寸
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outStream"></param>
        /// <returns></returns>
        private (long width, long height) GetSize(Stream? stream, out Stream? outStream)
        {
            long innerWidth = 0, innerHeight = 0;
            outStream = null;

            if (null == stream)
            {
                return (innerWidth, innerHeight);
            }

            try
            {
                using Bitmap img = Image.FromStream(stream) as Bitmap;
                innerWidth = img!.Width;
                innerHeight = img.Height;
                MemoryStream ms = new();
                img.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                stream.Dispose();
                outStream = ms;
            }
            catch (Exception)
            {
                outStream = default;
            }

            return (innerWidth, innerHeight);
        }

        /// <summary>
        /// 在含有特定文字的文本框里插入图片
        /// </summary>
        /// <param name="keyword"></param>
        /// <param name="path"></param>
        public void Replace(string keyword, string path)
        {
            if (string.IsNullOrWhiteSpace(keyword) || string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            // 生成流
            Stream stream = SlideUtils.GetStream(path);
            Replace(keyword, stream);
        }

        /// <summary>
        /// 在含有特定文字的文本框里插入多个图片
        /// </summary>
        /// <param name="keyword"></param>
        /// <param name="paths"></param>
        /// <param name="maxColumnCount"></param>
        /// <param name="margin"></param>
        public void Replace(string keyword, IEnumerable<string> paths, int maxColumnCount = 2, int margin = 0)
        {
            if (paths == null || paths.Count() == 0)
            {
                return;
            }

            var streamList = new List<Stream>();
            foreach (var path in paths)
            {
                Stream stream = SlideUtils.GetStream(path);
                if (stream == null)
                {
                    continue;
                }
                streamList.Add(stream);
            }

            Replace(keyword, streamList, maxColumnCount, margin);
        }

        /// <summary>
        /// 在含有特定文字的文本框里插入图片
        /// </summary>
        /// <param name="keyword"></param>
        /// <param name="imageStream"></param>
        public void Replace(string keyword, Stream imageStream)
        {
            if (string.IsNullOrWhiteSpace(keyword) || imageStream == null)
            {
                return;
            }
            var shape = _slides.SlidePart.Slide.Descendants<Shape>().FirstOrDefault(x => x.InnerText.Contains(keyword));
            var transform = shape.ShapeProperties!.Transform2D;
            long pX = SlideUtils.EMU2Pixel(transform.Offset.X);
            long pY = SlideUtils.EMU2Pixel(transform.Offset.Y);
            long shapeWidth = SlideUtils.EMU2Pixel(transform.Extents.Cx) - 2;
            long shapeHeight = SlideUtils.EMU2Pixel(transform.Extents.Cy) - 2;
            (long originWidth, long originHeight) = GetSize(imageStream, out var newStream);
            (long positionX, long positionY, long width, long height) = CalculatePositionAndSize(originWidth, originHeight, shapeWidth, shapeHeight, pX, pY);
            Create(newStream, positionX, positionY, width, height);
        }

        /// <summary>
        /// 在含有特定文字的文本框里插入图片
        /// </summary>
        /// <param name="keyword"></param>
        /// <param name="imageStreams"></param>
        /// <param name="maxColumnCount"></param>
        /// <param name="margin"></param>
        public void Replace(string keyword, IList<Stream> imageStreams, int maxColumnCount = 2, int margin = 0)
        {
            if (string.IsNullOrWhiteSpace(keyword) || imageStreams == null || imageStreams.Count() == 0)
            {
                return;
            }

            var rowCount = (int)Math.Ceiling(imageStreams.Count() / (decimal)maxColumnCount);
            var columnCount = rowCount > 1 ? maxColumnCount : imageStreams.Count();

            var shape = _slides.SlidePart.Slide.Descendants<Shape>().FirstOrDefault(x => x.InnerText.Contains(keyword));
            var transform = shape.ShapeProperties.Transform2D;
            long pX = SlideUtils.EMU2Pixel(transform.Offset.X);
            long pY = SlideUtils.EMU2Pixel(transform.Offset.Y);
            long shapeWidth = SlideUtils.EMU2Pixel(transform.Extents.Cx) - 2;
            long shapeHeight = SlideUtils.EMU2Pixel(transform.Extents.Cy) - 2;

            var cellWidth = shapeWidth / columnCount;
            var cellHeight = shapeHeight / rowCount;
            var streamIndex = 0;
            for (var x = 0; x < rowCount; x++)
            {
                for (var y = 0; y < columnCount; y++)
                {
                    if (streamIndex < imageStreams.Count())
                    {
                        var innerImageStream = imageStreams[streamIndex++];
                        (long originWidth, long originHeight) = GetSize(innerImageStream, out var newStream);
                        (long positionX, long positionY, long width, long height) = CalculatePositionAndSize(originWidth, originHeight, cellWidth, cellHeight, pX + y * cellWidth, pY + x * cellHeight, margin);
                        Create(newStream, positionX, positionY, width, height);
                    }
                }
            }
        }

        /// <summary>
        /// 最大适应尺寸计算
        /// </summary>
        /// <param name="originWidth">图片原始宽度</param>
        /// <param name="originHeight">图片原始高度</param>
        /// <param name="targetWidth">目标宽度</param>
        /// <param name="targetHeight">目标高度</param>
        /// <param name="startX"></param>
        /// <param name="startY"></param>
        /// <param name="margin"></param>
        /// <returns></returns>
        private (long positionX, long positionY, long width, long height) CalculatePositionAndSize(long originWidth, long originHeight, long targetWidth, long targetHeight, long startX, long startY, int margin = 0)
        {
            originWidth += margin * 2;
            originHeight += margin * 2;

            var shapeRatio = (double)targetWidth / targetHeight;
            var imageRatio = (double)originWidth / originHeight;
            long innerPositionX = 0, innerPositionY = 0, innerWidth = 0, innerHeight = 0;

            if (targetWidth > targetHeight && originWidth > originHeight)
            {
                if (shapeRatio > imageRatio)
                {
                    innerHeight = targetHeight;
                    innerWidth = (long)((double)originWidth / originHeight * targetHeight);
                    innerPositionY = startY;
                    innerPositionX = startX + (long)(((double)targetWidth - innerWidth) / 2);
                }
                else
                {
                    innerWidth = targetWidth;
                    innerHeight = (long)((double)originHeight / originWidth * targetWidth);
                    innerPositionX = startX;
                    innerPositionY = startY + (long)(((double)targetHeight - innerHeight) / 2);
                }
            }
            else if (targetWidth > targetHeight && originWidth <= originHeight)
            {
                innerHeight = targetHeight;
                innerWidth = (long)((double)originWidth / originHeight * targetHeight);
                innerPositionY = startY;
                innerPositionX = startX + (long)(((double)targetWidth - innerWidth) / 2);
            }
            else if (targetWidth <= targetHeight && originWidth <= originHeight)
            {
                if (shapeRatio > imageRatio)
                {
                    innerHeight = targetHeight;
                    innerWidth = (long)((double)originWidth / originHeight * targetHeight);
                    innerPositionY = startY;
                    innerPositionX = startX + (long)(((double)targetWidth - innerWidth) / 2);
                }
                else
                {
                    innerWidth = targetWidth;
                    innerHeight = (long)((double)originHeight / originWidth * targetWidth);
                    innerPositionX = startX;
                    innerPositionY = startY + (long)(((double)targetHeight - innerHeight) / 2);
                }
            }
            else if (targetWidth <= targetHeight && originWidth > originHeight)
            {
                innerWidth = targetWidth;
                innerHeight = (long)((double)originHeight / originWidth * targetWidth);
                innerPositionX = startX;
                innerPositionY = startY + (long)(((double)targetHeight - innerHeight) / 2);
            }
            return (innerPositionX + margin, innerPositionY + margin, innerWidth - 2 * margin, innerHeight - 2 * margin);
        }
    }
}
