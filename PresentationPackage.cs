using DocumentFormat.OpenXml.Packaging;
using System;

namespace SlideSharp
{
    public sealed class PresentationPackage : IDisposable
    {
        /// <summary>
        /// dispose 标志
        /// </summary>
        private bool _disposed = false;

        /// <summary>
        /// 文档实例
        /// </summary>
        internal PresentationDocument Document { get; }

        /// <summary>
        /// 幻灯片
        /// </summary>
        private Slides _slides;
        public Slides Slides
        {
            get
            {
                if (null == _slides)
                {
                    _slides = new Slides(this);
                }

                return _slides;
            }
        }

        /// <summary>
        /// 模板幻灯片实例
        /// </summary>
        private MasterSlides _masterSlides;
        public MasterSlides MasterSlides
        {
            get
            {
                if (null == _masterSlides)
                {
                    _masterSlides = new MasterSlides(this);
                }

                return _masterSlides;
            }
        }


        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="file"></param>
        public PresentationPackage(string file)
        {
            Document = PresentationDocument.Open(file, true);
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            Document.Save();
        }


        /// <summary>
        /// 释放句柄
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                _slides = null;
                _masterSlides = null;
            }
            if (null != Document)
            {
                Document.Dispose();
            }

            _disposed = true;
        }

        ~PresentationPackage()
        {
            if (!_disposed)
            {
                Dispose(false);
            }
        }
    }
}
