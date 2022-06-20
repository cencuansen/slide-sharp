using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Reflection;
using System.Xml;

namespace SlideSharp.Utils
{
    static class SlideUtils
    {
        /// <summary>
        /// 根据属性名获取属性值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        internal static object GetPropertyValue(this object obj, string key)
        {
            try
            {
                var param = Expression.Parameter(obj.GetType(), "p"); // [Type] p
                var prop = Expression.Property(param, key); // prop = p["key"]
                var getter = Expression.Lambda(prop, param);
                return getter.Compile().DynamicInvoke(obj);
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 根据属性名获取属性值
        /// </summary>
        /// <typeparam name="T">对象类型</typeparam>
        /// <param name="obj">对象</param>
        /// <param name="name">属性名</param>
        /// <returns>属性的值</returns>
        [Obsolete]
        internal static object GetPropertyValue2<T>(this T obj, string name)
        {
            try
            {
                Type type = obj.GetType();
                PropertyInfo p = type.GetProperty(name);
                if (p == null)
                {
                    throw new Exception(string.Format("该类型没有名为{0}的属性", name));
                }
                var param_obj = Expression.Parameter(typeof(T));
                var param_val = Expression.Parameter(typeof(object));
                //转成真实类型，防止Dynamic类型转换成object
                var body_obj = Expression.Convert(param_obj, type);
                var body = Expression.Property(body_obj, p);
                var getValue = Expression.Lambda<Func<T, object>>(body, param_obj).Compile();
                return getValue(obj);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 根据属性名称设置属性的值
        /// </summary>
        /// <typeparam name="T">对象类型</typeparam>
        /// <param name="t">对象</param>
        /// <param name="name">属性名</param>
        /// <param name="value">属性的值</param>
        internal static void SetPropertyValue<T>(this T t, string name, object value)
        {
            Type type = t.GetType();
            PropertyInfo p = type.GetProperty(name);
            if (p == null)
            {
                throw new Exception(string.Format("该类型没有名为{0}的属性", name));
            }
            var param_obj = Expression.Parameter(type);
            var param_val = Expression.Parameter(typeof(object));
            var body_obj = Expression.Convert(param_obj, type);
            var body_val = Expression.Convert(param_val, p.PropertyType);

            //获取设置属性的值的方法
            var setMethod = p.GetSetMethod(true);

            //如果只是只读,则setMethod==null
            if (setMethod != null)
            {
                var body = Expression.Call(param_obj, p.GetSetMethod(), body_val);
                var setValue = Expression.Lambda<Action<T, object>>(body, param_obj, param_val).Compile();
                setValue(t, value);
            }
        }

        /// <summary>
        /// 判断对象上是否存在指定字段
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="key"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        internal static bool HasField(this object obj, string key, Type type = null)
        {
            var properties = obj.GetType().GetProperties();

            if (type == null)
            {
                return properties.Any(property => property.Name == key);
            }
            else
            {
                return properties.Any(property => property.Name == key && property.PropertyType == type);
            }
        }

        /// <summary>
        /// 获取图片流
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static Stream GetStream(string url)
        {
            try
            {
                return WebRequest.Create(url).GetResponseAsync().GetAwaiter().GetResult().GetResponseStream();
            }
            catch
            {
                return null;
            }
        }

        ///// <summary>
        ///// 生成id
        ///// </summary>
        ///// <param name="prefix"></param>
        ///// <param name="randomPostfix"></param>
        ///// <returns></returns>
        //internal static string GenerateId(string prefix = "", bool randomPostfix = true)
        //{
        //    return $"{prefix}{DateTimeOffset.UtcNow.Ticks}{(randomPostfix ? new Random().Next(100, 999).ToString() : string.Empty)}";
        //}

        /// <summary>
        /// 单位转换：像素转EMU (English Metric Units)
        /// </summary>
        /// <param name="pixel"></param>
        /// <returns></returns>
        internal static long Pixel2EMU(long pixel)
        {
            return pixel * 9525;
        }

        /// <summary>
        /// 单位转换：EMU转像素
        /// </summary>
        /// <param name="eum"></param>
        /// <returns></returns>
        internal static long EMU2Pixel(long eum)
        {
            return eum / 9525 + 1;
        }

        /// <summary>
        /// 判断是否是图片文件地址
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static bool IsPicture(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            if (url.StartsWith("http://") || url.StartsWith("https://"))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 所有值都是字符串
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal static string ValueConvert(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }
            else if (value.GetType() == typeof(int))
            {
                return ((int)value).ToString();
            }
            else if (value.GetType() == typeof(decimal) || value.GetType() == typeof(double) || value.GetType() == typeof(float))
            {
                return Math.Round((decimal)value, 2).ToString();
            }
            else if (value.GetType() == typeof(string))
            {
                return (string)value;
            }

            return value.ToString();
        }

        /// <summary>
        /// 解析xml
        /// </summary>
        /// <param name="outXml"></param>
        /// <returns></returns>
        internal static string ParseXml(string outXml)
        {
            if (string.IsNullOrWhiteSpace(outXml))
            {
                return string.Empty;
            }
            var nameSpaces = new Dictionary<string, string>
            {
                { "a", "http://schemas.openxmlformats.org/drawingml/2006/main" },
                { "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" },
                { "p", "http://schemas.openxmlformats.org/presentationml/2006/main" },
                { "c", "http://schemas.openxmlformats.org/drawingml/2006/chart" },
            };
            var nameSpaceList = nameSpaces.Select(nameSpace => @$"xmlns:{nameSpace.Key}=""{nameSpace.Value}""");
            var nameSpaceString = nameSpaceList.StringJoin(" ");
            var xml = @$"<root {nameSpaceString}>{outXml}</root>";
            XmlDocument doc = new XmlDocument();
            //doc.LoadXml(HttpUtility.HtmlDecode(xml));
            doc.LoadXml(xml);
            return doc.FirstChild.FirstChild.OuterXml;
        }
    }
}
