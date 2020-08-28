using Alizhou.Office.Attribute;
using Alizhou.Office.Enum;
using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Alizhou.Office.Extensions
{
    public static class IWordExportTemplateExtensions
    {
        public static IEnumerable<PlaceholderEntity> GetReplacements<T>(this T wordData) where T : IWordExportTemplate
        {
            var placeholders = new List<PlaceholderEntity>();
            Type type = typeof(T);
            PropertyInfo[] props = type.GetProperties();
            foreach (var prop in props)
            {
                var placeholderEntity = new PlaceholderEntity();
                var placeholder = prop.IsDefined(typeof(PlaceholderAttribute)) ?
                  prop.GetCustomAttribute<PlaceholderAttribute>().Placeholder.ToString() : "{" + prop.Name + "}";

                if (prop.PropertyType == typeof(string))
                {
                    placeholderEntity.Placeholder = placeholder;
                    placeholderEntity.PlaceholderType = PlaceholderType.Text;
                    placeholderEntity.Data = new AlizhouText { Data = prop.GetValue(wordData)?.ToString() };
                    placeholders.Add(placeholderEntity);
                }
                else if (prop.PropertyType == typeof(AlizhouTable))
                {
                    placeholderEntity.Placeholder = placeholder;
                    placeholderEntity.PlaceholderType = PlaceholderType.Table;
                    placeholderEntity.Data = (AlizhouTable)prop.GetValue(wordData);
                    placeholders.Add(placeholderEntity);
                }
                else if (prop.PropertyType == typeof(AlizhouPicture))
                {
                    placeholderEntity.Placeholder = placeholder;
                    placeholderEntity.PlaceholderType = PlaceholderType.Picture;
                    var picture = (AlizhouPicture)prop.GetValue(wordData);
                    placeholderEntity.Pictures = new List<AlizhouPicture>() { picture };
                    placeholders.Add(placeholderEntity);
                }
                else if (typeof(IEnumerable<AlizhouPicture>).IsAssignableFrom(prop.PropertyType))
                {
                    placeholderEntity.Placeholder = placeholder;
                    placeholderEntity.PlaceholderType = PlaceholderType.Picture;
                    var pictures = (IEnumerable<AlizhouPicture>)prop.GetValue(wordData);
                    placeholderEntity.Pictures = pictures;
                    placeholders.Add(placeholderEntity);
                }
                else if (prop.PropertyType == typeof(AlizhouParagraph)) {
                    placeholderEntity.Placeholder = placeholder;
                    placeholderEntity.PlaceholderType = PlaceholderType.Paragraph;
                    placeholderEntity.Data = (AlizhouParagraph)prop.GetValue(wordData);
                    placeholders.Add(placeholderEntity);
                }
                else if (prop.PropertyType == typeof(AlizhouComplex))
                {
                    placeholderEntity.Placeholder = placeholder;
                    placeholderEntity.PlaceholderType = PlaceholderType.Complex;
                    placeholderEntity.Data = (AlizhouComplex)prop.GetValue(wordData);
                    placeholders.Add(placeholderEntity);
                }
            }
            return placeholders;
        }
    }
}
