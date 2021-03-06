using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace EvernoteToOneNote
{
    /// <summary>
    /// Evernote の .enex ファイルを表すクラス
    /// </summary>
    public class EvernoteEnex
    {
        /// <summary>
        /// ノートを表すクラス
        /// </summary>
        public class Note
        {
            public string Title { get; set; }
            public string Content { get; set; }
            public DateTime Created { get; set; }
            public DateTime Updated { get; set; }
            public List<string> Tags { get; set; } = new List<string>();

            /// <summary>
            /// 属性
            /// </summary>
            public class Attributes
            {
                public float Latitude { get; set; }
                public float Longitude { get; set; }
                public float Altitude { get; set; }
                public string Author { get; set; }
                public string SourceUrl { get; set; }
            }

            /// <summary>
            /// 属性値
            /// </summary>
            public Attributes Attribute { get; set; }

            /// <summary>
            /// 添付リソース
            /// </summary>
            public class Resource
            {
                public string FilePath { get; set; }
                public string Mime { get; set; }

                public int Width { get; set; }
                public int Height { get; set; }

                public class Attributes
                {
                    public string FileName { get; set; }
                    public string SourceUrl { get; set; }
                }

                public Attributes Attribute { get; set; }
            }

            /// <summary>
            /// 添付リソース
            /// </summary>
            public List<Resource> Resources = new List<Resource>();
        }

        /// <summary>
        /// ノート名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// .enex に含まれるノート
        /// </summary>
        public List<Note> Notes { get; } = new List<Note>();
        
        /// <summary>
        /// .enex ファイルをロードする
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool Load(string filePath)
        {
            Name = Path.GetFileNameWithoutExtension(filePath);

            var root = XElement.Load(filePath);
            foreach (var noteElement in root.Descendants("note"))
            {
                var note = new Note();
                var title = noteElement.Element("title")?.Value;
                var content = noteElement.Element("content")?.Value;
                var created = noteElement.Element("created")?.Value;
                var updated = noteElement.Element("updated")?.Value;
                var tags = noteElement.Element("tag")?.Value.Split(",");

                note.Title = title;
                note.Content = content;
                note.Created = DateTime.ParseExact(created, "yyyyMMddTHHmmssZ", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None);
                note.Updated = DateTime.ParseExact(updated, "yyyyMMddTHHmmssZ", System.Globalization.DateTimeFormatInfo.InvariantInfo, System.Globalization.DateTimeStyles.None);
                note.Tags.AddRange(tags ?? new string[0]);

                var attributes = new Note.Attributes();
                if (noteElement.Element("note-attributes") != null)
                {
                    var attributesElement = noteElement.Element("note-attributes");

                    if (float.TryParse(attributesElement.Element("latitude")?.Value, out var latitude))
                    {
                        attributes.Latitude = latitude;
                    }

                    if (float.TryParse(attributesElement.Element("longitude")?.Value, out var longitude))
                    {
                        attributes.Longitude = longitude;
                    }

                    if (float.TryParse(attributesElement.Element("altitude")?.Value, out var altitude))
                    {
                        attributes.Altitude = altitude;
                    }

                    attributes.Author = attributesElement.Element("author")?.Value;
                    attributes.SourceUrl = attributesElement.Element("source-url")?.Value;
                    note.Attribute = attributes;
                }

                foreach (var resourceElement in noteElement.Descendants("resource"))
                {
                    var resource = new Note.Resource();
                    resource.Mime = resourceElement.Element("mime")?.Value;
                    
                    if (int.TryParse(resourceElement.Element("width")?.Value, out var width))
                    {
                        resource.Width = width;
                    }

                    if (int.TryParse(resourceElement.Element("height")?.Value, out var height))
                    {
                        resource.Height = height;
                    }

                    if (resourceElement.Element("resource-attributes") != null)
                    {
                        var resourceAttributesElement = resourceElement.Element("resource-attributes");
                        resource.Attribute = new Note.Resource.Attributes();
                        resource.Attribute.FileName = resourceAttributesElement.Element("file-name")?.Value;
                        resource.Attribute.SourceUrl = resourceAttributesElement.Element("source-url")?.Value;
                    }

                    var data = resourceElement.Element("data")?.Value;
                    if (resourceElement.Element("data").Attribute("encoding").Value == "base64")
                    {
                        resource.FilePath = Path.GetTempFileName();
                        using (var stream = new FileStream(resource.FilePath, FileMode.Create))
                        using (var writer = new BinaryWriter(stream))
                        {
                            writer.Write(Convert.FromBase64String(data));
                        }
                        //File.Delete(resource.FilePath);
                    }
                    note.Resources.Add(resource);
                }
                Notes.Add(note);
            }
            return true;
        }

        /// <summary>
        /// 一時ファイルをクリア
        /// </summary>
        public void ClearTempResourceFiles()
        {
            foreach (var note in Notes)
            {
                foreach (var resource in note.Resources)
                {
                    File.Delete(resource.FilePath);
                }
            }
        }
    }
}
