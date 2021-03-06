using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;

namespace EvernoteToOneNote
{
    /// <summary>
    /// OneNote の Graph API のラッパークラス
    /// エラー処理がないので注意
    /// </summary>
    public class OneNoteWrapper
    {
        /// <summary>
        /// アクセストークン
        /// </summary>
        public string AccessToken { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="accessToken"></param>
        public OneNoteWrapper(string accessToken)
        {
            AccessToken = accessToken;
        }

        /// <summary>
        /// Notebook を取得
        /// </summary>
        /// <returns>Json フォーマットの文字列</returns>
        public string GetNotebooks()
        {
            string url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks";

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers["Authorization"] = $"Bearer {AccessToken}";

            using (var response = request.GetResponse())
            {
                using(var stream = response.GetResponseStream())
                using(var reader = new System.IO.StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// ノートブックを作成
        /// </summary>
        /// <param name="name">作成するノートブック名</param>
        /// <returns>作成したノートブックの ID</returns>
        public string CreateNotebook(string name)
        {
            string url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks";

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers["Authorization"] = $"Bearer {AccessToken}";
            request.Headers["Content-type"] = $"application/json";
            request.Method = "POST";

            using (var stream = request.GetRequestStream())
            using (var writer = new System.IO.StreamWriter(stream))
            {
                var bodys = new Dictionary<string, object>
                {
                    { "displayName", name }
                };   
                writer.Write(JsonSerializer.Serialize(bodys).ToString());
            }
   
            using (var response = request.GetResponse())
            {
                using (var stream = response.GetResponseStream())
                using (var reader = new System.IO.StreamReader(stream))
                {
                    var dictionary = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(reader.ReadToEnd());
                    return dictionary["id"].ToString();
                }
            }
        }

        /// <summary>
        /// セクションを作成
        /// </summary>
        /// <param name="notebookId">セクションを追加するノートブックの ID</param>
        /// <param name="name">作成するセクション名</param>
        /// <returns></returns>
        public string CreateSection(string notebookId, string name)
        {
            string url = $"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebookId}/sections";

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers["Authorization"] = $"Bearer {AccessToken}";
            request.Headers["Content-type"] = $"application/json";
            request.Method = "POST";

            using (var stream = request.GetRequestStream())
            using (var writer = new System.IO.StreamWriter(stream))
            {
                var bodys = new Dictionary<string, object>
                {
                    { "displayName", name }
                };
                writer.Write(JsonSerializer.Serialize(bodys).ToString());
            }

            using (var response = request.GetResponse())
            {
                using (var stream = response.GetResponseStream())
                using (var reader = new System.IO.StreamReader(stream))
                {
                    var dictionary = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(reader.ReadToEnd());
                    return dictionary["id"].ToString();
                }
            }
        }

        /// <summary>
        /// ページ作成パラメータ
        /// </summary>
        public class PageParameter
        {
            /// <summary>
            /// タイトル
            /// </summary>
            public string Title { get; set; }

            /// <summary>
            /// コンテンツ
            /// </summary>
            public string Content { get; set; }

            /// <summary>
            /// ページ作成時間
            /// </summary>
            public DateTime? Created { get; set; }

            /// <summary>
            /// 添付ファイルが存在するかどうか
            /// </summary>
            public bool HasAttachment { get { return Attachments?.Count > 0; } }

            /// <summary>
            /// 添付ファイル
            /// </summary>
            public class AttachmentParameter
            {
                /// <summary>
                /// 添付ファイルのタイプ
                /// </summary>
                public string ContentType { get; set; }

                /// <summary>
                /// 添付ファイル名
                /// </summary>
                public string Name { get; set; }

                /// <summary>
                /// 添付ファイルまでのパス
                /// </summary>
                public string FilePath { get; set; }

                /// <summary>
                /// 添付ファイルの幅
                /// </summary>
                public int Width { get; set; }

                /// <summary>
                /// 添付ファイルの高さ
                /// </summary>
                public int Height { get; set; }
            }

            /// <summary>
            /// 添付ファイルのリスト
            /// </summary>
            public List<AttachmentParameter> Attachments { get; set; } = new List<AttachmentParameter>();

            /// <summary>
            /// ユーザー情報
            /// </summary>
            public Object Tag { get; set; }

            /// <summary>
            /// ユーザー情報を保持しているか
            /// </summary>
            public bool HasTag { get { return Tag != null; } }

            /// <summary>
            /// URL 情報
            /// </summary>
            public string Url { get; set; }
        }

        /// <summary>
        /// ページを作成する
        /// </summary>
        /// <param name="id">ノートブックのID</param>
        /// <param name="title">ページのタイトル</param>
        /// <param name="body">ページの本文</param>
        /// <returns></returns>
        public string CreatePage(string sectionId, PageParameter param)
        {
            string url = $"https://graph.microsoft.com/v1.0/me/onenote/sections/{sectionId}/pages";

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers["Authorization"] = $"Bearer {AccessToken}";
            request.Headers["Content-type"] = $"application/xhtml+xml";
            request.Method = "POST";

            string boundary = string.Format("{0:N}", Guid.NewGuid());

            using (var stream = request.GetRequestStream())
            using (var writer = new System.IO.StreamWriter(stream))
            {
                if (param.HasAttachment)
                {
                    request.Headers["Content-type"] = $"multipart/form-data; boundary={boundary}";

                    writer.WriteLine($"--{boundary}");
                    writer.WriteLine($"Content-Disposition:form-data; name=\"Presentation\"");
                    writer.WriteLine($"Content-Type:text/html");
                    writer.WriteLine();
                    writer.Flush();
                }
                WritePageParameter(writer, param, boundary);
            }

            using (var response = request.GetResponse())
            {
                using (var stream = response.GetResponseStream())
                using (var reader = new System.IO.StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// PageParameter を Stream に出力する
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="param"></param>
        /// <param name="boundary"></param>
        private void WritePageParameter(System.IO.StreamWriter writer, PageParameter param, string boundary)
        {
            writer.WriteLine($"<!DOCTYPE html>");
            writer.WriteLine($"<html>");
            writer.WriteLine($"  <head>");
            if (!string.IsNullOrWhiteSpace(param.Title))
            {
                writer.WriteLine($"    <title>{param.Title}</title>");
            }
            if (param.Created.HasValue)
            {
                var dateTime = param.Created.Value;
                writer.WriteLine($"    <meta name=\"created\" content=\"{dateTime.ToString("yyyy-MM-ddTHH:mm:sszzzz")}\" />");
            }
            writer.WriteLine($"  </head>");
            writer.WriteLine($"  <body>");
            if (!string.IsNullOrWhiteSpace(param.Url))
            {
                writer.WriteLine($"    <blockquote>{param.Url}</blockquote>");
            }
            writer.WriteLine($"    {param.Content}");
            writer.WriteLine($"  </body>");
            writer.WriteLine($"</html>");
            writer.Flush();

            if (param.HasAttachment)
            {
                foreach (var attachment in param.Attachments)
                {
                    writer.WriteLine();
                    writer.WriteLine($"--{boundary}");

                    writer.WriteLine($"Content-Disposition:form-data; name=\"{attachment.Name}\"");
                    writer.WriteLine($"Content-Type:{attachment.ContentType}");
                    writer.WriteLine();
                    writer.Flush();
                    using (var stream = new System.IO.FileStream(attachment.FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    using (var reader = new System.IO.BinaryReader(stream))
                    {
                        var bytes = new byte[stream.Length];
                        stream.Read(bytes, 0, bytes.Length);
                        writer.BaseStream.Write(bytes, 0, bytes.Length);
                        writer.Flush();
                    }
                }

                writer.WriteLine();
                writer.WriteLine($"--{boundary}--");
                writer.Flush();
            }
        }
    }
}
