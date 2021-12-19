using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.Json;

namespace EvernoteToOneNote
{
    /// <summary>
    /// OneNote の Graph API のラッパークラス
    /// エラー処理がないので注意
    /// </summary>
    class OneNoteApi
    {
        /// <summary>
        /// アクセストークン
        /// </summary>
        public string AccessToken { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="accessToken"></param>
        public OneNoteApi(string accessToken)
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
            public string Title { get; set; }
            public string Body { get; set; }
            public DateTime? DateTime { get; set; }
            public bool HasAttachment { get; set; }
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

            string boundary = string.Format("---{0:N}", Guid.NewGuid());
            if (param.HasAttachment)
            {
                request.Headers["Content-type"] = $"multipart/form-data; boundary={boundary}";
            }

            using (var stream = request.GetRequestStream())
            using (var writer = new System.IO.StreamWriter(stream))
            {
                writer.Write(PageParameterToXHtml(param));
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

        private string PageParameterToXHtml(PageParameter param, string boundary)
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"<!DOCTYPE html>");
            stringBuilder.AppendLine($"<html>");
            stringBuilder.AppendLine($"  <head>");
            if (!string.IsNullOrWhiteSpace(param.Title))
            {
                stringBuilder.AppendLine($"    <title>{param.Title}</title>");
            }
            if (param.DateTime.HasValue)
            {
                var dateTime = param.DateTime.Value;
                stringBuilder.AppendLine($"    <meta name=\"created\" content=\"{dateTime.ToString("yyyy-MM-ddTHH:mm:sszzzz")}\" />");
            }
            stringBuilder.AppendLine($"  </head>");
            stringBuilder.AppendLine($"  <body>{param.Body}</body>");
            stringBuilder.AppendLine($"</html>");
            return stringBuilder.ToString();
        }
    }
}
