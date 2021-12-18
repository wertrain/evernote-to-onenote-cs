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
        /// ページを作成する
        /// </summary>
        /// <param name="id">ノートブックのID</param>
        /// <param name="title">ページのタイトル</param>
        /// <param name="body">ページの本文</param>
        /// <returns></returns>
        public string CreatePage(string id, string title, string body)
        {
            string url = $"https://graph.microsoft.com/v1.0/me/onenote/sections/{id}/pages";

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers["Authorization"] = $"Bearer {AccessToken}";
            request.Headers["Content-type"] = $"application/xhtml+xml";
            request.Method = "POST";

            using (var stream = request.GetRequestStream())
            using (var writer = new System.IO.StreamWriter(stream))
            {
                string xhtml = "" +
                    $"<!DOCTYPE html>" +
                    $"<html>" +
                    $"  <head>" +
                    $"    <title>A page with a block of HTML</title>" +
                    $"    <meta name=\"created\" content=\"2015-07-22T09:00:00-08:00\" />" +
                    $"  </head>" +
                    $"  <body>{body}</body>" +
                    $"</html>";
                writer.Write(xhtml);
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
    }
}
