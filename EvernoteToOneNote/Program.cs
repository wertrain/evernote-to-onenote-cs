using System;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;

namespace EvernoteToOneNote
{
    public class EvernoteToOneNoteBridge
    {
        /// <summary>
        /// 同じタイトルのノート名があった場合にノート名を変更
        /// （セクション名で同名は許可されていない）
        /// </summary>
        /// <param name="evernoteEnex"></param>
        public static void MakeNoteTitleUnique(EvernoteEnex evernoteEnex)
        {
            var titles = new List<string>();
            foreach (var note in evernoteEnex.Notes)
            {
                if (titles.Contains(note.Title))
                {
                    note.Title = note.Title + $"({note.Created})";
                }
                titles.Add(note.Title);
            }
        }

        /// <summary>
        /// ノート名を指定文字数にカットする
        /// </summary>
        /// <param name="evernoteEnex"></param>
        /// <param name="length"></param>
        public static void CutNoteTitle(EvernoteEnex evernoteEnex, int length)
        {
            foreach (var note in evernoteEnex.Notes)
            {
                if (note.Title.Length > length)
                {
                    note.Title = note.Title.Substring(0, length);
                }
            }
        }

        /// <summary>
        /// セクション名に使用できない文字を削除（と全角への置き換え）
        /// </summary>
        /// <param name="evernoteEnex"></param>
        public static void ReplaceTitleForbiddenCharacters(EvernoteEnex evernoteEnex)
        {
            foreach (var note in evernoteEnex.Notes)
            {
                var title = ReplaceForbiddenCharacters(
                    note.Title.
                        Replace("!", "！").
                        Replace("?", "？").
                        Replace("&", "＆").
                        Replace("%", "％").
                        Replace(":", "：")
                    );

                note.Title = title.Trim();
            }
        }

        /// <summary>
        /// Evernote から OneNote 用の本文に変換する
        /// </summary>
        /// <param name="evernoteEnex"></param>
        /// <returns></returns>
        public static IEnumerable<OneNoteWrapper.PageParameter> ConvertContent(EvernoteEnex evernoteEnex)
        {
            var hashList = new List<string>();

            foreach (var note in evernoteEnex.Notes)
            {
                var content = ReplaceSpecialCharacters(note.Content);

                var root = XElement.Parse(content);
                foreach (var element in root.Descendants())
                {
                    if (element.Name == "en-media")
                    {
                        // フォーマットのチェック
                        switch (element.Attribute("type")?.Value)
                        {
                            case "image/jpeg":
                            case "image/png":
                            case "image/gif":
                                break;
                            default:
                                continue;
                        }
                        var hash = element.Attribute("hash");
                        if (hash == null) continue;
                        hashList.Add(hash.Value);

                        // エレメント内容の置き換え
                        element.SetAttributeValue("src", $"name:{hash.Value}");
                        element.Name = "img";
                        hash.Remove();
                    }
                }

                var param = new OneNoteWrapper.PageParameter()
                {
                    Title = note.Title,
                    Content = StripNoteTag(root.ToString()),
                    Created = note.Created,
                    Url = note.Attribute.SourceUrl
                };
                foreach (var resource in note.Resources)
                {
                    string findHash = null;
                    foreach (var hash in hashList)
                    {
                        if (resource.Attribute.SourceUrl.Contains(hash))
                        {
                            findHash = hash;
                            break;
                        }
                    }
                    if (string.IsNullOrEmpty(findHash)) continue;
                    
                    var attachment = new OneNoteWrapper.PageParameter.AttachmentParameter();
                    attachment.Name = findHash;
                    attachment.Width = resource.Width;
                    attachment.Height = resource.Height;
                    attachment.FilePath = resource.FilePath;
                    attachment.ContentType = resource.Mime;
                    param.Attachments.Add(attachment);
                }

                yield return param;
            }
        }

        /// <summary>
        /// 特殊文字を削除
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static string ReplaceSpecialCharacters(string text)
        {
            return text.Replace("&nbsp;", " ");
        }

        /// <summary>
        /// OneNote の禁止文字を削除
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string ReplaceForbiddenCharacters(string text)
        {
            var separators = new char[] { '~', '#', '%', '&', '*', '{', '}', '|', '\\', ':', '\"', '<', '>', '?', '/', '^'};
            string[] temp = text.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", temp);
        }

        /// <summary>
        /// "en-note" タグを取り除く
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static string StripNoteTag(string text)
        {
            string targetStartTag = "<en-note>";
            if (text.IndexOf(targetStartTag) == 0)
            {
                string targetLastTag = "</en-note>";
                var strip = text.Remove(0, targetStartTag.Length).Trim();
                
                var lastIndex = strip.Length - targetLastTag.Length;
                if (strip.LastIndexOf(targetLastTag) == lastIndex)
                {
                    return strip.Substring(0, lastIndex);
                }
            }
            return text;
        }
    }

    class Program
    {
        static int Main(string[] args)
        {
            var accessToken = args[0];
            if (string.IsNullOrWhiteSpace(accessToken))
            {
                Console.WriteLine("Access token is required for argument 0.");
                return 1;
            }

            var enexFilePath = args[1];
            if (System.IO.File.Exists(enexFilePath))
            {
                Console.WriteLine("Evernote export file (*.enex) path must be specified as argument 1.");
                return 1;
            }

            var enex = new EvernoteEnex();
            enex.Load(enexFilePath);
            try
            {
                EvernoteToOneNoteBridge.MakeNoteTitleUnique(enex);
                EvernoteToOneNoteBridge.ReplaceTitleForbiddenCharacters(enex);
                EvernoteToOneNoteBridge.CutNoteTitle(enex, 49);

                OneNoteWrapper oneNote = new OneNoteWrapper(accessToken);
                var noteBookId = oneNote.CreateNotebook(enex.Name);
                foreach (var param in EvernoteToOneNoteBridge.ConvertContent(enex))
                {
                    var sectionId = oneNote.CreateSection(noteBookId, param.Title);
                    oneNote.CreatePage(sectionId, param);
                }
            }
            catch
            {
                return 2;
            }
            finally
            {
                enex.ClearTempResourceFiles();
            }

            return 0;
        }
    }
}
