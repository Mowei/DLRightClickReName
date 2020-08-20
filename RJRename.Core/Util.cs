using HtmlAgilityPack;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace RJRename.Core
{
    public class Util
    {
        public AppSetting AppSetting { get; set; } = new AppSetting();

        public Util()
        {
            var builder = new ConfigurationBuilder()
              .SetBasePath(Directory.GetCurrentDirectory())
              .AddJsonFile("appsettings.json");
            var config = builder.Build();
            config.Bind(AppSetting);
        }
        public string GetRJNumber(string rjStr)
        {
            Regex rgx = new Regex("(RJ\\d{6})", RegexOptions.IgnoreCase);
            var result = rgx.Match(rjStr);
            if (result.Success)
            {
                return result.Value;
            }
            return null;
        }
        public string GetRJNewName(string rjStr)
        {
            //RJ號
            var RJNumber = GetRJNumber(rjStr);
            if (!string.IsNullOrEmpty(RJNumber))
            {
                var htmlWeb = new HtmlWeb();
                var query = $@"https://www.dlsite.com/home/work/=/product_id/{RJNumber}.html";
                var doc = htmlWeb.Load(query);

                //名稱
                var work_name = doc.DocumentNode.SelectSingleNode(AppSetting.WorkNameXPath).InnerText.Trim();
                //社團
                var maker_name = doc.DocumentNode.SelectSingleNode(AppSetting.MakerNameXPath).InnerText.Trim();
                //販售日
                var sale_Date = doc.DocumentNode.SelectSingleNode(AppSetting.SaleDateXPath).InnerText.Trim().Replace("年", "").Replace("月", "").Replace("日", "").Substring(2);
                //作品形式
                var work_genre = string.Join("", doc.DocumentNode.SelectNodes(AppSetting.WorkGenreXPath).Select(x => $"({x.InnerText.Trim()})"));

                var customTypes = "";
                var genreTypes = AppSetting.GenreTypes;
                if (genreTypes.Count >= 0)
                {
                    foreach (var genType in genreTypes)
                    {
                        var genreVal = doc.DocumentNode.SelectSingleNode(genType.Value);
                        if (genreVal != null)
                        {
                            customTypes += $@"({genreVal.InnerText.Trim()})";
                        }
                    }
                }

                var name = AppSetting.NameFormatTemplate;
                var customXPath = AppSetting.CustomXPath;
                if (customXPath.Count >= 0)
                {
                    foreach (var custom in customXPath)
                    {
                        var customVal = doc.DocumentNode.SelectSingleNode(custom.Value);
                        if (customVal != null)
                        {
                            name = name.Replace(custom.Key, customVal.InnerText.Trim());
                        }
                    }
                }
                name = name
                .Replace("%maker_name%", maker_name)
                .Replace("%sale_date%", sale_Date)
                .Replace("%number%", RJNumber)
                .Replace("%work_name%", work_name)
                .Replace("%work_genre%", work_genre)
                .Replace("%custom_types%", customTypes)
                .Replace("?", "？")
                .Replace("~", "～")
                .Replace("*", "＊")
                .Replace("/", "／")
                .Replace("\\", "＼")
                .Replace(":", "：")
                .Replace("\"", "＂")
                .Replace("<", "＜")
                .Replace(">", "＞")
                .Replace("|", "｜");

                return name;
            }
            return null;
        }

        public void ReNameFile(string rjStr, string filePath)
        {

            var RJNewName = GetRJNewName(rjStr);
            if (!string.IsNullOrEmpty(RJNewName))
            {
                try
                {
                    var fileinfo = new FileInfo(filePath);
                    var Extensions = fileinfo.Name.Split('.').ToList();
                    Extensions.RemoveAt(0);
                    var newFilename = RJNewName + "." + string.Join(".", Extensions);
                    if (newFilename.Length >= 200) return;
                    fileinfo.MoveTo(Path.Combine(fileinfo.Directory.FullName, newFilename));
                }
                catch (Exception ex)
                {
                    throw;
                }

            }
        }

        public void DownLoadImg(string rjStr, string dirPath)
        {
            //RJ號
            var RJNumber = GetRJNumber(rjStr);
            if (!string.IsNullOrEmpty(RJNumber))
            {
                var htmlWeb = new HtmlWeb();
                var query = $@"https://www.dlsite.com/home/work/=/product_id/{RJNumber}.html";
                var doc = htmlWeb.Load(query);
                var results = doc.DocumentNode.SelectNodes(AppSetting.ProductSampleImagesXPath);

                if (results != null)
                {
                    var targetDir = Path.GetDirectoryName(dirPath);
                    foreach (var node in results)
                    {
                        var imageUrl = "https:" + node.Attributes["data-src"].Value;
                        Uri uri = new Uri(imageUrl);
                        string filename = Path.GetFileName(uri.LocalPath);

                        using (WebClient client = new WebClient())
                        {
                            client.DownloadFile(imageUrl, Path.Combine(targetDir, filename));
                        }
                    }
                }

            }
        }
    }
}
