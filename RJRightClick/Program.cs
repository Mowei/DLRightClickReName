using CommandLine;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RJRightClick
{
    class Program
    {
        public class Options
        {
            [Option("Install", Required = false, HelpText = "Install Menu")]
            public bool Install { get; set; }

            [Option("UnInstall", Required = false, HelpText = "UnInstall Menu")]
            public bool UnInstall { get; set; }

            [Option('s', "site", Required = false, HelpText = "Go to DLSite")]
            public IEnumerable<string> Site { get; set; }

            [Option('r', "rename", Required = false, HelpText = "rename file")]
            public IEnumerable<string> Rename { get; set; }

            [Option('i', "image", Required = false, HelpText = "get image")]
            public IEnumerable<string> Image { get; set; }

        }

        static void Main(string[] args)
        {

            Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(o =>
            {
                if (o.Install)
                {
                    var installTool = new MenuInstall();
                    installTool.DoRegister(false);

                    installTool.NotifyShell();

                    MessageBox.Show("Install finished.",
                        "RJ Context Menu Installer",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                if (o.UnInstall)
                {
                    var installTool = new MenuInstall();
                    installTool.DoUnRegister(false);

                    installTool.NotifyShell();

                    MessageBox.Show("Uninstall finished.",
                        "RJ Context Menu Installer",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                if (o.Site.Count() > 0)
                {
                    foreach (var filePath in o.Site)
                    {
                        Regex rgx = new Regex("(RJ\\d{6})", RegexOptions.IgnoreCase);
                        var result = rgx.Match(filePath);
                        if (result.Success)
                        {
                            var RJNumber = result.Value;
                            System.Diagnostics.Process.Start($@"https://www.dlsite.com/home/work/=/product_id/{RJNumber}.html");
                        }
                    }
                }

                if (o.Rename.Count() > 0)
                {
                    var NameFormatTemplate = ConfigurationManager.AppSettings["NameFormatTemplate"];
                    var WorkNameXPath = ConfigurationManager.AppSettings["WorkNameXPath"];
                    var MakerNameXPath = ConfigurationManager.AppSettings["MakerNameXPath"];
                    var SaleDateXPath = ConfigurationManager.AppSettings["SaleDateXPath"];
                    var WorkGenreXPath = ConfigurationManager.AppSettings["WorkGenreXPath"];

                    foreach (var filePath in o.Rename)
                    {
                        Regex rgx = new Regex("(RJ\\d{6})", RegexOptions.IgnoreCase);
                        var result = rgx.Match(filePath);
                        if (result.Success)
                        {
                            //RJ號
                            var RJNumber = result.Value;

                            var htmlWeb = new HtmlWeb();
                            var query = $@"https://www.dlsite.com/home/work/=/product_id/{result.Value}.html";
                            var doc = htmlWeb.Load(query);

                            //名稱
                            var work_name = doc.DocumentNode.SelectSingleNode(WorkNameXPath).InnerText;
                            //社團
                            var maker_name = doc.DocumentNode.SelectSingleNode(MakerNameXPath).InnerText;
                            //販售日
                            var sale_Date = doc.DocumentNode.SelectSingleNode(SaleDateXPath).InnerText.Replace("年", "").Replace("月", "").Replace("日", "").Substring(2);
                            //作品形式
                            var work_genre = string.Join("", doc.DocumentNode.SelectNodes(WorkGenreXPath).Select(x => $"({x.InnerText})"));

                            var customTypes = "";
                            var genreTypes = ConfigurationManager.GetSection("genreTypes") as NameValueCollection;
                            if (genreTypes.Count >= 0)
                            {
                                foreach (var key in genreTypes.AllKeys)
                                {
                                    var genType = doc.DocumentNode.SelectSingleNode(genreTypes[key]);
                                    if (genType != null)
                                    {
                                        customTypes += $@"({genType.InnerText})";
                                    }
                                }
                            }

                            var name = NameFormatTemplate
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

                            try
                            {
                                var fileinfo = new FileInfo(filePath);
                                var Extensions = fileinfo.Name.Split('.').ToList();
                                Extensions.RemoveAt(0);
                                var newFilename = name + "." + string.Join(".", Extensions);
                                if (newFilename.Length >= 200) return;
                                fileinfo.MoveTo(Path.Combine(fileinfo.Directory.FullName, newFilename));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"命名失敗: {ex.InnerException.Message}", "失敗");
                                return;
                            }

                        }
                    }
                    MessageBox.Show("命名成功!");
                }
                if (o.Image.Count() > 0)
                {
                    var ProductSampleImagesXPath = ConfigurationManager.AppSettings["ProductSampleImagesXPath"];

                    foreach (var filePath in o.Image)
                    {
                        var targetDir = Path.GetDirectoryName(filePath);

                        Regex rgx = new Regex("(RJ\\d{6})", RegexOptions.IgnoreCase);
                        var result = rgx.Match(filePath);
                        if (result.Success)
                        {
                            //RJ號
                            var RJNumber = result.Value;

                            var htmlWeb = new HtmlWeb();
                            var query = $@"https://www.dlsite.com/home/work/=/product_id/{result.Value}.html";
                            var doc = htmlWeb.Load(query);
                            var results = doc.DocumentNode.SelectNodes(ProductSampleImagesXPath);

                            if (results != null)
                            {
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
            });

        }

    }
}
