using CommandLine;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace RJRename.CLI
{
    class Program
    {
        public class Options
        {
            public bool UnInstall { get; set; }

            [Option('s', "site", Required = false, HelpText = "Go to DLSite")]
            public IEnumerable<string> Site { get; set; }

            [Option('g', "GetNewname", Required = false, HelpText = "Get Newname")]
            public IEnumerable<string> Newname { get; set; }

            [Option('r', "rename", Required = false, HelpText = "Rename file")]
            public IEnumerable<string> Rename { get; set; }

            [Option('i', "image", Required = false, HelpText = "Get image")]
            public IEnumerable<string> Image { get; set; }

        }
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(o =>
            {
                if (o.Site.Count() > 0)
                {
                    var rjUtil = new RJRename.Core.Util();
                    foreach (var filePath in o.Site)
                    {
                        var RJNumber = rjUtil.GetRJNumber(filePath);
                        if (!string.IsNullOrEmpty(RJNumber))
                        {
                            ProcessStartInfo startInfo = new ProcessStartInfo($@"https://www.dlsite.com/home/work/=/product_id/{RJNumber}.html")
                            {
                                UseShellExecute = true
                            };
                            Process.Start(startInfo);
                        }
                    }
                }

                if (o.Newname.Count() > 0)
                {
                    var rjUtil = new RJRename.Core.Util();
                    foreach (var str in o.Newname)
                    {
                        var newName = rjUtil.GetRJNewName(str);
                        if (!string.IsNullOrEmpty(newName))
                        {
                            Console.WriteLine(newName);
                        }
                        else
                        {
                            Console.WriteLine("Not Found!");
                        }      
                    }
                }

                if (o.Rename.Count() > 0)
                {
                    var rjUtil = new RJRename.Core.Util();

                    foreach (var filePath in o.Rename)
                    {
                        var RJNumber = rjUtil.GetRJNumber(filePath);
                        if (!string.IsNullOrEmpty(RJNumber))
                        {
                            try
                            {
                                rjUtil.ReNameFile(RJNumber, filePath);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"命名失敗: {ex.InnerException.Message}", "失敗");
                                return;
                            }

                        }
                    }
                    Console.WriteLine("命名成功!");
                }
                if (o.Image.Count() > 0)
                {
                    var rjUtil = new RJRename.Core.Util();

                    foreach (var filePath in o.Image)
                    {
                        var RJNumber = rjUtil.GetRJNumber(filePath);
                        if (!string.IsNullOrEmpty(RJNumber))
                        {
                            rjUtil.DownLoadImg(RJNumber, filePath);
                        }
                    }
                }
            });
        }
    }
}
