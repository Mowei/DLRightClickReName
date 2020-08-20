using CommandLine;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

            [Option('g', "GetNewname", Required = false, HelpText = "Get Newname")]
            public IEnumerable<string> Newname { get; set; }

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
                    var rjUtil = new RJRename.Core.Util();
                    foreach (var filePath in o.Site)
                    {
                        var RJNumber = rjUtil.GetRJNumber(filePath);
                        if (!string.IsNullOrEmpty(RJNumber))
                        {
                            System.Diagnostics.Process.Start($@"https://www.dlsite.com/home/work/=/product_id/{RJNumber}.html");
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
                            MessageBox.Show(newName);
                        }
                        else
                        {
                            MessageBox.Show("Not Found!");
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
                                MessageBox.Show($"命名失敗: {ex.InnerException.Message}", "失敗");
                                return;
                            }

                        }
                    }
                    MessageBox.Show("命名成功!");
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
