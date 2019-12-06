using CommandLine;
using HtmlAgilityPack;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
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
            public string Site { get; set; }

            [Option('r', "rename", Required = false, HelpText = "rename file")]
            public string Rename { get; set; }

            [Option('i', "image", Required = false, HelpText = "get image")]
            public string Image { get; set; }

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

                if (!string.IsNullOrWhiteSpace(o.Site))
                {
                    Regex rgx = new Regex("(RJ\\d{6})", RegexOptions.IgnoreCase);
                    var result = rgx.Match(o.Site);
                    if (result.Success)
                        System.Diagnostics.Process.Start($@"https://www.dlsite.com/home/work/=/product_id/{result.Value}.html");
                }
                if (!string.IsNullOrWhiteSpace(o.Rename))
                {

                    Regex rgx = new Regex("(RJ\\d{6})", RegexOptions.IgnoreCase);
                    var result = rgx.Match(o.Rename);
                    if (result.Success)
                    {
                        //RJ號
                        var RJNumber = result.Value;

                        var htmlWeb = new HtmlWeb();
                        var query = $@"https://www.dlsite.com/home/work/=/product_id/{result.Value}.html";
                        var doc = htmlWeb.Load(query);
                        var response = doc.DocumentNode.SelectSingleNode("//div[@id='work_right_inner']");

                        //名稱
                        var work_name = response.SelectSingleNode("//h1[@id='work_name']//a").InnerText;

                        //社團
                        var maker_name = response.SelectSingleNode("//span[@class='maker_name']").InnerText;


                        var results = response.SelectNodes("//table[@id='work_outline']//td");
                        if (results != null)
                        {
                            //販售日
                            var saleDate = results[0].InnerText.Replace("年", "").Replace("月", "").Replace("日", "").Substring(2);
                            //作品形式
                            var rjtype = string.Join("", results[3].SelectNodes("//div[@class='work_genre']//span").Select(x => $"({x.InnerText})"));

                            var name = $"[{maker_name}][{saleDate}][{RJNumber}]{work_name}{rjtype}"
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
                                var fileinfo = new FileInfo(o.Rename);
                                var Extensions = fileinfo.Name.Split('.').ToList();
                                Extensions.RemoveAt(0);
                                var newFilename = name + "." + string.Join(".", Extensions);
                                if (newFilename.Length >= 200) return;
                                fileinfo.MoveTo(Path.Combine(fileinfo.Directory.FullName, newFilename));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"命名失敗: {ex.InnerException.Message}", "失敗");
                            }
                            MessageBox.Show("命名成功!");
                        }


                    }
                }
                if (!string.IsNullOrWhiteSpace(o.Image))
                {

                }

            });

        }


        public class MenuInstall
        {
            private readonly string crtDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            public void DoRegister(bool allUsers)
            {
                try
                {
                    if (!allUsers)
                        OverrideHKCR();

                    var rs = new RegistrationServices();
                    rs.RegisterAssembly(Assembly.LoadFrom(Path.Combine(crtDir, @"CSShellExtContextMenuHandler.dll")),
                        AssemblyRegistrationFlags.SetCodeBase);

                    ShellExtensionManager.RegisterShellExtContextMenuHandler(allUsers);

                    if (!allUsers)
                        OverrideHKCR(true);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message + "\r\n\r\n" + e.StackTrace);
                }
            }

            public void DoUnRegister(bool allUsers)
            {
                try
                {
                    if (!allUsers)
                        OverrideHKCR();

                    var rs = new RegistrationServices();
                    rs.UnregisterAssembly(Assembly.LoadFrom(Path.Combine(crtDir, @"CSShellExtContextMenuHandler.dll")));

                    ShellExtensionManager.UnregisterShellExtContextMenuHandler(allUsers);

                    if (!allUsers)
                        OverrideHKCR(true);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message + "\r\n\r\n" + e.StackTrace);
                }
            }

            private void OverrideHKCR(bool restore = false)
            {

                UIntPtr HKEY_CLASSES_ROOT = Is64BitOS() ? new UIntPtr(0xFFFFFFFF80000000) : new UIntPtr(0x80000000);
                UIntPtr HKEY_CURRENT_USER = Is64BitOS() ? new UIntPtr(0xFFFFFFFF80000001) : new UIntPtr(0x80000001);

                // 0xF003F = KEY_ALL_ACCESS
                UIntPtr key = UIntPtr.Zero;

                RegOpenKeyEx(HKEY_CURRENT_USER, @"Software\Classes", 0, 0xF003F, out key);
                RegOverridePredefKey(HKEY_CLASSES_ROOT, restore ? UIntPtr.Zero : key);
            }


            public void NotifyShell()
            {
                const uint SHCNE_ASSOCCHANGED = 0x08000000;
                const ushort SHCNF_IDLIST = 0x0000;

                SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
            }

            public static bool IsAdministrator()
            {
                var wp = new WindowsPrincipal(WindowsIdentity.GetCurrent());

                return wp.IsInRole(WindowsBuiltInRole.Administrator);
            }

            private static bool Is64BitOS()
            {
                //The code below is from http://1code.codeplex.com/SourceControl/changeset/view/39074#842775
                //which is under the Microsoft Public License: http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.

                if (UIntPtr.Size == 8) // 64-bit programs run only on Win64
                {
                    return true;
                }
                // Detect whether the current process is a 32-bit process 
                // running on a 64-bit system.
                bool flag;
                return DoesWin32MethodExist("kernel32.dll", "IsWow64Process") &&
                       IsWow64Process(GetCurrentProcess(), out flag) && flag;
            }

            private static bool DoesWin32MethodExist(string moduleName, string methodName)
            {
                var moduleHandle = GetModuleHandle(moduleName);
                if (moduleHandle == UIntPtr.Zero)
                {
                    return false;
                }
                return GetProcAddress(moduleHandle, methodName) != UIntPtr.Zero;
            }

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern int DeleteFile(string name);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern void SetLastError(int errorCode);

            [DllImport("shell32.dll", SetLastError = true)]
            private static extern void SHChangeNotify(uint wEventId, ushort uFlags, IntPtr dwItem1, IntPtr dwItem2);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern UIntPtr GetCurrentProcess();

            [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            private static extern UIntPtr GetModuleHandle(string moduleName);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern UIntPtr GetProcAddress(UIntPtr hModule, [MarshalAs(UnmanagedType.LPStr)] string procName);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool IsWow64Process(UIntPtr hProcess, out bool wow64Process);

            [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
            private static extern int RegOpenKeyEx(UIntPtr hKey, string subKey, int ulOptions, uint samDesired,
                out UIntPtr hkResult);

            [DllImport("advapi32.dll", SetLastError = true)]
            private static extern int RegOverridePredefKey(UIntPtr hKey, UIntPtr hNewKey);

            [DllImport("advapi32.dll", SetLastError = true)]
            public static extern int RegCloseKey(UIntPtr hKey);

            [DllImport("user32.dll", SetLastError = true)]
            static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern UInt32 SendMessage(IntPtr hWnd, UInt32 msg, UInt32 wParam, UInt32 lParam);

            [DllImport("user32.dll")]
            private static extern bool SetProcessDPIAware();
        }
    }
}
