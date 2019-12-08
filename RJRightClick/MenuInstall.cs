using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;

namespace RJRightClick
{
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
