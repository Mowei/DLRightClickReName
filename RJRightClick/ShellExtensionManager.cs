using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RJRightClick
{
    class ShellExtensionManager
    {
        private static string clsid = "{6E49F0EC-FC7F-41D9-B1CE-3553E9FBBCE5}";
        private static string fileType = "*";
        private static string friendlyName = "RJ Class";
        private static string keyName = $@"Software\Classes\{fileType}\shellex\ContextMenuHandlers\{clsid}";

        public static void RegisterShellExtContextMenuHandler(bool allUsers)
        {
            var rootName = allUsers ? Registry.LocalMachine : Registry.CurrentUser;

            using (var key = rootName.CreateSubKey(keyName))
            {
                key?.SetValue(null, friendlyName);
            }
        }

        public static void UnregisterShellExtContextMenuHandler(bool allUsers)
        {
            var rootName = allUsers ? Registry.LocalMachine : Registry.CurrentUser;

            rootName.DeleteSubKeyTree(keyName);
        }

        public static bool IsInstalled(bool allUsers)
        {
            var rootName = allUsers ? Registry.LocalMachine : Registry.CurrentUser;

            return rootName.OpenSubKey(keyName, false) != null;
        }
    }
}