/********************************** Module Header **********************************\
Module Name:  FileContextMenuExt.cs
Project:      CSShellExtContextMenuHandler
Copyright (c) Microsoft Corporation.

The FileContextMenuExt.cs file defines a context menu handler by implementing the 
IShellExtInit and IContextMenu interfaces.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***********************************************************************************/

#region Using directives

using CSShellExtContextMenuHandler.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

#endregion

namespace CSShellExtContextMenuHandler
{
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6E49F0EC-FC7F-41D9-B1CE-3553E9FBBCE5"), ComVisible(true)]
    public class FileContextMenuExt : IShellExtInit, IContextMenu
    {
        // The name of the selected file.
        private readonly List<MenuItem> menuItems = new List<MenuItem>();
        private List<string> SelectedFile = new List<string>();
        private IntPtr menuBmp = IntPtr.Zero;

        public FileContextMenuExt()
        {
            // Load the bitmap for the menu item.
            Bitmap bmp = Resources.OK;
            bmp.MakeTransparent(bmp.GetPixel(0, 0));
            this.menuBmp = bmp.GetHbitmap();

            //Load default items.
            menuItems.Add(new MenuItem("瀏覽DLSite網站", true, null, "--site %FILE_PATH%"));
            menuItems.Add(new MenuItem("DLSite管理選單", true, null, ""));
            menuItems.Add(new MenuItem("檔案重新命名", true, null, "--rename %FILE_PATH%"));
            menuItems.Add(new MenuItem("下載sample圖片", true, null, "--image %FILE_PATH%"));
            menuItems.Add(new MenuItem("測試命名功能", true, null, "--GetNewname %FILE_PATH%"));
        }

        #region IShellExtInit Members

        /// <summary>
        ///     Initialize the context menu handler.
        /// </summary>
        /// <param name="pidlFolder">
        ///     A pointer to an ITEMIDLIST structure that uniquely identifies a folder.
        /// </param>
        /// <param name="pDataObj">
        ///     A pointer to an IDataObject interface object that can be used to retrieve
        ///     the objects being acted upon.
        /// </param>
        /// <param name="hKeyProgId">
        ///     The registry key for the file object or folder type.
        /// </param>
        public void Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgId)
        {
            if (pDataObj == IntPtr.Zero)
            {
                throw new ArgumentException();
            }

            var fe = new FORMATETC
            {
                cfFormat = (short)CLIPFORMAT.CF_HDROP,
                ptd = IntPtr.Zero,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                tymed = TYMED.TYMED_HGLOBAL
            };
            STGMEDIUM stm;

            // The pDataObj pointer contains the objects being acted upon. In this 
            // example, we get an HDROP handle for enumerating the selected files 
            // and folders.
            var dataObject = (IDataObject)Marshal.GetObjectForIUnknown(pDataObj);
            dataObject.GetData(ref fe, out stm);

            try
            {
                // Get an HDROP handle.
                var hDrop = stm.unionmember;
                if (hDrop == IntPtr.Zero)
                {
                    throw new ArgumentException();
                }

                // Determine how many files are involved in this operation.
                var nFiles = NativeMethods.DragQueryFile(hDrop, uint.MaxValue, null, 0);

                // This code sample displays the custom context menu item when only 
                // one file is selected. 
                if (nFiles >= 1)
                {
                    for (uint i = 0; i < nFiles; i++)
                    {
                        // Get the path of the file.
                        var fileName = new StringBuilder(260);
                        if (0 == NativeMethods.DragQueryFile(hDrop,
                                                             i,
                                                             fileName,
                                                             fileName.Capacity))
                        {
                            Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                        }

                        Regex rgx = new Regex("(RJ\\d{6})", RegexOptions.IgnoreCase);
                        var result = rgx.Match(fileName.ToString());
                        if (!result.Success)
                        {
                            Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                        }
                        SelectedFile.Add(fileName.ToString());
                    }
                }
                else
                {
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }
            finally
            {
                NativeMethods.ReleaseStgMedium(ref stm);
            }
        }

        #endregion


        private void OnVerbDisplayFileName(string cmd)
        {
            string cmdstring = string.Join(" ", SelectedFile.Select(x => $"\"{x}\""));
            Process.Start(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "RJRightClick.exe"),
                          cmd.Replace("%FILE_PATH%", cmdstring));
        }

        private int RegisterMenuItem(uint id,
                                     uint idCmdFirst,
                                     string text,
                                     bool enabled,
                                     IntPtr bitmap,
                                     IntPtr subMenu,
                                     uint position,
                                     IntPtr registerTo)
        {
            var sub = new MENUITEMINFO();
            sub.cbSize = (uint)Marshal.SizeOf(sub);

            var m = MIIM.MIIM_STRING | MIIM.MIIM_FTYPE | MIIM.MIIM_ID | MIIM.MIIM_STATE;
            if (bitmap != IntPtr.Zero)
                m |= MIIM.MIIM_BITMAP;
            if (subMenu != IntPtr.Zero)
                m |= MIIM.MIIM_SUBMENU;
            sub.fMask = m;

            sub.wID = idCmdFirst + id;
            sub.fType = MFT.MFT_STRING;
            sub.dwTypeData = text;
            sub.hSubMenu = subMenu;
            sub.fState = enabled ? MFS.MFS_ENABLED : MFS.MFS_DISABLED;
            sub.hbmpItem = bitmap;

            if (!NativeMethods.InsertMenuItem(registerTo, position, true, ref sub))
                return Marshal.GetHRForLastWin32Error();
            return 0;
        }

        #region Shell Extension Registration

        [ComRegisterFunction]
        public static void Register(Type t)
        {
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
        }

        #endregion

        #region IContextMenu Members

        /// <summary>
        ///     Add commands to a shortcut menu.
        /// </summary>
        /// <param name="hMenu">A handle to the shortcut menu.</param>
        /// <param name="iMenu">
        ///     The zero-based position at which to insert the first new menu item.
        /// </param>
        /// <param name="idCmdFirst">
        ///     The minimum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="idCmdLast">
        ///     The maximum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="uFlags">
        ///     Optional flags that specify how the shortcut menu can be changed.
        /// </param>
        /// <returns>
        ///     If successful, returns an HRESULT value that has its severity value set
        ///     to SEVERITY_SUCCESS and its code value set to the offset of the largest
        ///     command identifier that was assigned, plus one.
        /// </returns>
        public int QueryContextMenu(
            IntPtr hMenu,
            uint iMenu,
            uint idCmdFirst,
            uint idCmdLast,
            uint uFlags)
        {
            // If uFlags include CMF_DEFAULTONLY then we should not do anything.
            if (((uint)CMF.CMF_DEFAULTONLY & uFlags) != 0)
            {
                return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, 0);
            }

            // Add a separator.
            var sep = new MENUITEMINFO();
            sep.cbSize = (uint)Marshal.SizeOf(sep);
            sep.fMask = MIIM.MIIM_TYPE;
            sep.fType = MFT.MFT_SEPARATOR;
            if (!NativeMethods.InsertMenuItem(hMenu, 0, true, ref sep))
                return Marshal.GetHRForLastWin32Error();

            // Register item
            var item = menuItems[0];
            RegisterMenuItem(0, idCmdFirst, item.Text, true, menuBmp, IntPtr.Zero, 1, hMenu);

            // Register PopupMenu
            var hSubMenu = NativeMethods.CreatePopupMenu();
            item = menuItems[1];
            RegisterMenuItem(1, idCmdFirst, item.Text, true, IntPtr.Zero, hSubMenu, 2, hMenu);


            // Register item
            item = menuItems[2];
            RegisterMenuItem(2, idCmdFirst, item.Text, true, IntPtr.Zero, IntPtr.Zero, 0, hSubMenu);

            // Register item
            item = menuItems[3];
            RegisterMenuItem(3, idCmdFirst, item.Text, true, IntPtr.Zero, IntPtr.Zero, 1, hSubMenu);

            // Register item
            item = menuItems[4];
            RegisterMenuItem(4, idCmdFirst, item.Text, true, IntPtr.Zero, IntPtr.Zero, 2, hSubMenu);

            // Add a separator.
            sep = new MENUITEMINFO();
            sep.cbSize = (uint)Marshal.SizeOf(sep);
            sep.fMask = MIIM.MIIM_TYPE;
            sep.fType = MFT.MFT_SEPARATOR;
            NativeMethods.InsertMenuItem(hMenu, 3, true, ref sep);

            // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
            // Set the code value to the total number of items added.
            return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, 3 + (uint)menuItems.Count);
        }

        /// <summary>
        ///     Carry out the command associated with a shortcut menu item.
        /// </summary>
        /// <param name="pici">
        ///     A pointer to a CMINVOKECOMMANDINFO or CMINVOKECOMMANDINFOEX structure
        ///     containing information about the command.
        /// </param>
        public void InvokeCommand(IntPtr pici)
        {
            var ici = (CMINVOKECOMMANDINFO)Marshal.PtrToStructure(pici, typeof(CMINVOKECOMMANDINFO));

            var item = menuItems[NativeMethods.LowWord(ici.verb.ToInt32())];

            OnVerbDisplayFileName(item.Commands);
        }

        /// <summary>
        ///     Get information about a shortcut menu command, including the help string
        ///     and the language-independent, or canonical, name for the command.
        /// </summary>
        /// <param name="idCmd">Menu command identifier offset.</param>
        /// <param name="uFlags">
        ///     Flags specifying the information to return. This parameter can have one
        ///     of the following values: GCS_HELPTEXTA, GCS_HELPTEXTW, GCS_VALIDATEA,
        ///     GCS_VALIDATEW, GCS_VERBA, GCS_VERBW.
        /// </param>
        /// <param name="pReserved">Reserved. Must be IntPtr.Zero</param>
        /// <param name="pszName">
        ///     The address of the buffer to receive the null-terminated string being
        ///     retrieved.
        /// </param>
        /// <param name="cchMax">
        ///     Size of the buffer, in characters, to receive the null-terminated string.
        /// </param>
        public void GetCommandString(
            UIntPtr idCmd,
            uint uFlags,
            IntPtr pReserved,
            StringBuilder pszName,
            uint cchMax)
        {
        }

        #endregion
    }
}