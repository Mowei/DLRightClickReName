using System;

namespace CSShellExtContextMenuHandler
{
    internal struct MenuItem
    {
        internal string Commands;
        internal bool? ShowInMainMenu;
        internal string Text;
        internal bool Enabled;

        internal MenuItem(string text, bool enabled, bool? showInMainMenu, string commands)
        {
            ShowInMainMenu = showInMainMenu;
            Text = text;
            Enabled = enabled;
            Commands = commands;
        }
    }
}