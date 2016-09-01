using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;

namespace mainUI
{
    public class MruLoader
    {
        // The application's name.
        private string ApplicationName;

        // A list of the files.
        private int NumFiles;
        private List<FileInfo> FileInfos;

        // The File menu.
        private RibbonOrbDropDown MyMenu;
        private RibbonOrbRecentItem[] MenuItems;

        // Raised when the user selects a file from the MRU list.
        public delegate void FileSelectedEventHandler(string file_name);
        public event FileSelectedEventHandler FileSelected;

        // Constructor.
        public MruLoader(string application_name, RibbonOrbDropDown menu, int num_files)
        {
            ApplicationName = application_name;
            MyMenu = menu;
            NumFiles = num_files;
            FileInfos = new List<FileInfo>();
            MenuItems = new RibbonOrbRecentItem[NumFiles + 1];

            for (int i = 0; i < NumFiles; i++)
            {
                MenuItems[i] = new RibbonOrbRecentItem();
                MenuItems[i].Visible = false;
                MyMenu.RecentItems.Add(MenuItems[i]);
            }

            // Reload items from the registry.
            LoadFiles();

            // Display the items.
            ShowFiles();
        }

        // Load saved items from the Registry.
        private void LoadFiles()
        {
            // Reload items from the registry.
            for (int i = 0; i < NumFiles; i++)
            {
                string file_name = (string)RegistryTools.GetSetting(
                    ApplicationName, "FilePath" + i.ToString(), "");
                if (file_name != "")
                {
                    FileInfos.Add(new FileInfo(file_name));
                }
            }
        }

        // Save the current items in the Registry.
        private void SaveFiles()
        {
            // Delete the saved entries.
            for (int i = 0; i < NumFiles; i++)
            {
                RegistryTools.DeleteSetting(ApplicationName, "FilePath" + i.ToString());
            }

            // Save the current entries.
            int index = 0;
            foreach (FileInfo file_info in FileInfos)
            {
                RegistryTools.SaveSetting(ApplicationName,
                    "FilePath" + index.ToString(), file_info.FullName);
                index++;
            }
        }

        // Remove a file's info from the list.
        private void RemoveFileInfo(string file_name)
        {
            // Remove occurrences of the file's information from the list.
            for (int i = FileInfos.Count - 1; i >= 0; i--)
            {
                if (FileInfos[i].FullName == file_name) FileInfos.RemoveAt(i);
            }
        }

        // Add a file to the list, rearranging if necessary.
        public void AddFile(string file_name)
        {
            // Remove the file from the list.
            RemoveFileInfo(file_name);

            // Add the file to the beginning of the list.
            FileInfos.Insert(0, new FileInfo(file_name));

            // If we have too many items, remove the last one.
            if (FileInfos.Count > NumFiles) FileInfos.RemoveAt(NumFiles);

            // Display the files.
            ShowFiles();

            // Update the Registry.
            SaveFiles();
        }

        // Remove a file from the list, rearranging if necessary.
        public void RemoveFile(string file_name)
        {
            // Remove the file from the list.
            RemoveFileInfo(file_name);

            // Display the files.
            ShowFiles();

            // Update the Registry.
            SaveFiles();
        }

        // Display the files in the menu items.
        private void ShowFiles()
        {
            for (int i = 0; i < FileInfos.Count; i++)
            {
                MenuItems[i].Text = string.Format("{0} {1}", i + 1, FileInfos[i].Name);
                MenuItems[i].Visible = true;
                MenuItems[i].Tag = FileInfos[i];
                MenuItems[i].Click -= File_Click;
                MenuItems[i].Click += File_Click;
                MenuItems[i].ToolTip = FileInfos[i].FullName;
            }
            for (int i = FileInfos.Count; i < NumFiles; i++)
            {
                MenuItems[i].Visible = false;
                MenuItems[i].Click -= File_Click;
            }
        }

        // The user selected a file from the menu.
        private void File_Click(object sender, EventArgs e)
        {
            // Don't bother if no one wants to catch the event.
            if (FileSelected != null)
            {
                // Get the corresponding FileInfo object.
                RibbonOrbRecentItem menu_item = sender as RibbonOrbRecentItem;
                FileInfo file_info = menu_item.Tag as FileInfo;

                // Raise the event.
                FileSelected(file_info.FullName);
            }
        }
    }

    public class RegistryTools
    {
        // Save a value.
        public static void SaveSetting(string app_name, string name, object value)
        {
            RegistryKey reg_key = Registry.CurrentUser.OpenSubKey("Software", true);
            RegistryKey sub_key = reg_key.CreateSubKey(app_name);
            sub_key.SetValue(name, value);
        }

        // Get a value.
        public static object GetSetting(string app_name, string name, object default_value)
        {
            RegistryKey reg_key = Registry.CurrentUser.OpenSubKey("Software", true);
            RegistryKey sub_key = reg_key.CreateSubKey(app_name);
            return sub_key.GetValue(name, default_value);
        }

        // Delete a value.
        public static void DeleteSetting(string app_name, string name)
        {
            RegistryKey reg_key = Registry.CurrentUser.OpenSubKey("Software", true);
            RegistryKey sub_key = reg_key.CreateSubKey(app_name);
            try
            {
                if(sub_key.GetValue(name) != null)
                    sub_key.DeleteValue(name);
            }
            catch
            {
            }
        }
    }
}
