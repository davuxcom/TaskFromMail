using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private const string REG_PATH = @"Software\DSASoftware\TaskFromMail";
        private const string CREATE_MENU_TEXT = "Create task...";
        private const string APPEND_MENU_TEXT = "Append to task...";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Attach a handler for when a context menu is opened on an item.
            Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Cleanup the handler
            Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay);
        }

        void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            // only show if one item is selected and it is a mail item.
            if (Selection.Count == 1 && GetMessageClass(Selection[1]) == "IPM.Note")
            {
                var button = (Office.CommandBarButton)CommandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                button.Caption = CREATE_MENU_TEXT;
                button.Visible = true;
                button.Click += new Office
                    ._CommandBarButtonEvents_ClickEventHandler(CreateTask_Click);

                button = (Office.CommandBarButton)CommandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                button.Caption = APPEND_MENU_TEXT;
                button.Visible = true;
                button.Click += new Office
                    ._CommandBarButtonEvents_ClickEventHandler(AppendTask_Click);
            }
        }

        private void CreateTask_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            SelectFolder(selected =>
                {
                    try
                    {
                        // FUTURE flag message with GREEN if posible
                        // Get the mail message in focus
                        var SelectedItem = Application.ActiveExplorer().Selection[1] as MailItem;
                        // create a new task _in_ the specified folder
                        var task = selected.Folder.Items.Add(OlItemType.olTaskItem) as TaskItem;
                        // space out the attachment a bit
                        task.Body = "\n\n\n";
                        // Attach the email message to the new task
                        task.Attachments.Add(SelectedItem);
                        // Give the task a default subject
                        task.Subject = SelectedItem.Subject;
                        // FUTURE: extract more data from the email to generate the task
                        task.Display(false);
                    }
                    catch (System.Exception ex)
                    {
                        // COMException for 'access deined' with reasonably friendly error message
                        MessageBox.Show(ex.Message, "TaskFromMail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
        }

        List<object> Pinned = new List<object>();

        private void AppendTask_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            SelectFolder(selected =>
            {
                // NOTE:  If we don't pin the COM object here, the RCW will free it when
                // we leave scope.  We'll unpin it after the task dialog comes back.
                Pinned.Add(selected);
                System.Action Unpin = () => Pinned.Remove(selected);

                try
                {
                    // Get the mail message in focus
                    var SelectedItem = Application.ActiveExplorer().Selection[1] as MailItem;
                    // request that the user choose a task to append it to
                    SelectTaskForm.AskUserForTask(selected.Folder.Items,
                        task =>
                        {
                            task.Attachments.Add(SelectedItem);
                            // FUTURE: extract more data from the email to generate the task
                            task.Display(false);
                            Unpin();
                        },
                        () =>
                        {
                            // user clicked cancel
                            Unpin();
                        });
                }
                catch (System.Exception ex)
                {
                    // COMException for 'access deined' with reasonably friendly error message
                    MessageBox.Show(ex.Message, "TaskFromMail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
        }

        private void SelectFolder(Action<NavigationFolder> Callback)
        {
            var exp = this.Application.ActiveExplorer();
            var nav = (TasksModule)exp.NavigationPane.Modules.GetNavigationModule(OlNavigationModuleType.olModuleTasks);

            NavigationFolder def = LoadDefaultFolder(nav.NavigationGroups);
            if (def != null)
            {
                // if CTRL is down, always display the selection dialog
                if (!IsKeyPushedDown(Keys.ControlKey))
                {
                    Callback(def);
                    return;
                }
            }

            // We don't have a default selection, so invoke the selection dialog
            SelectFolderForm.AskUserForFolder(nav.NavigationGroups,
                (folder, save) =>   // success
                {
                    try
                    {
                        if (save)
                        {
                            SaveDefaultFolder(nav.NavigationGroups, folder);
                        }
                        else
                        {
                            // NOTE: clear out the old folder, since this might be an
                            // override using CTRL, and we'll need to respect new settings
                            ClearDefaultFolder();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        // if we can't save, just log the error and continue
                        Debug.WriteLine(ex);
                    }
                    Callback(folder);
                },
                () =>       // cancelled
                {
                    // the user canceled the dialog
                }
            );
        }

        #region Keyboard Interop

        [DllImport("user32.dll")]
        static extern ushort GetAsyncKeyState(int vKey);

        public static bool IsKeyPushedDown(System.Windows.Forms.Keys vKey)
        {
            return 0 != (GetAsyncKeyState((int)vKey) & 0x8000);
        }

        #endregion

        #region Save & Load default values from registry

        private NavigationFolder LoadDefaultFolder(NavigationGroups nav)
        {
            try
            {
                RegistryKey rKey = Registry.CurrentUser.CreateSubKey(REG_PATH);
                string group = rKey.GetValue("DefaultGroup").ToString();
                string folder = rKey.GetValue("DefaultFolder").ToString();

                if (string.IsNullOrEmpty(group) || string.IsNullOrEmpty(folder))
                {
                    // nothing saved
                    return null;
                }

                Debug.WriteLine("Group: " + group + " Folder: " + folder);

                foreach (NavigationGroup g in nav)
                {
                    if (g.Name == group)
                    {
                        foreach (NavigationFolder f in g.NavigationFolders)
                        {
                            if (f.DisplayName == folder)
                            {
                                return f;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                // if we can't load... log the error and continue
                Debug.WriteLine(ex);
            }

            // if the default folder is missing, just invoke the dialog.
            return null;
        }

        private void ClearDefaultFolder()
        {
            RegistryKey rKey = Registry.CurrentUser.CreateSubKey(REG_PATH);
            rKey.SetValue("DefaultGroup", "");
            rKey.SetValue("DefaultFolder", "");
        }

        private void SaveDefaultFolder(NavigationGroups nav, NavigationFolder folder)
        {
            NavigationGroup group = null;
            foreach (NavigationGroup g in nav)
            {
                if (group == null)
                {
                    foreach (NavigationFolder f in g.NavigationFolders)
                    {
                        if (f == folder)
                        {
                            group = g;
                            break;
                        }
                    }
                }
            }

            // NOTE: this should be impossible, but don't crash.
            if (group == null) return;

            RegistryKey rKey = Registry.CurrentUser.CreateSubKey(REG_PATH);

            rKey.SetValue("DefaultGroup", group.Name);
            rKey.SetValue("DefaultFolder", folder.DisplayName);
        }

        #endregion

        private string GetMessageClass(object item)
        {
            // Use reflection to find out the message class
            object[] args = new Object[] { };
            Type t = item.GetType();
            return t.InvokeMember("messageClass",
                BindingFlags.Public | BindingFlags.GetField | BindingFlags.GetProperty,
                null, item, args).ToString();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
