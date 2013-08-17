using System;
using System.Threading;
using System.Windows.Forms;
using app = System.Windows.Forms.Application;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    // NOTE:  Using WinForms instead of WPF because the perf hit to load up
    // the first WPF Window is significant, and may confuse the user. :(
    public partial class SelectFolderForm : Form
    {
        public static void AskUserForFolder(Outlook.NavigationGroups nav, 
            Action<Outlook.NavigationFolder, bool> SuccessCallback, 
            System.Action CancelCallback)
        {
            var t = new Thread(() =>
                {
                    SelectFolderForm frm = new SelectFolderForm(nav);
                    frm.ShowDialog();
                    // TODO: get the hwnd for this instance of outlook
                    // NOTE: looks like this is not possible, we could easily
                    // attach to the wrong explorer in the right process
                    // and end up focusing a window poorly and confusing the user
                    app.Run();  // NOTE: Run() will block until Exit()
                    if (frm.DialogResult == DialogResult.OK)
                    {
                        SuccessCallback(frm.Result, frm.chk.Checked);
                    }
                    else
                    {
                        CancelCallback();
                    }
                });
            t.SetApartmentState(ApartmentState.STA);    // Required for UI
            t.Start();
        }

        public SelectFolderForm(Outlook.NavigationGroups groups)
        {
            InitializeComponent();

            ColumnHeader col = new ColumnHeader();
            col.Width = lsv.Width - lsv.Margin.Left - lsv.Margin.Right;
            lsv.Columns.Add(col);
            lsv.View = View.Details;

            foreach (Outlook.NavigationGroup group in groups)
            {
                ListViewGroup lvg = new ListViewGroup(group.Name);
                lsv.Groups.Add(lvg);
                foreach (Outlook.NavigationFolder folder in group.NavigationFolders)
                {
                    lsv.Items.Add(new ListViewItem
                    {
                        Text = folder.DisplayName,
                        Group = lvg,
                        Tag = folder
                    });
                }
            }

            btn.Enabled = false;    // nothing is selected yet
            lsv.SelectedIndexChanged += (_, __) =>
                {
                    btn.Enabled = lsv.SelectedItems.Count > 0;
                };
            lsv.DoubleClick += (s, e) =>
                {
                    if (btn.Enabled) btn_Click(null, null);
                };

            FormClosing += (_, __) => app.Exit();
        }

        private void btn_Click(object sender, EventArgs e)
        {
            if (lsv.SelectedItems.Count == 1)
            {
                Result = lsv.SelectedItems[0].Tag as Outlook.NavigationFolder;
                DialogResult = System.Windows.Forms.DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Select a folder for the new task to be created in",
                    "TaskFromMail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        // store the users selection
        private Outlook.NavigationFolder Result { get; set; }
    }
}
