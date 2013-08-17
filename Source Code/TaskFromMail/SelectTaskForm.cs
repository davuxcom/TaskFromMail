using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using app = System.Windows.Forms.Application;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace OutlookAddIn1
{
    // NOTE:  Using WinForms instead of WPF because the perf hit to load up
    // the first WPF Window is significant, and may confuse the user. :(
    public partial class SelectTaskForm : Form
    {
        public static void AskUserForTask(Outlook.Items items,
            Action<Outlook.TaskItem> SuccessCallback, System.Action CancelCallback)
        {
            var t = new Thread(() =>
            {
                SelectTaskForm frm = new SelectTaskForm(items);
                frm.ShowDialog();
                // TODO: get the hwnd for this instance of outlook
                // NOTE: looks like this is not possible, we could easily
                // attach to the wrong explorer in the right process
                // and end up focusing a window poorly and confusing the user
                app.Run();  // NOTE: Run() will block until Exit()
                if (frm.DialogResult == DialogResult.OK)
                {
                    SuccessCallback(frm.Result);
                }
                else
                {
                    CancelCallback();
                }
            });
            t.SetApartmentState(ApartmentState.STA);    // Required for UI
            t.Start();

            // TODO: fix this.  scope issue frees the COM object from the RCW
            t.Join();
        }

        // store the users selection
        private Outlook.TaskItem Result { get; set; }

        public SelectTaskForm(Outlook.Items items)
        {
            InitializeComponent();

            ColumnHeader col = new ColumnHeader();
            col.Width = lsv.Width - lsv.Margin.Left - lsv.Margin.Right;
            lsv.Columns.Add(col);
            lsv.View = View.Details;

            List<ListViewGroup> Groups = new List<ListViewGroup>();

            foreach (Outlook.TaskItem item in items)
            {
                try
                {
                    if (item.Status == Outlook.OlTaskStatus.olTaskComplete) continue;

                    string[] categories = new string[] { "Uncategorized" };

                    if (!string.IsNullOrEmpty(item.Categories))
                    {
                        categories = item.Categories.Split(new string[] { ",", ";" }, StringSplitOptions.RemoveEmptyEntries);
                    }

                    foreach (var catE in categories)
                    {
                        var cat = catE.Trim();
                        var group = Groups.Where(g => g.Header == cat).FirstOrDefault();

                        if (group == default(ListViewGroup))
                        {
                            Groups.Add(new ListViewGroup(cat));
                        }
                    }

                    lsv.Items.Add(new ListViewItem
                    {
                        Text = item.Subject,
                        Group = Groups.Where(g => g.Header == categories[0]).FirstOrDefault(),
                        Tag = item
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            try
            {
                Groups.Sort((a, b) => a.Header.CompareTo(b.Header));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            lsv.Groups.AddRange(Groups.ToArray());

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
            Resize += (_, __) => lsv.Columns[0].Width = lsv.Width - lsv.Margin.Left - lsv.Margin.Right;
        }

        private void btn_Click(object sender, EventArgs e)
        {
            if (lsv.SelectedItems.Count == 1)
            {
                Result = lsv.SelectedItems[0].Tag as Outlook.TaskItem;
                DialogResult = System.Windows.Forms.DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Select a task for the message to be apended to",
                    "TaskFromMail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
