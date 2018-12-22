using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// http://tinqwill.blog59.fc2.com/blog-entry-84.html
using Shell32;
using SHDocVw;
using System.IO;
using System.Collections;

using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace FolderAnalyzer
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void ListView1_initialize()
        {
            int colwidth0 = 1000;// -2;// 500;
            int colwidth1 = 100;
            int colwidth2 = 20;
            if (listView1.Columns.Count > 0)
            {
                colwidth0 = listView1.Columns[0].Width;
                colwidth1 = listView1.Columns[1].Width;
                colwidth2 = listView1.Columns[2].Width;
            }
            listView1.Clear();
            listView1.View = View.Details;
            listView1.Columns.Add("Path", colwidth0);
            listView1.Columns.Add("Value", colwidth1);
            listView1.Columns.Add("Match", colwidth2);
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            /*
            lvsort = new ListViewSort();
            lvsort.ColumnModes = new ListViewSort.ComparerMode[]
            {
                ListViewSort.ComparerMode.String,
                ListViewSort.ComparerMode.Integer,
                ListViewSort.ComparerMode.String
            };
            listView1.ListViewItemSorter = lvsort;
            */
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenExplorer();
        }

        const int SWF_NOSIZE = 0x0001;
        const int SWF_NOMOVE = 0x0002;
        const int SWF_SHOWWINDOW = 0x0040;
        const int HWND_TOPMOST = -1;
        const int HWND_NOTOPMOST = -2;
        
        [DllImport("user32.dll")]
        //[DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

        private void OpenExplorer()
        {
            if (listView1.SelectedItems.Count > 0)
            {
                IntPtr hWnd = (IntPtr)int.Parse(listView1.SelectedItems[0].SubItems[1].Text);
                SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWF_NOMOVE | SWF_NOSIZE);
                SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWF_SHOWWINDOW | SWF_NOMOVE | SWF_NOSIZE);
            }
        }

        public void EnumerateWindows()
        {
            ListView1_initialize();

            listView1.BeginUpdate();

            Shell shell = new Shell();
            ShellWindows win = shell.Windows();
            foreach (IWebBrowser2 web in win)
            {
                try
                {
                    if (Path.GetFileName(web.FullName).ToUpper() == "EXPLORER.EXE")
                    {
                        string str = web.LocationURL;
                        if (str.Length == 0) { continue; } // LocationURL is empty for special folders (ex. Documents)

                        IntPtr hWnd = (IntPtr)web.HWND;

                        // Add
                        string[] newitem = { str, hWnd.ToString(), "" };
                        listView1.Items.Add(new ListViewItem(newitem));
                    }
                }
                catch (Exception ex)
                {
                    // Skip this window if something is wrong
                }
            }
            listView1.EndUpdate();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            EnumerateWindows();
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                OpenExplorer();
            }
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            EnumerateWindows();
        }
    }
}
