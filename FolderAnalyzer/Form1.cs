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

namespace FolderAnalyzer
{
    public partial class Form1 : Form
    {
        private Timer m_timer = new Timer();
        private int m_interval = 1000; // ms
        private bool m_bLogging = true; // logging or not

        private int m_elapsed = 0;

        private string m_curpaths; // currently opened paths in explorer windows

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            m_timer.Tick += new EventHandler(MyCallback);
            m_timer.Interval = m_interval;
            m_timer.Enabled = true;

            listView1.View = View.Details;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;

            ColumnHeader chpath = new ColumnHeader();
            chpath.Text = "Path";
            chpath.Width = 500;
            ColumnHeader chval = new ColumnHeader();
            chval.Text = "Value";
            chval.Width = 100;
            ColumnHeader[] charray = { chpath, chval };
            listView1.Columns.AddRange(charray);

        }

        public void MyCallback(object sender, EventArgs e)
        {
            m_elapsed += m_interval;
            label2.Text = m_elapsed.ToString();

            label1.Text = "";

            string paths = "";

            Shell shell = new Shell();
            ShellWindows win = shell.Windows();
            foreach (IWebBrowser2 web in win)
            {
                if (Path.GetFileName(web.FullName).ToUpper() == "EXPLORER.EXE")
                {
                    string str = web.LocationURL;
                    ListViewItem obj = listView1.FindItemWithText(str);
                    if (obj != null)
                    {
                        label1.Text = "obj=" + obj.Text;
                    }
                    else
                    {
                        string val = "1";
                        string[] newitem = { str, val };
                        listView1.Items.Add(new ListViewItem(newitem));

                    }
                    int idx = comboBox1.Items.IndexOf(str);
                    if (idx < 0)
                    {
                        comboBox1.Items.Insert(0, str);
                        paths += str;
                    }
                }
            }
            if (!paths.Equals(m_curpaths))
            {
                // update
                label1.Text += "Updated";
                m_curpaths = paths;
            }
            else
            {
                label1.Text += "Skip";

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Text = "";

            Shell shell = new Shell();
            ShellWindows win = shell.Windows();
            foreach (IWebBrowser2 web in win)
            {
                if (Path.GetFileName(web.FullName).ToUpper() == "EXPLORER.EXE")
                {
                    label1.Text += web.LocationURL + " : " + web.LocationName + "\n";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("EXPLORER.EXE", comboBox1.Text);
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //ListViewItem target = (ListViewItem)sender;
            string seltext = "";
            //seltext = target.SubItems[0].Text;
            seltext = listView1.SelectedItems[0].SubItems[0].Text;
            int val = int.Parse(listView1.SelectedItems[0].SubItems[1].Text);
            val++;
            listView1.SelectedItems[0].SubItems[1].Text = val.ToString();
            System.Diagnostics.Process.Start("EXPLORER.EXE", seltext);
        }
    }
}
