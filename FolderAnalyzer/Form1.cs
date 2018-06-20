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

        private Dictionary<string, int> m_dict = new Dictionary<string, int>();
        private Dictionary<string, bool> m_curPaths = new Dictionary<string, bool>();

        private string setting_filename = "folder_log.txt";
        
        public Form1()
        {
            InitializeComponent();
        }

        private void ListView1_initialize()
        {
            listView1.Clear();
            listView1.View = View.Details;
            listView1.Columns.Add("Path", 500);
            listView1.Columns.Add("Value", 100);
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            m_timer.Tick += new EventHandler(MyCallback);
            m_timer.Interval = m_interval;
            m_timer.Enabled = true;

            LoadData(Application.UserAppDataPath + setting_filename);

            ListView1_initialize();

            m_curPaths.Clear();
        }

        public void MyCallback(object sender, EventArgs e)
        {
            m_elapsed += m_interval;
            label2.Text = m_elapsed.ToString();

            string paths = "";

            List<string> keylist = new List<string>(m_curPaths.Keys);

            foreach (string key in keylist)
            {
                m_curPaths[key] = false; // once disable
            }
            
            Shell shell = new Shell();
            ShellWindows win = shell.Windows();
            foreach (IWebBrowser2 web in win)
            {
                if (Path.GetFileName(web.FullName).ToUpper() == "EXPLORER.EXE")
                {
                    string str = web.LocationURL;
                    if (m_dict.ContainsKey(str))
                    {
                        if (!m_curPaths.ContainsKey(str))
                        {
                            m_dict[str] += 1;
                        }
                    }
                    else
                    {
                        m_dict[str] = 1;
                    }
                    m_curPaths[str] = true;
                    paths += str;
                }
            }
            if (!paths.Equals(m_curpaths))
            {
                // update
                label1.Text = "Updated";
                m_curpaths = paths;

                ListView1_initialize();
                IOrderedEnumerable<KeyValuePair<string, int>> sorted = m_dict.OrderByDescending(pair => pair.Value);
                foreach (KeyValuePair<string, int> pair in sorted)
                {
                    string[] newitem = { pair.Key, pair.Value.ToString() };
                    listView1.Items.Add(new ListViewItem(newitem));
                }

                foreach (string key in keylist)
                {
                    if (m_curPaths[key] == false)
                    {
                        m_curPaths.Remove(key);
                    }
                }
            }
            else
            {
                label1.Text = "Skip";
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveData(Application.UserAppDataPath + setting_filename);
            label1.Text = Application.UserAppDataPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenExplorer();
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenExplorer();
        }

        private void OpenExplorer()
        {
            string path = listView1.SelectedItems[0].SubItems[0].Text;
            int val = int.Parse(listView1.SelectedItems[0].SubItems[1].Text);
            val++;
            m_dict[path]++;
            listView1.SelectedItems[0].SubItems[1].Text = val.ToString();
            System.Diagnostics.Process.Start("EXPLORER.EXE", path);
        }

        private void SaveData(string filename)
        {
            StreamWriter sw = new StreamWriter(filename, false, Encoding.GetEncoding("Shift_JIS"));
            IOrderedEnumerable<KeyValuePair<string, int>> sorted = m_dict.OrderByDescending(pair => pair.Value);
            foreach (KeyValuePair<string, int> pair in sorted/*m_dict*/)
            {
                if (pair.Key.Length > 0)
                {
                    sw.WriteLine(pair.Key + "\t" + pair.Value/* + sw.NewLine*/);
                }
            }
            sw.Close();
        }

        private void LoadData(string filename)
        {
            string line = "";
            StreamReader sr;

            m_dict.Clear();
            ListView1_initialize();

            try
            {
                sr = new StreamReader(filename, Encoding.GetEncoding("Shift_JIS"));
            }
            catch (System.IO.FileNotFoundException ex)
            {
                return;
            }

            while ((line = sr.ReadLine()) != null)
            {
                string[] delimiter = { "\t" };
                string[] sttmp = line.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);
                string key = sttmp[0];
                int value = int.Parse(sttmp[1]);
                if (key.Length > 0)
                {
                    m_dict[key] = value;

                }
            }
            sr.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveData(Application.UserAppDataPath + setting_filename);
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
