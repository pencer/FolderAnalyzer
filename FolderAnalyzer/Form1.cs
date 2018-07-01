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

        private ListViewSort lvsort;

        public Form1()
        {
            InitializeComponent();
        }

        private void ListView1_initialize()
        {
            int colwidth0 = 500;
            int colwidth1 = 100;
            if (listView1.Columns.Count > 0)
            {
                colwidth0 = listView1.Columns[0].Width;
                colwidth1 = listView1.Columns[1].Width;
            }   
            listView1.Clear();
            listView1.View = View.Details;
            listView1.Columns.Add("Path", colwidth0);
            listView1.Columns.Add("Value", colwidth1);
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            lvsort = new ListViewSort();
            lvsort.ColumnModes = new ListViewSort.ComparerMode[]
            {
                ListViewSort.ComparerMode.String,
                ListViewSort.ComparerMode.Integer
            };
            listView1.ListViewItemSorter = lvsort;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            m_timer.Tick += new EventHandler(MyCallback);
            m_timer.Interval = m_interval;
            m_timer.Enabled = true;

            LoadData(Application.UserAppDataPath + setting_filename);

            ListView1_initialize();

            m_curPaths.Clear();

            // Receive children's key event.
            this.KeyPreview = true;
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

            bool newitemfound = false;
            
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
                            // an already registered folder is opened
                            m_dict[str] += 1;
                            ListViewItem obj = listView1.FindItemWithText(str);
                            if (obj != null)
                            {
                                obj.SubItems[1].Text = (int.Parse(obj.SubItems[1].Text) + 1).ToString();
                            }
                        }
                    }
                    else
                    {
                        // not registered folder
                        newitemfound = true;
                        m_dict[str] = 1;
                    }
                    m_curPaths[str] = true;
                    paths += str;
                }
            }
            if (!paths.Equals(m_curpaths))
            {
                if (newitemfound || listView1.Items.Count == 0)
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
            if (listView1.SelectedItems.Count > 0)
            {
                string path = listView1.SelectedItems[0].SubItems[0].Text;
                int val = int.Parse(listView1.SelectedItems[0].SubItems[1].Text);
                val++;
                m_dict[path]++;
                listView1.SelectedItems[0].SubItems[1].Text = val.ToString();
                System.Diagnostics.Process.Start("EXPLORER.EXE", path);
            }
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

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                OpenExplorer();
            }
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            lvsort.Column = e.Column;
            listView1.Sort();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SearchString();
        }

        private void SearchString()
        { 
            string search_text = textBox1.Text;
            bool found = false;
            // Search Button
            if (search_text.Length > 0)
            {
                for (int i = 0; i < listView1.SelectedItems.Count; i++)
                {
                    listView1.SelectedItems[i].Selected = false; // Unselect all items
                }
                foreach (KeyValuePair<string, int> kvp in m_dict)
                {
                    if (kvp.Key.Contains(search_text))
                    {
                        ListViewItem obj = listView1.FindItemWithText(kvp.Key);
                        if (obj != null)
                        {
                            obj.Selected = true; // Select a matched item
                            label1.Text = search_text + " found.";
                            found = true;
                        }
                    }
                }
            }
            if (found)
            {
                listView1.Focus();
            }
            else
            {
                label1.Text = "Not found.";
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                SearchString();
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyData == Keys.F3) || (e.KeyData == Keys.Divide))
            {
                label1.Text = "focused";
                textBox1.Focus();
            }
        }
    }

    // http://www.eonet.ne.jp/~maeda/cs/listsort.htm
    internal class ListViewSort : IComparer
    {
        public enum ComparerMode
        {
            String, Integer, DateTime
        };

        private ComparerMode[] _columnModes;
        private ComparerMode _mode;
        private int _column;
        private SortOrder _order;

        public ComparerMode[] ColumnModes
        {
            set
            {
                _columnModes = value;
            }
        }

        public int Column
        {
            set
            {
                if (_column == value)
                {
                    if (_order == SortOrder.Ascending)
                        _order = SortOrder.Descending;
                    else if (_order == SortOrder.Descending)
                        _order = SortOrder.Ascending;
                }
                _column = value;
            }
            get
            {
                return _column;
            }
        }

        public ListViewSort(int col, SortOrder ord, ComparerMode cmod)
        {
            _column = col;
            _order = ord;
            _mode = cmod;
        }
        public ListViewSort()
        {
            _column = 0;
            _order = SortOrder.Ascending;
            _mode = ComparerMode.String;
        }

        public int Compare(object x, object y)
        {
            int result = 0;

            ListViewItem itemx = (ListViewItem)x;
            ListViewItem itemy = (ListViewItem)y;

            if (_columnModes != null && _columnModes.Length > _column)
                _mode = _columnModes[_column];

            switch (_mode)
            {
                case ComparerMode.String:
                    result = string.Compare(itemx.SubItems[_column].Text,
                        itemy.SubItems[_column].Text);
                    break;
                case ComparerMode.Integer:
                    result = int.Parse(itemx.SubItems[_column].Text) -
                        int.Parse(itemy.SubItems[_column].Text);
                    break;
                case ComparerMode.DateTime:
                    result = DateTime.Compare(
                        DateTime.Parse(itemx.SubItems[_column].Text),
                        DateTime.Parse(itemy.SubItems[_column].Text));
                    break;
            }

            if (_order == SortOrder.Descending)
                result = -result;
            else if (_order == SortOrder.None)
                result = 0;

            return result;
        }
    }
}
