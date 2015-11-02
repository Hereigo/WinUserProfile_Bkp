namespace WinUserProfile_Bkp
{
    using System;
    using System.IO;
    using System.Windows.Forms;

    public partial class Form1 : Form
    {
        public static string profilesRootDir;
        public static string profileFullPath;

        public Form1()
        {
            InitializeComponent();
            Form1_Load();
            SetupDataGridView();
        }


        private void SetupDataGridView()
        {
            this.dataGridView1.Columns.Add("chk", "");
            this.dataGridView1.Columns.Add("name", "Data Type");
            this.dataGridView1.Columns["name"].DefaultCellStyle.Font = new System.Drawing.Font(Font, System.Drawing.FontStyle.Bold);
            this.dataGridView1.Columns.Add("size", "Size");
            this.dataGridView1.Columns.Add("path", "Full Path To Location");
            this.dataGridView1.Columns.Add("date", "Modified");
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }


        private void Form1_Load()
        {
            this.Text = "Win 7 & XP User Profile BackUp Tool";
            this.MaximizeBox = false;
            //this.BackColor = Color.Brown;
            //this.Size = new Size(350, 125);
            //this.Location = new Point(300, 300);
            InitializeStartParams();
        }


        private void InitializeStartParams()
        {
            String[] profilesFolders = {"C:\\Documents and Settings\\", "C:\\Пользователи\\", "C:\\Users\\"};
            try
            {
                foreach (var profDir in profilesFolders)
                {
                    if (Directory.Exists(profDir)) profilesRootDir = profDir;
                }

                if (profilesRootDir != "")
                {
                    DirectoryInfo[] profiles = new DirectoryInfo(profilesRootDir).GetDirectories();

                    foreach (var item in profiles)
                    {
                        comboBox1.Items.Add(item.Name);
                    }
                }
                else
                {
                    MessageBox.Show("Profiles Directory Not Found!", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show("Oooops!... " + exc.Message);
            }
        }


        private void button1_SearchDataForBkp(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Need Profile To Select !!!");
            }
            else
            {
                string selectedProfile = comboBox1.SelectedItem.ToString();
                profileFullPath = profilesRootDir + selectedProfile;

                string[] possibleOutlookPstDirs = {
                     profileFullPath + "\\Documents\\",
                     profileFullPath + "\\Documents\\Outlook Files\\",
                     profileFullPath + "\\Documents\\Файлы Outlook\\",
                     profileFullPath + "\\Documents\\Файли Outlook\\",
                     profileFullPath + "\\AppData\\Local\\Microsoft\\Outlook\\",
                     profileFullPath +"\\Local Settings\\Application Data\\Microsoft\\Outlook\\"
                };

                string[] possibleBrowsersFavDirs = {
                     profileFullPath + "\\Local Settings\\Application Data\\Google\\Chrome\\User Data\\Default\\",
                     profileFullPath + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\",
                     profileFullPath + "\\AppData\\Local\\Chromium\\User Data\\Default\\",
                     profileFullPath + "\\Favorites\\",
                     profileFullPath + "\\Избранное\\"
                };

                string[] possibleOutlookSignDirs =
                {
                    profileFullPath + "\\AppData\\Roaming\\Microsoft\\Signatures\\",
                    profileFullPath + "\\Application Data\\Microsoft\\Signatures\\"
                };

                string[] possible1CcfgDirs =
                {
                    profileFullPath + "\\AppData\\Roaming\\1C\\1CEStart\\",
                    profileFullPath + "\\Application Data\\1C\\1CEStart\\"
                };

                SearchDataForBkp(possibleOutlookPstDirs, "*.pst");
                SearchDataForBkp(possibleBrowsersFavDirs, "Bookmarks*");
                SearchDataForBkp(possibleBrowsersFavDirs, "*.url");
                SearchDataForBkp(possibleOutlookSignDirs, "*.*");
                SearchDataForBkp(possible1CcfgDirs, "ibases.v8i");
            }
        }


        private void RecursingSignatureFolder(string path)
        {
            foreach (var file in new DirectoryInfo(path).GetFiles())
            {

            }
        }


        private void SearchDataForBkp(string[] dirsForSearch, string fileMask)
        {
            try
            {
                long serchingFilesSize = 0;
                DateTime SearchingFilesDate = DateTime.Now.AddYears(-1);
                string path = "";

                foreach (var pstPath in dirsForSearch)
                {
                    if (Directory.Exists(pstPath))
                    {
                        FileInfo[] filesPst = new DirectoryInfo(pstPath).GetFiles(fileMask);

                        for (int i = 0; i < filesPst.Length; i++)
                        {
                            serchingFilesSize += filesPst[i].Length;

                            if (DateTime.Compare(filesPst[i].LastWriteTime, SearchingFilesDate) > -1)
                            {
                                SearchingFilesDate = filesPst[i].LastWriteTime;
                                path = pstPath;
                            }
                        }
                    }
                }

                // TODO:  OPTIMIZE THIS !!!!!!!!
                // TODO:  OPTIMIZE THIS !!!!!!!!
                // TODO:  OPTIMIZE THIS !!!!!!!!

                string sizeInStr;
                if (serchingFilesSize < 1024) sizeInStr = serchingFilesSize + " bytes.";
                else if ((serchingFilesSize / 1024) < 1024) sizeInStr = serchingFilesSize / 1024 + " KB.";
                else sizeInStr = (serchingFilesSize / 1024 / 1024) + " MB.";

                dataGridView1.Rows.Add("", fileMask, sizeInStr, path, SearchingFilesDate);
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
