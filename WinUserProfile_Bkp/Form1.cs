namespace WinUserProfile_Bkp
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Security.AccessControl;
    using System.Windows.Forms;

    public partial class Form1 : Form
    {
        public static string profilesRootDir;
        public static string profileFullPath;
        public static string resultFolderForeachSearch = "";

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
            String[] profilesFolders = { "C:\\Documents and Settings\\", "C:\\Пользователи\\", "C:\\Users\\" };
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
                // CLEAR D.G.V. BEFORE FILL IT !
                if (dataGridView1.Rows.Count > 1) dataGridView1.Rows.Clear();

                string selectedProfile = comboBox1.SelectedItem.ToString();
                profileFullPath = profilesRootDir + selectedProfile;

                string[] possibleOutlookPstDirs = {
                     profileFullPath + "\\Documents\\Outlook Files\\",
                     profileFullPath + "\\Documents\\Файлы Outlook\\",
                     profileFullPath + "\\Documents\\Файли Outlook\\",
                     profileFullPath + "\\AppData\\Local\\Microsoft\\Outlook\\",
                     profileFullPath +"\\Local Settings\\Application Data\\Microsoft\\Outlook\\"
                     
                     // profileFullPath + "\\Documents\\"  

                     // Acces Denied !!! on Win 7...10  !!!!!!!!!
                     // Acces Denied !!! on Win 7...10  !!!!!!!!!
                     // Acces Denied !!! on Win 7...10  !!!!!!!!!
                };

                string[] possibleChromeFavDirs = {
                     profileFullPath + "\\Local Settings\\Application Data\\Google\\Chrome\\User Data\\Default\\",
                     profileFullPath + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\",
                     profileFullPath + "\\AppData\\Local\\Chromium\\User Data\\Default\\"
                };

                string[] possibleIExplorerFavDirs = {
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
                SearchDataForBkp(possibleChromeFavDirs, "Bookmarks*");
                SearchDataForBkp(possibleIExplorerFavDirs, "*.url");
                SearchDataForBkp(possibleOutlookSignDirs, "*.*");
                SearchDataForBkp(possible1CcfgDirs, "ibases.v8i");

                string[] myDoxPathes =
                {
                    profileFullPath + "\\Documents\\",
                    profileFullPath + "\\My Documents\\",
                    profileFullPath + "\\Мои Документы\\",
                    profileFullPath + "\\Мої Документи\\"
                };

                string[] myDeskPathes =
                {
                    profileFullPath + "\\Desktop\\",
                    profileFullPath + "\\My Desktop\\",
                    profileFullPath + "\\Рабочий стол\\",
                    profileFullPath + "\\Робочий стіл\\"
                };

                SearchDoxAndDesk(myDoxPathes);
                SearchDoxAndDesk(myDeskPathes);
            }
        }

        private void SearchDoxAndDesk(string[] myDoxPathes)
        {
            foreach (string usersFolder in myDoxPathes)
            {
                if (Directory.Exists(usersFolder))
                {
                    DirectoryInfo di = new DirectoryInfo(usersFolder);
                    long dirSize = 0;
                    try
                    {
                        dirSize = di.EnumerateFiles("*.*").Sum(fi => fi.Length);
                        // TODO: HAVE TO CHECK ACCESS ALLOWED !!!
                        // TODO: HAVE TO CHECK ACCESS ALLOWED !!!
                        // TODO: HAVE TO CHECK ACCESS ALLOWED !!!
                        // dirSize = di.EnumerateFiles("*.*", SearchOption.AllDirectories).Sum(fi => fi.Length);
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message);
                    }

                    if (dirSize > 0)
                    {
                        string sizeInStr;
                        if (dirSize < 1024) sizeInStr = dirSize + " bytes.";
                        else if ((dirSize / 1024) < 1024) sizeInStr = dirSize / 1024 + " KB.";
                        else sizeInStr = (dirSize / 1024 / 1024) + " MB.";
                        dataGridView1.Rows.Add("", di.Name, sizeInStr, di.FullName, di.LastAccessTime.ToShortDateString());
                    }
                }
            }
        }

        private void SearchDataForBkp(string[] dirsForSearch, string fileMask)
        {
            try
            {
                long serchingDataSize = 0;
                DateTime searchingFileDate = DateTime.Parse("01.01.01");
                string path = "";

                foreach (var currentDir in dirsForSearch)
                {
                    if (Directory.Exists(currentDir))
                    {
                        // PROTECT ROOT FOLDER SETTINGS FROM CHANGE BY RECURSE:
                        if (resultFolderForeachSearch == "") resultFolderForeachSearch = currentDir;

                        recursiveSearch(currentDir, fileMask, ref serchingDataSize, ref searchingFileDate, ref path);
                    }
                }
                // TODO:  OPTIMIZE THIS !!!!!!!!
                // TODO:  OPTIMIZE THIS !!!!!!!!
                // TODO:  OPTIMIZE THIS !!!!!!!!
                string sizeInStr;
                if (serchingDataSize < 1024) sizeInStr = serchingDataSize + " bytes.";
                else if ((serchingDataSize / 1024) < 1024) sizeInStr = serchingDataSize / 1024 + " KB.";
                else sizeInStr = (serchingDataSize / 1024 / 1024) + " MB.";

                dataGridView1.Rows.Add("", fileMask, sizeInStr, resultFolderForeachSearch, searchingFileDate.ToShortDateString());

                // CLEAR ROOT FOLDER FOR THE NEXT SEARCH CYCLE:
                resultFolderForeachSearch = "";

            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private static void recursiveSearch(string currentDir, string fileMask, ref long serchingDataSize, ref DateTime mostFreshDate, ref string path)
        {
            FileInfo[] searchingFiles = new DirectoryInfo(currentDir).GetFiles(fileMask);

            for (int i = 0; i < searchingFiles.Length; i++)
            {
                serchingDataSize += searchingFiles[i].Length;

                if (DateTime.Compare(searchingFiles[i].LastWriteTime, mostFreshDate) > -1)
                {
                    mostFreshDate = searchingFiles[i].LastWriteTime;
                    path = currentDir;
                }
            }

            DirectoryInfo[] subDirsHere = new DirectoryInfo(currentDir).GetDirectories();

            foreach (DirectoryInfo dir in subDirsHere)
            {
                recursiveSearch(dir.FullName, fileMask, ref serchingDataSize, ref mostFreshDate, ref path);
            }
        }


        public static bool FolderCanRead(string path)
        {
            var readAllow = false;
            var readDeny = false;
            var accessControlList = Directory.GetAccessControl(path);
            if (accessControlList == null)
                return false;
            var accessRules = accessControlList.GetAccessRules(true, true, typeof(System.Security.Principal.SecurityIdentifier));
            if (accessRules == null)
                return false;

            foreach (FileSystemAccessRule rule in accessRules)
            {
                if ((FileSystemRights.Read & rule.FileSystemRights) != FileSystemRights.Read) continue;

                if (rule.AccessControlType == AccessControlType.Allow)
                    readAllow = true;
                else if (rule.AccessControlType == AccessControlType.Deny)
                    readDeny = true;
            }

            return readAllow && !readDeny;
        }

        private void button_RunZip_Click(object sender, EventArgs e)
        {
            const string wRar = @"C:\Progra~1\WinRAR\Rar.exe";

            //// string userName = comboBox1.SelectedItem.ToString().Replace(' ', '_') + "_" + archName + ".RAR";

            ////String path = root + comboBox1.SelectedItem + "\\" + dirFrom;

            //String zipParams = "m[f] -ms -rr3p " + excluMask + " -r -t \"" + textBox1.Text + "\\" + userName + "\" \"" + path + "\\*." + incluMask + "\"";

            //ProcessStartInfo psi = new ProcessStartInfo(rarPath, zipParams);
            //Process.Start(psi).WaitForExit();

            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if (!String.IsNullOrWhiteSpace(dataGridView1[3, i].Value.ToString()))
                    {
                        string mask;
                        if (dataGridView1[1, i].Value.ToString().Contains("*.pst")) mask = "*.pst";
                        else if (dataGridView1[1, i].Value.ToString().Contains("*.url")) mask = "*.url";
                        else if (dataGridView1[1, i].Value.ToString().Contains("Bookmarks*")) mask = "Bookmarks*";
                        else if (dataGridView1[1, i].Value.ToString().Contains("ibases.v8i")) mask = "ibases.v8i";
                        // Special For myDoxPathes;
                        // Special For myDeskPathes;
                        else mask = "*.*";

                        string archiveName = comboBox1.SelectedItem.ToString().Replace(' ', '_') + "_" + mask.Replace("*", "").ToUpper() + ".RAR";

                        // TODO: EXCLUDE - AVI, MPEG4, MP3,  . . . ????  !!!!!!!!!!
                        // TODO: EXCLUDE - AVI, MPEG4, MP3,  . . . ????  !!!!!!!!!!
                        // TODO: EXCLUDE - AVI, MPEG4, MP3,  . . . ????  !!!!!!!!!!



                        //  PATH = dataGridView1[3, i].Value.ToString();
                    }


                }

                //Process.Start(dataGridView1[3, 0].Value.ToString());

            }
            catch (Exception exc)
            {
                MessageBox.Show("Ooops!... " + exc);
            }

            //SearchDataForBkp(possibleOutlookPstDirs, "*.pst");
            //SearchDataForBkp(possibleChromeFavDirs, "Bookmarks*");
            //SearchDataForBkp(possibleIExplorerFavDirs, "*.url");
            //SearchDataForBkp(possibleOutlookSignDirs, "*.*");
            //SearchDataForBkp(possible1CcfgDirs, "ibases.v8i");

            //SearchDoxAndDesk(myDoxPathes);
            //SearchDoxAndDesk(myDeskPathes);
        }
    }
}
