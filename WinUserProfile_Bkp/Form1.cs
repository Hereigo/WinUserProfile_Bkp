namespace WinUserProfile_Bkp
{
    using System;
    using System.IO;
    using System.Windows.Forms;

    public partial class Form1 : Form
    {
        public static string profilesRootDir;
        public static string profileFullPath;
        // public static string outlookPstPath;

        public Form1()
        {
            InitializeComponent();
            Form1_Load();
        }


        private void Form1_Load()
        {
            this.Text = "Change Prperties Through Coding";
            this.MaximizeBox = false;
            //this.BackColor = Color.Brown;
            //this.Size = new Size(350, 125);
            //this.Location = new Point(300, 300);
            InitializeStartParams();
        }


        private void InitializeStartParams()
        {
            String[] profilesFolders =
            {
                "C:\\Documents and Settings\\",
                "C:\\Пользователи\\",
                "C:\\Users\\"
            };

            try
            {
                foreach (var profDir in profilesFolders)
                {
                    if (Directory.Exists(profDir)) profilesRootDir = profDir;
                }

                if (profilesRootDir != "")
                {
                    loadProfListToCombo(profilesRootDir);
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


        private void loadProfListToCombo(string profDir)
        {
            DirectoryInfo[] profiles = new DirectoryInfo(profDir).GetDirectories();
            foreach (var item in profiles)
            {
                comboBox1.Items.Add(item.Name);
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
                     profileFullPath + "\\AppData\\Local\\Grome\\Chrome\\User Data\\Default\\",
                     profileFullPath + "\\Favorites\\",
                     profileFullPath + "\\Избранное\\"
                };

                SearchDataForBkp(possibleOutlookPstDirs, "*.pst");
            }
        }


        private void SearchBrowsersFav(string profileFullPath)
        {
            try
            {
                // \\192.168.0.23\c$\Users\D.Mohylnitskyy\AppData\Roaming\Microsoft\Signatures\
                // \\192.168.0.23\c$\Users\D.Mohylnitskyy\AppData\Roaming\1C\1CEStart\

                //  Bookmarks
                
                
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }


        private void SearchDataForBkp(string[] dirsForSearch, string fileMask)
        {
            try
            {
                long filesPstSize = 0;
                DateTime yearAgoDate = DateTime.Now.AddYears(-1);

                foreach (var pstPath in dirsForSearch)
                {
                    if (Directory.Exists(pstPath))
                    {
                        FileInfo[] filesPst = new DirectoryInfo(pstPath).GetFiles(fileMask);

                        for (int i = 0; i < filesPst.Length; i++)
                        {
                            filesPstSize += filesPst[i].Length;

                            if (DateTime.Compare(filesPst[i].LastWriteTime, yearAgoDate) > -1)
                            {
                                yearAgoDate = filesPst[i].LastWriteTime;
                                //outlookPstPath = pstPath;
                            }
                        }

                        label_Outlook_Size.Text = (filesPstSize / 1024 / 1024).ToString() + " (" + yearAgoDate + ")";
                        label_4outlookPstPath.Text = pstPath;
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
