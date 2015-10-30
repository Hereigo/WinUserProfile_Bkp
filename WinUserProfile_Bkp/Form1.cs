namespace WinUserProfile_Bkp
{
    using System;
    using System.IO;
    using System.Windows.Forms;

    public partial class Form1 : Form
    {
        public static string profilesRootDir = "";
        public static string profileName;

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
                profileName = comboBox1.SelectedItem.ToString();

                SearchDataForBkp(profilesRootDir + profileName);
            }
        }


        private void SearchDataForBkp(string profileFullPath)
        {
            try
            {
                string[] possibleOutlookPstDirs = {
                     profileFullPath + "\\Documents\\",
                     profileFullPath + "\\Documents\\Outlook Files\\",
                     profileFullPath + "\\Documents\\Файлы Outlook\\",
                     profileFullPath + "\\Documents\\Файли Outlook\\",
                     profileFullPath + "\\AppData\\Local\\Microsoft\\Outlook\\",
                     profileFullPath +"\\Local Settings\\Application Data\\Microsoft\\Outlook\\"
                };

                long filesPstSize = 0;

                foreach (var pstPath in possibleOutlookPstDirs)
                {
                    if (Directory.Exists(pstPath))
                    {
                        FileInfo[] filesPst = new DirectoryInfo(pstPath).GetFiles("*.pst");

                        for (int i = 0; i < filesPst.Length; i++)
                        {
                            filesPstSize += filesPst[i].Length;
                        }

                        label_Outlook_Size.Text = (filesPstSize / 1024 / 1024).ToString();
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
