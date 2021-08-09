/*
##############################################################################################################################################################
#   Date      #  Version   #     Author        #                               Comments                                                                      #
##############################################################################################################################################################
# 06/03/2021  # 1.0.0.0    # Remi G Grandsire   #  Initial development                                                                                       #
# 06/08/2021  # 1.0.0.1    # Remi G Grandsire   # Added: display version                                                                                     #
#             #            #                    # Save file as with version added                                                                            #
# 08/09/2021  # 1.0.0.2    # Remi G Grandsire   # Set default dir to the download folder                                                                     #
#             #            #                    # Do not save if no file was open or version was not found                                                   #
##############################################################################################################################################################


*/

using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Win32;
using System.Reflection;
using System.Text.RegularExpressions;

namespace PresetVersionCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string zFile = string.Empty;
        private string zipPath = "C:\\temp\\RemUnzip";
        private string zVersion = string.Empty;


        private string GetVersion(string xmlFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFile);
            return doc.OuterXml.ToString();
        }
        public static void Decompress(string filename, string outputfolder)
        {
            ZipStorer zip;

            if (File.Exists(filename))
                // Opens existing zip file
                zip = ZipStorer.Open(filename, FileAccess.Read);
            else
                return;

            // Read all directory contents
            List<ZipStorer.ZipFileEntry> dir = zip.ReadCentralDir();

            // Extract all files in target directory
            string path;
            bool result;
            foreach (ZipStorer.ZipFileEntry entry in dir)
            {
                path = Path.Combine(outputfolder, entry.FilenameInZip);
                result = zip.ExtractFile(entry, path);
            }
            zip.Close();
        }
        private void UnComp()
        {
            if (Directory.Exists(zipPath) != true)
            {
                Directory.CreateDirectory(zipPath);
                AddLog("created folder: " + zipPath, listBox1);
            }
            else AddLog("Temp folder exists", listBox1);
            AddLog("Preparing to uncompress: ", listBox1);
            Decompress(zFile, zipPath);
            AddLog("Uncompress completed", listBox1);
        }
        private static void AddLog(string zMess, ListBox listBox2)
        {
           listBox2.Items.Add(DateTime.Now.ToString()+"  :   "+ zMess);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Let's sdave the file with the version if needed
            if (checkBox1.Checked && zVersion != string.Empty)
            {
                zFile = Path.GetFileNameWithoutExtension(zFile);
                
                if (!zFile.Contains("V1_0_"))
                {
                    zVersion = zVersion.Replace('.', '_');
                    zFile += "V" + zVersion +".mpt";
                    AddLog(zFile, listBox1);
                    saveFileDialog1.FileName = zFile;
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        File.Move(openFileDialog1.FileName, saveFileDialog1.FileName);
                    }

                }
            }

            Application.Exit();
        }

        private void GetiItDone()
        {
            int start = 0;
            int stop = 0;
            string zResult = string.Empty;

            zFile = openFileDialog1.FileName;
            AddLog("Preset file: " + zFile, listBox1);
            AddLog("Opening Preset file for version check", listBox1);
            UnComp();
            zResult = GetVersion(zipPath + "\\Preset.xml");
            AddLog(zResult, listBox1);
            start = zResult.IndexOf("build");
            stop = zResult.IndexOf("name");
            zVersion = zResult.Substring(start, stop - start);
            AddLog(zVersion, listBox1);
            zVersion = zVersion.Substring(7, zVersion.Length - 9);

            label1.Text = "Preset Version: " + zVersion;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string defaultDir = string.Empty;

            defaultDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
            openFileDialog1.InitialDirectory = defaultDir;
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.DefaultExt = "*.mpt";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                GetiItDone();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            zFile = string.Empty;
            openFileDialog1.FileName = string.Empty;
            label1.Text = string.Empty;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string currentSavedDefault = "";
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(@"filetype\shell\open\command"))
                {
                    if (key != null)
                    {
                        currentSavedDefault = key.GetValue("").ToString();
                    }
                }
                String myExecutable = Assembly.GetEntryAssembly().Location;
                if (currentSavedDefault != myExecutable)
                {
                    Registry.ClassesRoot.CreateSubKey(".type").SetValue("", "filetype", Microsoft.Win32.RegistryValueKind.String);
                    Registry.ClassesRoot.CreateSubKey(@"filetype\shell\open\command").SetValue("", myExecutable + " %1", Microsoft.Win32.RegistryValueKind.String);
                }
            }

            catch (Exception)
            {
            }
    string[] args = Environment.GetCommandLineArgs();

                if (args.Length > 1)
                {
                    zFile  = args[1];
                    openFileDialog1.FileName = zFile;
                    GetiItDone();
                }
        }
    }
}
