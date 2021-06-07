/*
##############################################################################################################################################################
#   Date      #  Version   #     Author        #                               Comments                                                                      #
##############################################################################################################################################################
# 06/03/2021  # 1.0.0.0    # Remi G Grandsire  #  Initial development                                                                                        #
#_____________#____________#___________________#_____________________________________________________________________________________________________________#


*/

using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;

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
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                int start = 0;
                int stop = 0;
                string zResult = string.Empty;
                string zVersion = string.Empty;
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
                zVersion = zVersion.Substring(7, zVersion.Length -9);

                label1.Text = "Preset Version: " + zVersion;    
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            zFile = string.Empty;
            openFileDialog1.FileName = string.Empty;
            label1.Text = string.Empty;
        }
    }
}
