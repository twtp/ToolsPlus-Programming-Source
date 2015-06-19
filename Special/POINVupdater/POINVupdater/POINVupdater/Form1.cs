using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
namespace POINVupdater
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Shown += new EventHandler(Form1_Shown);
            
        }

        void Form1_Shown(object sender, EventArgs e)
        {
            label1.Refresh();
            JustDoIt();
        }

        private void JustDoIt()
        {
            this.Visible = false;
            string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string DocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string POINVPath = @"S:\mastest\mas90-signs\A_Dist\POINV Betas\poinv - B5";
            string POINVDependenciesPath = POINVPath + "\\dependencies";

            bool poinvOnDesktopExists = File.Exists(DesktopPath + "\\poinv-b5.exe");

            bool poinvOnServerExists = File.Exists(POINVPath + "\\poinv-b5.exe");

            bool is64Bit = Directory.Exists("C:\\Windows\\SysWow64");

            if (poinvOnServerExists == false)
            {
                MessageBox.Show("Could not find the copy of POINV that resides on the server. Either it was removed or you do not have your S:\\ drive mapped. Call Tom and let him know.");
                return;
            }
            else
            {
                //File.Copy(POINVPath + "\\poinv-b5.exe", DesktopPath + "\\poinv-b5.exe",true);

                //now check for dependencies...
                string[] deps = Directory.GetFiles(POINVDependenciesPath);
                foreach (string dep in deps)
                {
                    if (is64Bit == true)
                    {
                        string filePath = "C:\\Windows\\SysWow64\\" + dep.Split('\\').Last<string>();
                        if (File.Exists(filePath) == true)
                        {
                            if (File.Exists(filePath) == false)
                            {
                                File.Copy(dep, filePath, false);
                            }
                        }
                        string command = "/C regsvr32.exe /s" + filePath;
                        System.Diagnostics.Process.Start("cmd.exe", command);
                    }
                    else
                    {
                        string filePath = "C:\\Windows\\System32\\" + dep.Split('\\').Last<string>();
                        if (File.Exists(filePath) == false)
                        {
                            if (File.Exists(filePath) == false)
                            {
                                File.Copy(dep, filePath, false);
                            }
                        }
                        string command = "/C regsvr32.exe /s" + filePath;
                        System.Diagnostics.Process.Start("cmd.exe", command);
                    }

                }

                System.Diagnostics.Process[] pList = System.Diagnostics.Process.GetProcessesByName("poinv-b5");
                if (pList.GetLength(0) > 0)
                {
                    //DialogResult res = MessageBox.Show("Warning: Close all POINVs before running!");
                    //Application.Exit();



                    /*if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
                    {
                        foreach (System.Diagnostics.Process p in pList)
                        {
                            try
                            {
                                p.Kill();
                                p.WaitForExit();
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        Application.Exit();
                    }*/
                }



                try
                {
                    //Kirks Way:
                    //File.Copy(POINVPath + "\\poinv-b5.exe", "c:\\poinv-b5.exe", true);
                    //MessageBox.Show("Updated OK. Press OK to Launch.");
                    //System.Diagnostics.Process.Start("C:\\poinv-b5.exe");
                    //Application.Exit();


                    //Erics Way:
                    File.Copy(POINVPath + "\\poinv-b5.exe", DocumentsPath + "\\poinv-b5.exe", true);
                    //MessageBox.Show("Updated OK. Press OK to Launch.");
                               
                }
                catch
                {
                    MessageBox.Show("POINV was not updated. Make sure you arent running a copy already.");
                    //Application.Exit();

                }
                try
                {
                    System.Diagnostics.Process.Start(DocumentsPath + "\\poinv-b5.exe");
                    Application.Exit();
                }
                catch
                {
                    MessageBox.Show("Could not find POINV on your computer. Perhaps it did not update from the server. Rerun this app and try again.");
                }

            }




        }
    }
}
