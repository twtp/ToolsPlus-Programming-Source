using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace DocumentationWriter
{
    public partial class Form1 : Form
    {
        public struct MethodInfo
        {
           public string MethodName;
           public string MethodDescription;
        }
        public struct VariablesInfo
        {
           public string VariableName;
           public string VariableDescription;
        }
        public struct ClassInfo
        {
            public string ClassName;
            public string ClassDescription;
            public List<VariablesInfo> VariablesList;
            public List<MethodInfo> MethodList;
        }
        public struct Documentation
        {
            public List<ClassInfo> ClassInfoList;
            public string ProjectName;
            public string ProjectVersion;
            public string ProjectLastBuildDate;

        }
        //public List<ClassInfo> ClassInfoList = new List<ClassInfo>();

        public Documentation thisDocumentation = new Documentation();

        
        public List<VariablesInfo> VariableInfoList = new List<VariablesInfo>();
        public List<MethodInfo> MethodInfoList = new List<MethodInfo>();
        public string CurrentDataFile;

        public Form1()
        {
            InitializeComponent();
            thisDocumentation.ClassInfoList = new List<ClassInfo>();
            
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text += "{VERSION}";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text += "{LAST_BUILD_DATE}";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox2.Text += "{VERSION}";

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text += "{LAST_BUILD_DATE}";
        }

        private void AddToDictionary()
        {
            string ClassName = textBox3.Text;
            string ClassDescription = textBox4.Text;
            List<VariablesInfo> VariableInfoList = new List<VariablesInfo>();
            


        }

        private void button5_Click(object sender, EventArgs e)
        {
            VariablesInfo newVariableInfo = new VariablesInfo();
            newVariableInfo.VariableName = textBox5.Text;
            newVariableInfo.VariableDescription = textBox6.Text;
            VariableInfoList.Add(newVariableInfo);
            listBox1.Items.Add(newVariableInfo.VariableName);
            textBox5.Text = "";
            textBox6.Text = "";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            MethodInfo newMethodInfo = new MethodInfo();
            newMethodInfo.MethodName = textBox8.Text;
            newMethodInfo.MethodDescription = textBox7.Text;
            MethodInfoList.Add(newMethodInfo);
            listBox2.Items.Add(newMethodInfo.MethodName);
            textBox8.Text = "";
            textBox7.Text = "";
        }

      

        private void CreateDataFile()
        {
            CurrentDataFile = "";
            string bodyFormatted = textBox1.Text;
            bodyFormatted = bodyFormatted.Replace("{VERSION}", textBox2.Text);
            bodyFormatted = bodyFormatted.Replace("{LAST_BUILD_DATE}", textBox9.Text);
            CurrentDataFile += "\r\n<projectname>" + textBox10.Text + "</projectname>\r\n";
            CurrentDataFile += "\r\n<body>\r\n" + bodyFormatted + "\r\n</body>\r\n";
            CurrentDataFile += "\r\n<version>" + textBox2.Text + "</version>\r\n";
            CurrentDataFile += "\r\n<lastbuild>" + textBox9.Text + "</lastbuild>\r\n";
            foreach (ClassInfo classInformation in thisDocumentation.ClassInfoList)
            {
                CurrentDataFile += "\r\n<class>\r\n";
                CurrentDataFile += " <name>" + classInformation.ClassName + "</name>\r\n";
                CurrentDataFile += " <description>" + classInformation.ClassDescription + "</description>\r\n";
                foreach(VariablesInfo variableInformation in classInformation.VariablesList)
                {
                    CurrentDataFile += "\r\n <variable>\r\n";
                    CurrentDataFile += "  <name>" + variableInformation.VariableName + "</name>\r\n";
                    CurrentDataFile += "  <description>" + variableInformation.VariableDescription + "</description>\r\n";
                    CurrentDataFile += " </variable>\r\n";
                }
                foreach (MethodInfo methodInformation in classInformation.MethodList)
                {
                    CurrentDataFile += "\r\n <method>\r\n";
                    CurrentDataFile += "  <name>" + methodInformation.MethodName + "</name>\r\n";
                    CurrentDataFile += "  <description>" + methodInformation.MethodDescription + "</description>\r\n";
                    CurrentDataFile += " </method>\r\n";

                }
                CurrentDataFile += "</class>\r\n";
            }
            

        }

        private void button9_Click(object sender, EventArgs e)
        {
            ClassInfo newClassInfo = new ClassInfo();
            newClassInfo.ClassName = textBox3.Text;
            newClassInfo.ClassDescription = textBox4.Text;
            newClassInfo.VariablesList = VariableInfoList;
            newClassInfo.MethodList = MethodInfoList;
            //ClassInfoList.Add(newClassInfo);
            thisDocumentation.ClassInfoList.Add(newClassInfo);
            listBox3.Items.Add(newClassInfo.ClassName);
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";

        }

        private void button15_Click(object sender, EventArgs e)
        {
            CreateDataFile();

            folderBrowserDialog1.Description = "Choose a folder to save to:";
            folderBrowserDialog1.ShowDialog();
            string savePath = folderBrowserDialog1.SelectedPath;

            savePath = savePath + "\\" + CurrentDataFile.Split(new string[] {"<projectname>"},StringSplitOptions.None)[1].Split(new string[] { "</projectname>" }, StringSplitOptions.None)[0] + ".txt";
            System.IO.File.WriteAllText(savePath,CurrentDataFile);
            MessageBox.Show("Check '"+savePath+"' for the file!");
        }


        private string GenerateHTMLFromDataFile(string DataFile)
        {
            DataFileInfo getDataFileInfo = DataFileToVariables(DataFile);
            string HTMLCODE = "";
            HTMLCODE += "<html><body><font size=\"14px\">Project Name: <b>" + getDataFileInfo.ProjectName + "</b></font><br>";
            HTMLCODE += "<font size=\"12px\"> &nbspVersion: <b>" + getDataFileInfo.Version + "</b></font><br>";
            HTMLCODE += "<font size=\"10px\"> &nbspLast Build Date: <b>" + getDataFileInfo.LastBuildDate + "</b></font><br>";
            HTMLCODE += "<br><br><font size = \"10px\"> <b>Description: </b><br> &nbsp &nbsp " + getDataFileInfo.MainDescription + "</font><br>";
            HTMLCODE += "<br><br>";
            HTMLCODE += "<table border=\"2\" width=\"600px\">";
            HTMLCODE += "<th><b><font size = \"12px\">Documentation</font></b></th>";
            HTMLCODE += "</table><br>";

            foreach (ClassInfo classInfo in getDataFileInfo.Classes)
            {
                HTMLCODE += "<table border=\"2\" width=\"600px\">";
                HTMLCODE += "<tr><td><b>Class Name:</b></td><td>" + classInfo.ClassName + "</td></tr>";
                HTMLCODE += "<tr><td><b>Class Description:</b></td><td>" + classInfo.ClassDescription + "</td></tr>";
                HTMLCODE += "<tr><td><hr></td><td><hr></td></tr>";
                HTMLCODE += "<tr><th><center><b>Variables</b></center></th></tr>";
                foreach (VariablesInfo variablesInfo in classInfo.VariablesList)
                {
                    HTMLCODE += "<tr><td><hr></td><td><hr></td></tr>";
                    HTMLCODE += "<tr><td><b>Variable</b></td><td>" + variablesInfo.VariableName + "</td></tr>";
                    HTMLCODE += "<tr><td><b>Description</b></td><td>" + variablesInfo.VariableDescription + "</td></tr>";
                    
                }
                HTMLCODE += "<tr><td><hr></td><td><hr></td></tr>";
                HTMLCODE += "<tr><td><center><b>Methods</b></center></td></tr>";
                foreach (MethodInfo methodInfo in classInfo.MethodList)
                {
                    HTMLCODE += "<tr><td><hr></td><td><hr></td></tr>";
                    HTMLCODE += "<tr><td><b>Method</b></td><td>" + methodInfo.MethodName + "</td></tr>";
                    HTMLCODE += "<tr><td><b>Description</b></td><td>" + methodInfo.MethodDescription + "</td></tr>";

                }
                HTMLCODE += "</table><br><br>";

            }


            HTMLCODE += "</table>";
            HTMLCODE += "<br><br>";
            HTMLCODE += "</body></html>";


            return HTMLCODE;
              
        }
        private string GenerateTEXTFromDataFile(string DataFile)
        {
            DataFileInfo getDataFileInfo = DataFileToVariables(DataFile);
            return "";
        }

        List<ClassInfo> INDATA_ClassInfoList = new List<ClassInfo>();
        //List<VariablesInfo> INDATA_VariableInfoList = new List<VariablesInfo>();
        //List<MethodInfo> INDATA_MethodInfoList = new List<MethodInfo>();
        public struct DataFileInfo
        {
           public string ProjectName;
           public string MainDescription;
           public string Version;
           public string LastBuildDate;
           public List<ClassInfo> Classes;           
        }

        private DataFileInfo DataFileToVariables(string DataFile)
        {
            DataFileInfo newDataFileInfo = new DataFileInfo();
            
            //List<VariablesInfo> varInfoList = new List<VariablesInfo>();
            //List<MethodInfo> methInfoList = new List<MethodInfo>();
            string MainDescription = "";
            string Version = "";
            string LastBuildDate = "";
            string ProjectName = "";
            ProjectName = DataFile.Split(new string[] { "<projectname>" }, StringSplitOptions.None)[1].Split(new string[] { "</projectname>" }, StringSplitOptions.None)[0].Trim();
            MainDescription = DataFile.Split(new string[] { "<body>" }, StringSplitOptions.None)[1].Split(new string[] { "</body>" }, StringSplitOptions.None)[0].Trim();
            Version = DataFile.Split(new string[] { "<version>" }, StringSplitOptions.None)[1].Split(new string[] { "</version>" }, StringSplitOptions.None)[0].Trim();
            LastBuildDate = DataFile.Split(new string[] { "<lastbuild>" }, StringSplitOptions.None)[1].Split(new string[] { "</lastbuild>" }, StringSplitOptions.None)[0].Trim();

            int classCount = DataFile.Split(new string[] { "<class>" }, StringSplitOptions.None).GetLength(0) - 1;

            for (int classCounter = 0; classCounter < classCount; classCounter++)
            {
                ClassInfo newClass = new ClassInfo();
                newClass.MethodList = new List<MethodInfo>();
                newClass.VariablesList = new List<VariablesInfo>();
                string thisClass = DataFile.Split(new string[] { "<class>" }, StringSplitOptions.None)[classCounter+1].Split(new string[] { "</class>" }, StringSplitOptions.None)[0];

                newClass.ClassName = thisClass.Split(new string[] { "<name>" }, StringSplitOptions.None)[1].Split(new string[] { "</name>" }, StringSplitOptions.None)[0];
                newClass.ClassDescription = thisClass.Split(new string[] { "<description>" }, StringSplitOptions.None)[1].Split(new string[] { "</description>" }, StringSplitOptions.None)[0];

                int variableCount = thisClass.Split(new string[] { "<variable>" }, StringSplitOptions.None).GetLength(0) - 1;
                int methodCount = thisClass.Split(new string[] { "<method>" }, StringSplitOptions.None).GetLength(0) - 1;


                for (int variableCounter = 0; variableCounter < variableCount; variableCounter++)
                {
                    VariablesInfo varinfo = new VariablesInfo();
                    string thisVariable = thisClass.Split(new string[] { "<variable>" }, StringSplitOptions.None)[variableCounter].Split(new string[] { "</variable>" }, StringSplitOptions.None)[0];
                    varinfo.VariableName = thisVariable.Split(new string[] { "<name>" }, StringSplitOptions.None)[1].Split(new string[] { "</name>" }, StringSplitOptions.None)[0];
                    varinfo.VariableDescription = thisVariable.Split(new string[] { "<description>" }, StringSplitOptions.None)[1].Split(new string[] { "</description>" }, StringSplitOptions.None)[0];
                    newClass.VariablesList.Add(varinfo);

                }
                for (int methodCounter = 0; methodCounter < methodCount; methodCounter++)
                {
                    MethodInfo methinfo = new MethodInfo();
                    string thisMethod = thisClass.Split(new string[] { "<method>" }, StringSplitOptions.None)[methodCounter].Split(new string[] { "</method>" }, StringSplitOptions.None)[0];
                    methinfo.MethodName = thisMethod.Split(new string[] { "<name>" }, StringSplitOptions.None)[1].Split(new string[] { "</name>" }, StringSplitOptions.None)[0];
                    methinfo.MethodDescription = thisMethod.Split(new string[] { "<description>" }, StringSplitOptions.None)[1].Split(new string[] { "</description>" }, StringSplitOptions.None)[0];
                    newClass.MethodList.Add(methinfo);
                }

                INDATA_ClassInfoList.Add(newClass);


            }

            newDataFileInfo.ProjectName = ProjectName;
            newDataFileInfo.MainDescription = MainDescription;
            newDataFileInfo.Version = Version;
            newDataFileInfo.Classes = INDATA_ClassInfoList;


            return newDataFileInfo;


        }

        private void button12_Click(object sender, EventArgs e)
        {
            
            
        }

        private void button11_Click(object sender, EventArgs e)
        {
            CreateDataFile();
            string HTML = GenerateHTMLFromDataFile(CurrentDataFile);
            //MessageBox.Show(HTML);
            Form2 htmlPreviewForm = new Form2();
            htmlPreviewForm.Show();
            htmlPreviewForm.LoadPreviewDocument(HTML);
            
        }
        
        

    }
}
