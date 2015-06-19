using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public class PriceComparisonReturnData
        {
            public string ItemNumber;
            public decimal ItemPrice;


        }

        public struct CSharpScript
        {
            public string ScriptName;
            public string ScriptDescription;
            public string ScriptData;
            public string ScriptFilePath;
            public bool ScriptEnabled;




            //make a builder for first-timers lol!
            public bool LoadScript(string scriptName, string scriptDescription, string scriptData, string scriptFilePath, bool scriptEnabled)
            {
                try
                {
                    ScriptName = scriptName;
                    ScriptDescription = scriptDescription;
                    ScriptData = scriptData;
                    ScriptFilePath = scriptFilePath;
                    ScriptEnabled = scriptEnabled;
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
        public List<CSharpScript> VAR_ListOfScripts = new List<CSharpScript>();
        public List<string> VAR_CSharpCorruptFiles = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        public void initializeListView()
        {
            scriptViewerLV.Items.Clear();
            scriptViewerLV.Columns.Clear();
            scriptViewerLV.View = View.Details;
            ColumnHeader colEnabledCheck = new ColumnHeader();
            colEnabledCheck.Text = "Enabled";

            ColumnHeader colScriptName = new ColumnHeader();
            colScriptName.Text = "Script Name";

            ColumnHeader colScriptDesc = new ColumnHeader();
            colScriptDesc.Text = "Script Description";

            //make some hidden cols for the rest of the data incase we want to use it from the list view...
            ColumnHeader colScriptData = new ColumnHeader();
            colScriptData.Text = "Script Data";

            ColumnHeader colScriptFilePath = new ColumnHeader();
            colScriptFilePath.Text = "ScriptFilePath";

            scriptViewerLV.Columns.Add(colEnabledCheck);
            scriptViewerLV.Columns.Add(colScriptName);
            scriptViewerLV.Columns.Add(colScriptDesc);
            scriptViewerLV.Columns.Add(colScriptData);
            scriptViewerLV.Columns.Add(colScriptFilePath);
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            scriptViewerLV.ItemChecked += new ItemCheckedEventHandler(scriptViewerLV_ItemChecked);
            string[] CSharpScriptPaths = Directory.GetFiles(Application.StartupPath + "\\PretendServerFolder","*.csharp");
            //now instead of using foreach, let's for it up allow for better customization later on...
            //and to be a tad bit memory efficient, I declare the CSharpScript var here and reset it each time
            //in the for loop.
            initializeListView();

            CSharpScript currentLoadedScript = new CSharpScript();
            VAR_ListOfScripts = new List<CSharpScript>();
            for (int scriptCount = 0; scriptCount < CSharpScriptPaths.Length; scriptCount++)
            {
                //reset our CShareScript
                currentLoadedScript = new CSharpScript();
                try
                {
                    currentLoadedScript.ScriptFilePath = CSharpScriptPaths[scriptCount];
                    currentLoadedScript.ScriptData = File.ReadAllText(CSharpScriptPaths[scriptCount]);
                }
                catch
                {
                }
                try
                {
                    currentLoadedScript.ScriptName = currentLoadedScript.ScriptData.Split(new string[] { "[-[" }, StringSplitOptions.None)[1].Split(new string[] { "]-]" }, StringSplitOptions.None)[0].Split('-')[1];
                }
                catch 
                {
                    currentLoadedScript.ScriptName = "N/A";
                }
                try
                {
                    currentLoadedScript.ScriptDescription = currentLoadedScript.ScriptData.Split(new string[] { "[-[" }, StringSplitOptions.None)[2].Split(new string[] {"]-]"}, StringSplitOptions.None)[0].Split('-')[1];
                }
                catch
                {
                    currentLoadedScript.ScriptDescription = "N/A";
                }
                try
                {
                    string tmpEnabledData = currentLoadedScript.ScriptData.Split(new string[] {"[-["},StringSplitOptions.None)[3].Split(new string[] {"]-]"},StringSplitOptions.None)[0].Split('-')[1];
                    if (tmpEnabledData.ToUpper() == "YES")
                    {
                        currentLoadedScript.ScriptEnabled = true;
                    }
                    else if (tmpEnabledData.ToUpper() == "NO")
                    {
                        currentLoadedScript.ScriptEnabled = false;
                    }
                    else
                    {
                        //this file has no enabled tag.. error in some way... and shouldn't be by messagebox.
                        //also, might as well add em to a bad file list for later processing..
                        VAR_CSharpCorruptFiles.Add(currentLoadedScript.ScriptFilePath);
                    }
                   
                }
                catch
                {
                }


                try
                {
                    VAR_ListOfScripts.Add(currentLoadedScript);
                }
                catch
                {
                    VAR_CSharpCorruptFiles.Add(currentLoadedScript.ScriptFilePath);
                }

                
            }
            //now that the loop is over, let's parse into our listview!
            foreach (CSharpScript csScript in VAR_ListOfScripts)
            {
                ListViewItem newScript = new ListViewItem();
                if (csScript.ScriptEnabled == true)
                {
                    newScript.Checked = true;
                }
                else
                {
                    newScript.Checked = false;
                }

                ListViewItem.ListViewSubItem csScriptName = new ListViewItem.ListViewSubItem();
                csScriptName.Text = csScript.ScriptName;

                ListViewItem.ListViewSubItem csScriptDescription = new ListViewItem.ListViewSubItem();
                csScriptDescription.Text = csScript.ScriptDescription;

                ListViewItem.ListViewSubItem csScriptData = new ListViewItem.ListViewSubItem();
                csScriptData.Text = csScript.ScriptData;

                ListViewItem.ListViewSubItem csScriptFilePath = new ListViewItem.ListViewSubItem();
                csScriptFilePath.Text = csScript.ScriptFilePath;

                newScript.SubItems.Clear();
                newScript.SubItems.Add(csScriptName);
                newScript.SubItems.Add(csScriptDescription);
                newScript.SubItems.Add(csScriptData);
                newScript.SubItems.Add(csScriptFilePath);

                scriptViewerLV.Items.Add(newScript);


            }

        }

        void scriptViewerLV_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (e == null)
            {
                return;
            }
            scriptViewerLV.Items[e.Item.Index].Selected = true;
            if (scriptViewerLV.SelectedItems[0].Checked == true)
            {
                testBtn.Enabled = true;
            }
            else
            {
                testBtn.Enabled = false;
            }
        }

        private void testBtn_Click(object sender, EventArgs e)
        {
            ListViewItem scriptToTest = scriptViewerLV.SelectedItems[0];
            System.Collections.Specialized.StringCollection AssemblyReferences = new System.Collections.Specialized.StringCollection();

            System.CodeDom.Compiler.CompilerParameters newParameters = new System.CodeDom.Compiler.CompilerParameters();
            
            //Call any used assembly reference here for the csharp external scripts...
            newParameters.ReferencedAssemblies.Add(@"System.dll");
            newParameters.ReferencedAssemblies.Add(@"System.Windows.Forms.dll");

            newParameters.GenerateInMemory = true;
           

            using (Microsoft.CSharp.CSharpCodeProvider foo =
           new Microsoft.CSharp.CSharpCodeProvider())
            {
                var res = foo.CompileAssemblyFromSource(newParameters,scriptToTest.SubItems[3].Text);
                if (res.Errors.HasErrors == true)
                {
                    MessageBox.Show("This script has " + res.Errors.Count.ToString() + " errors.");
                    foreach (System.CodeDom.Compiler.CompilerError text in res.Errors)
                    {
                        MessageBox.Show(text.ErrorText);
                    }
                   
                    return;
                }
                var type = res.CompiledAssembly.GetType("CSharpScriptClass");

                var obj = Activator.CreateInstance(type);
                var testprice = new PriceComparisonReturnData();
                var returnData = type.GetMethod("TestMethod").Invoke(obj, new object[] { "test" });

                //returnData.GetType().GetFields().GetValue(;
                
                
                System.Reflection.FieldInfo[] testFields = returnData.GetType().GetFields();

                //MessageBox.Show( returnData.GetType().ToString());
                PriceComparisonReturnData newCompare = new PriceComparisonReturnData();
                newCompare = (PriceComparisonReturnData)testFields.GetType().GetProperties();
                MessageBox.Show(test);

                

                
                //TypedReference typeRef = TypedReference.MakeTypedReference(testFields[0].GetValue(
                
                



            }



        }

        private void scriptViewerLV_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (e == null)
            {
                return;
            }
            
            ItemCheckedEventArgs args = e as ItemCheckedEventArgs;
            
            scriptViewerLV_ItemChecked(sender, args);
        }


    }
}
