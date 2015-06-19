using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace UniversalXMLParser
{
    public partial class Form1 : Form
    {
        public struct XMLTag
        {
            public string XMLTagParent;
            public string XMLTagName;
            public string XMLTagValue;
        }
        public struct XMLObject
        {
            public string XMLObjectName;
            public List<XMLTag> XMLTags;
        }
        public struct XMLDocument
        {
            public string XMLVersion;
            public string XMLEncoding;
            public List<XMLObject> XMLObjects;
        }
        public XMLDocument CreateXMLDocument(string XML_Input)
        {
             XMLDocument newXML = new XMLDocument();
             XMLObject newObject = new XMLObject();
             XMLTag newTag = new XMLTag();

            foreach(string XMLObj in XML_Input.Split(new string[] {"\n<"},StringSplitOptions.None))
            {
                if (XMLObj[1] != '?')
                {
                    string objectTag = "<" + XMLObj.Split('>')[0] + ">";
                    string endObjectTag = "</" + XMLObj.Split('>')[0] + ">";
                    string xmlObject = XML_Input.Split(new string[]{objectTag},StringSplitOptions.None)[1].Split(new string[] {endObjectTag},StringSplitOptions.None)[0];
                    newObject.XMLObjectName = xmlObject.Split('>')[2].Split('<')[0];
                    
                    
                    //now parse the xml tags of this object...
                    foreach (string xmlTag in xmlObject.Split(new string[] {" <"},StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (xmlTag != "\r\n  ")
                        {
                            newTag.XMLTagValue = xmlTag.Split('>')[1].Split('<')[0];
                            newTag.XMLTagParent = newObject.XMLObjectName;
                            newTag.XMLTagName = xmlTag.Split('>')[0];
                            newObject.XMLTags.Add(newTag);
                        }
                    }
                    newXML.XMLObjects.Add(newObject);


                }
                else
                {
                    //header info.. get version and encoding data.
                    string version = XMLObj.Split(new string[] {"version=\""},StringSplitOptions.None)[1].Split('"')[0];
                    string encoding = XMLObj.Split(new string[] {"encoding=\""},StringSplitOptions.None)[1].Split('"')[0];
                    newXML.XMLVersion = version;
                    newXML.XMLEncoding = encoding;

                }
            }

            return newXML;
        }


        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //XMLDocument testDoc = CreateXMLDocument(textBox1.Text);
            XmlDocument newDoc = new XmlDocument();
            newDoc.LoadXml(textBox1.Text);
            //MessageBox.Show(newDoc.ChildNodes.Count.ToString());
            foreach (XmlDeclaration node in newDoc.ChildNodes)
            {
                MessageBox.Show("this node has " + node.ChildNodes.Count.ToString() + " childs.");
                for (int x = 0; x < node.ChildNodes.Count; x++)
                {
                    MessageBox.Show("Node " + x.ToString() + "'s name is " + node.ChildNodes[x].Name + " and has a value of "+ node.ChildNodes[x].Value);
                   
                }
            }
            //MessageBox.Show(nodeList.Count.ToString());
        }


        
    }
}
