using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TrueValueEmployeeManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void refreshListLnk_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RefreshEmployeeList();
        }

        private void RefreshEmployeeList()
        {
            employeesLst.Items.Clear();
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT first_name + ' ' + last_name as Name FROM TrueValueContacts");
            if (qr.Rows.Count > 0)
            {
                foreach (string row in qr.Rows)
                {
                    employeesLst.Items.Add(row.Split('|')[0]);
                }
            }

        }

        private void EditSelectedEmployee(string EmployeeFullName)
        {
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT id,rdc_id,job_title,first_name,last_name,office_phone_ext,phone_number,cell_phone,active,photo_path FROM TrueValueContacts WHERE first_name + ' ' + last_name='" + EmployeeFullName.Replace("'","''") + "'");
            if (qr.Rows.Count > 0)
            {
                foreach (string row in qr.Rows)
                {
                    databaseIDTxt.Text = "";
                    string db_id = row.Split('|')[0];
                    string rdc_id = row.Split('|')[1];
                    string job_title = row.Split('|')[2];
                    string first_name = row.Split('|')[3];
                    string last_name = row.Split('|')[4];
                    string office_phone_ext = row.Split('|')[5];
                    string phone_number = row.Split('|')[6];
                    string cell_phone = row.Split('|')[7];
                    string active = row.Split('|')[8];
                    string photo_path = row.Split('|')[9];

                    

                    databaseIDTxt.Text = db_id;
                    rdcIDTxt.Text = rdc_id;
                    jobTitleTxt.Text = job_title;
                    firstNameTxt.Text = first_name;
                    lastNameTxt.Text = last_name;
                    officeExtTxt.Text = office_phone_ext;
                    phoneNumberTxt.Text = phone_number;
                    cellPhoneTxt.Text = cell_phone;
                    activeChk.Checked = (active == "True" ? true : false);

                    if (photo_path.Length > 0)
                    {
                        imagePathLbl.Text = photo_path;
                        previewPhotoImg.Image = Image.FromFile(photo_path);
                        previewPhotoImg.Refresh();
                    }
                    else
                    {
                        imagePathLbl.Text = "...";
                    }

                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RefreshEmployeeList();
        }

        private void employeesLst_SelectedIndexChanged(object sender, EventArgs e)
        {
            previewPhotoImg.Image = null;
            string EmployeeFullName = employeesLst.Items[employeesLst.SelectedIndex].ToString();
            EditSelectedEmployee(EmployeeFullName);
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            string photoPath = "";
            if (imagePathLbl.Text.Length > 3)
            {
                photoPath = imagePathLbl.Text;
            }

            if (databaseIDTxt.Text.Length > 0)
            {
                //update
                Connectivity.SQLCalls.sqlQuery("UPDATE TrueValueContacts SET rdc_id=" + rdcIDTxt.Text.Replace("'", "''") + ",job_title='" + jobTitleTxt.Text.Replace("'", "''") + "', first_name='" + firstNameTxt.Text.Replace("'", "''") + "', last_name='" + lastNameTxt.Text.Replace("'", "''") + "', office_phone_ext='" + officeExtTxt.Text.Replace("'", "''") + "',phone_number='" + phoneNumberTxt.Text.Replace("'", "''") + "',cell_phone='" + cellPhoneTxt.Text.Replace("'", "''") + "',active=" + (activeChk.Checked == true ? "1" : "0") + ", photo_path='" + photoPath  + "' WHERE id=" + databaseIDTxt.Text);
            }
            else
            {
                //add new
                Connectivity.SQLCalls.sqlQuery("INSERT INTO TrueValueContacts (rdc_id,job_title,first_name,last_name,office_phone_ext,phone_number,cell_phone,active,photo_path) VALUES(" + rdcIDTxt.Text.Replace("'", "''") + ",'" + jobTitleTxt.Text.Replace("'", "''") + "','" + firstNameTxt.Text.Replace("'", "''") + "','" + lastNameTxt.Text.Replace("'", "''") + "','" + officeExtTxt.Text.Replace("'", "''") + "','" + phoneNumberTxt.Text.Replace("'", "''") + "','" + cellPhoneTxt.Text.Replace("'", "''") + "'," + (activeChk.Checked == true ? "1" : "0") + ",'" + photoPath + "')");

            }

            //if (imagePathLbl.Text.Length > 3)
            //{
            //    Connectivity.SQLCalls.sqlQuery("UPDATE TrueValueContacts SET photo=SELECT * FROM OPENROWSET(BULK N'S:\\mastest\\mas90-signs\\A_Dist\\toms beta\\TrueValueImages\\peoples\\PaulMotylinski.png', SINGLE_BLOB) as imagefile WHERE first_name='" + firstNameTxt.Text + "' AND last_name='" + lastNameTxt.Text + "' AND rdc_id=" + rdcIDTxt.Text);
            //}



        }

        private void clearFormLnk_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            databaseIDTxt.Text = "";
            rdcIDTxt.Text = "";
            jobTitleTxt.Text = "";
            firstNameTxt.Text = "";
            lastNameTxt.Text = "";
            officeExtTxt.Text = "";
            phoneNumberTxt.Text = "";
            cellPhoneTxt.Text = "";
            activeChk.Checked = false;
            imagePathLbl.Text = "...";
            previewPhotoImg.Image = null;
        }

        private void browseLnk_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFileDialog1.Filter = "*.png|*.png";
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK | dr == System.Windows.Forms.DialogResult.Yes)
            {
                imagePathLbl.Text = openFileDialog1.FileName;
                previewPhotoImg.Image = Image.FromFile(openFileDialog1.FileName);

            }

        }
    }
}
