using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Xml;
using System.DirectoryServices;

namespace ReportGenerator
{
    public static class UserSecurity
    {
        private static List<string> UserAllowedModules = new List<string>();
        private static bool IsLoggedIn = false;
        private static bool isUserElevated = false;
        private static Form loginFrm;

        private static void GetUserPermissionsModuleList(int userID)
        {
            List<string> tmp = Connectivity.SQLCalls.sqlQuery("SELECT ModuleID FROM UserModulePermissions WHERE PersonID=" + userID.ToString());
            foreach (string t in tmp)
            {
                UserAllowedModules.Add(t.Split('|')[0]);
            }
        }
        private static List<string> GetAllUsers()
        {
            //return Connectivity.SQLCalls.sqlQuery("SELECT CAST(COALESCE(ID,'') AS VARCHAR(30)) + ':' + COALESCE(NTUsername,'') AS 'UserList' FROM Users WHERE Deleted=0 AND SpecialUser=0");
            List<string> finalUsers = new List<string>();
            List<string> Users =  Connectivity.SQLCalls.sqlQuery("SELECT NTUsername FROM Users WHERE Deleted=0 AND SpecialUser=0");
            foreach (string user in Users)
            {
                finalUsers.Add(user.Split('|')[0]);
            }
            return finalUsers;
        }
        private static int GetUserID(string Username)
        {
            List<string> ID = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM Users WHERE NTUsername='" + Username + "'");
            if (ID.Count > 0)
            {
                return int.Parse(ID[0].Split('|')[0]);
            }
            else
            {
                return -999;
            }
        }
        private static void LogUserIn(object sender, EventArgs e)
        {
            
            IsLoggedIn = false;
            string Username = loginFrm.Controls[3].Text;
            string Password = loginFrm.Controls[4].Text;
            bool auth = false;
            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://TOOLBOX", Username, Password);
                object nativeObject = entry.NativeObject;
                auth = true;
            }
            catch (DirectoryServicesCOMException f)
            {
                string res = f.Message.ToString();
            }

            if (auth == true)
            {
                GetUserPermissionsModuleList(GetUserID(Username));
                foreach (string permission in UserAllowedModules)
                {
                    if (permission == "35") //35 is the Report Generator access ID
                    {
                        IsLoggedIn = true;
                        isUserElevated = true;
                        break;
                    }
                }
                if (isUserElevated == false)
                {
                    //isUserElevated = false;
                    MessageBox.Show("You do not have access to the Report Generator. If you believe this is in error tell Tom.");
                }
            }
            else
            {
                IsLoggedIn = false;
                MessageBox.Show("The Username/Password supplied does not match!");
                loginFrm.Controls[3].Text = "";
                loginFrm.Controls[4].Text = "";
            }

 

        }
        private static void CancelClick(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private static void FormExit(object sender, FormClosingEventArgs e)
        {
            if (IsLoggedIn == false)
            {
                Application.Exit();
            }
        }
        private static void PasswordKeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                LogUserIn(loginFrm, (EventArgs)e);
            }
        }

        public static bool IsUserLoggedIn()
        {
            return IsLoggedIn;
        }
        public static List<string> GetAllowedModules()
        {
            return UserAllowedModules;
        }        
        public static void LoginFrm()
        {
            loginFrm = new Form();
            loginFrm.Width = 350;
            loginFrm.Height = 200;
            loginFrm.Text = "Login";
            loginFrm.FormBorderStyle = FormBorderStyle.FixedSingle;
            loginFrm.MaximizeBox = false;
            loginFrm.MinimizeBox = false;            
            
            Label titleLbl = new Label();
            Label userLbl = new Label();            
            Label passLbl = new Label();
            ComboBox userCmb = new ComboBox();
            TextBox passTxt = new TextBox();
            Button loginBtn = new Button();
            Button cancelBtn = new Button();
            
            titleLbl.Text = "Sign In";
            titleLbl.AutoSize = true;
            titleLbl.Left = 20;
            titleLbl.Top = 10;
            titleLbl.Font = new System.Drawing.Font(System.Drawing.FontFamily.GenericSansSerif, 16);            
            
            userLbl.Text = "Username:";
            userLbl.AutoSize = true;
            userLbl.Left = 20;
            userLbl.Top = 50;

            passLbl.Text = "Password:";
            passLbl.AutoSize = true;
            passLbl.Left = 20;
            passLbl.Top = 75;

            userCmb.Text = "Select User";
            userCmb.Items.AddRange(GetAllUsers().ToArray());
            userCmb.Width = 150;
            userCmb.Height = 20;
            userCmb.Left = userLbl.Right;
            userCmb.Top = userLbl.Top;
            userCmb.AutoCompleteSource = AutoCompleteSource.ListItems;
            userCmb.AutoCompleteMode = AutoCompleteMode.Append; //not working???

            passTxt.Text = "";
            passTxt.PasswordChar = '*';
            passTxt.Width = 150;
            passTxt.Height = 20;
            passTxt.Left = passLbl.Right;
            passTxt.Top = passLbl.Top;

            loginBtn.Text = "Sign In";
            loginBtn.Left = userCmb.Left;            
            loginBtn.Top = passTxt.Bottom + 10;
            loginBtn.AutoSize = true;

            cancelBtn.Text = "Cancel";
            cancelBtn.Left = loginBtn.Right + 15;
            cancelBtn.Top = loginBtn.Top;
            cancelBtn.AutoSize = true;

            loginFrm.Controls.Add(titleLbl);
            loginFrm.Controls.Add(userLbl);
            loginFrm.Controls.Add(passLbl);
            loginFrm.Controls.Add(userCmb);
            loginFrm.Controls.Add(passTxt);
            loginFrm.Controls.Add(loginBtn);
            loginFrm.Controls.Add(cancelBtn);

            loginBtn.Click+=new EventHandler(LogUserIn);
            cancelBtn.Click +=new EventHandler(CancelClick);
            passTxt.KeyPress += new KeyPressEventHandler(PasswordKeyPress);
            loginFrm.FormClosing += new FormClosingEventHandler(FormExit);

            loginFrm.Show();
            loginFrm.BringToFront();
        }
        public static void CloseLoginFrm()
        {
            loginFrm.Close();
        }
        public static void LoginBringToFront()
        {
            loginFrm.BringToFront();
        }


    }
}
