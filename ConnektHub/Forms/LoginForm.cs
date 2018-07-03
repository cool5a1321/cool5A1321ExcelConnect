using Prospecta.ConnektHub.Controllers;
using Prospecta.ConnektHub.Core;
using Prospecta.ConnektHub.Services.User;
using System;
using System.Windows.Forms;

namespace Prospecta.ConnektHub.Forms
{
    public partial class LoginForm : Form
    {
        IUserService _login;
        private string strFullName = string.Empty, strUserName = string.Empty, strPassword = string.Empty;
        private int authenticationResult = 0;

        public string StrUserName
        {
            get { return strUserName; }
            set
            {
                if (value.Equals(strUserName)) { return; }
                strUserName = value;
            }
        }

        public string StrPassword
        {
            get { return strPassword; }
            set
            {
                if (value.Equals(strPassword)) { return; }
                strPassword = value;
            }
        }


        public int AuthenticationResult
        {
            get { return authenticationResult; }
            set
            {
                if (value.Equals(authenticationResult))
                { return; }
                authenticationResult = value;
            }
        }

        public string StrFullName
        {
            get { return strFullName; }
            set
            {
                if (value.Equals(strFullName))
                { return; }
                strFullName = value;
            }
        }

        private void TxtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            { UserLogin(); }
        }

        public LoginForm(IUserService login)
        {
            _login = login;
            InitializeComponent();

            txtUserName.Text = "quality";
            txtPassword.Text = "Welcome123";
        }

        private void BtnSubmit_Click(object sender, EventArgs e)
        {
            UserLogin();    
        }

        private void UserLogin()
        {
            strUserName = txtUserName.Text.Trim();
            strPassword = txtPassword.Text.Trim();

            if (strUserName.Length > 0 && strPassword.Length > 0)
            {
                var userLogin = new UserController(_login);
                authenticationResult = userLogin.UserLogin(strUserName, strPassword, out strFullName);

                if (authenticationResult == 1)
                {
                    var userDetails = new UserDetails
                    {
                        userName = strUserName,
                        password = strPassword,
                        fullName = strFullName
                    };

                    //var retVal = userLogin.AddUserDetails(userDetails);
                    DialogResult = DialogResult.OK;
                }
                else if (authenticationResult == 2)
                {
                    MessageBox.Show("Unable to connect server");
                    DialogResult = DialogResult.Cancel;
                }
                else
                {
                    MessageBox.Show("Please enter correct username and password");
                    DialogResult = DialogResult.Cancel;
                }
            }
            else
            {
                MessageBox.Show("Please enter username and password");
                DialogResult = DialogResult.Cancel;
            }
            this.Close();
            this.Dispose();
            #region Code to be used later
            //System.Data.DataTable dt = sQLiteDatabase.GetDataTable("select * from userdetails");
            //foreach (DataRow dr in dt.Rows)
            //{
            //    foreach (DataColumn dc in dt.Columns)
            //    {
            //        MessageBox.Show(dr[dc].ToString());
            //    }
            //}
            #endregion
        }
    }
}