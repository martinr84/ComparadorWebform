using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ComparadorWebform
{
    public partial class Login2 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {


        }

      
        protected void btnLogin_Click(object sender, EventArgs e)
        {
            // Authenticate againts the list stored in web.config

            //if (FormsAuthentication.Authenticate(txtUserName.Text, txtPassword.Text))
            //{

            //    FormsAuthentication.RedirectFromLoginPage(txtUserName.Text, false);
            //}

            if (1 == 1)

            {
                FormsAuthentication.RedirectFromLoginPage(txtUsuario.Text, false);
            }




            else
            {
                // lblMessage.Text = "Invalid UserName and/or password";
            }
        }
    }
}