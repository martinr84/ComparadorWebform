<%@ Page Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Login.aspx.cs" Inherits="ComparadorWebform.Login2" %>



<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server"> 
    <div class="container login-form">
	<h2 class="login-title">- Please Login -</h2>
	<div class="panel panel-default">
		<div class="panel-body">
			<form>
				<div style="text-align:center"><h1 >Login</h1></div>
				<div class="input-group login-userinput">				
					<span class="input-group-addon"><span class="glyphicon glyphicon-user"></span></span>                  
					  <asp:TextBox ID="txtUsuario" runat="server"  class="form-control" ></asp:TextBox>
				</div>
				<div class="input-group">
					<span class="input-group-addon"><span class="glyphicon glyphicon-lock"></span></span>				
					  <asp:TextBox ID="txtPassword" runat="server"  class="form-control" type="password" ></asp:TextBox>
					<span id="showPassword" class="input-group-btn">
            <button class="btn btn-default reveal" type="button"><i class="glyphicon glyphicon-eye-open"></i></button>
          </span>  
				</div>
				<!--button class="btn btn-primary btn-block login-button" type="submit"><i class="fa fa-sign-in"></i> Login</--button!-->
             
				<asp:Button ID="btnLogin"  Width="380px" runat="server" class="btn btn-primary btn-block login-button" Text="Login" OnClick="btnLogin_Click"  />
					
				<div class="checkbox login-options">
					<label><input type="checkbox"/> Remember Me</label>
					<a href="#" class="login-forgot">Forgot Username/Password?</a>
				</div>		
			</form>			
		</div>
	</div>
</div>
</asp:Content>