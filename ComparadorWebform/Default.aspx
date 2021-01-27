<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ComparadorWebform._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div>
        <h1>COMPARADOR ARCHIVOS</h1>
        
    </div>

    <div class="row">
        <div class="col-md-4">
        </div>
       
     </div>
        
   


    <br />
    <br />
    <br />
    Banco<br />
    <asp:DropDownList ID="ddlBancos" runat="server" CssClass="btn btn-primary">
        <asp:ListItem Selected="True" Value="0">Banco</asp:ListItem>
        <asp:ListItem Value="1">Santander</asp:ListItem>
        <asp:ListItem Value="2">ICBC</asp:ListItem>
        <asp:ListItem Value="3">Galicia</asp:ListItem>
    </asp:DropDownList>
    <br />
    <br />
    <br />
    <asp:Label ID="lblBanco" runat="server" Text="Archivo Banco"></asp:Label>
    <br />
    <asp:FileUpload ID="fuBanco" runat="server"  class="btn btn-primary"/>    
    <br />
    <asp:Label ID="lblOctopus" runat="server" Text="Archivo Octopus"></asp:Label>
    <br />
    <asp:FileUpload ID="fuOctopus" runat="server" class="btn btn-primary" />
    <br />
    <br />
    <asp:RadioButtonList ID="rbdDecimales" runat="server" OnSelectedIndexChanged="RadioButtonList1_SelectedIndexChanged">
        <asp:ListItem Value="0">punto para decimales</asp:ListItem>
        <asp:ListItem Value="1">coma para decimales</asp:ListItem>
    </asp:RadioButtonList>
    <br />
    
  
    
<table>
  <tr>
      <td>
   <asp:Button ID="btnEjecutarProceso" runat="server" OnClick="BtnEjecutarProceso_Click" Text="Ejecutar proceso"   CssClass="btn btn-primary" Height="71px" OnClientClick= "return validarIngresos()" Width="246px"  />
</td>        
      <td>               &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;</td>
    <td >     
        &nbsp;&nbsp;&nbsp;<asp:Label ID="lblNombreArchivoBanco0" runat="server" Font-Size="Small" Text="lblBanco"></asp:Label>
                   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     
        <asp:Panel ID="panelInfoProcesada" runat="server" GroupingText=" " Width="200px">
                   &nbsp;&nbsp;&nbsp;<asp:Label ID="lblNombreArchivoBanco" runat="server" Font-Size="Small" Text="lblBanco"></asp:Label>
                   <br />
                   <br />
                   &nbsp;
                   <asp:Label ID="lblNombreArchivoOctopus" runat="server" Font-Size="Small" Text="lblOctopus"></asp:Label>
                     </asp:Panel>
    </td>
    
  </tr> 

</table>                

    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                

    <asp:Button ID="btnDownload" runat="server" Text="Descargar Archivo"  CssClass="btn btn-success"
            OnClick="BtnDownload_Click" Width="800" Height="89px" Visible="False"/>
   
    



    


</asp:Content>
