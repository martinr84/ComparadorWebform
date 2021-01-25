//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows;
using System.Diagnostics;
using ClosedXML.Excel;
using Spire.Xls;

namespace ComparadorWebform
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

           // ((Label)Master.FindControl("lblUsuario")).Text = HttpContext.Current.User.Identity.Name;

            //    if (!IsPostBack)

            //    fuBanco.Enabled = false;
            //fuOctopus.Enabled = false;
            //btnEjecutarProceso.Enabled = false;


            //   }
            //document.getElementById("MainContent_fuBanco").disabled = true;
            //document.getElementById("MainContent_fuOctopus").disabled = true;
            //document.getElementById("MainContent_btnEjecutarProceso").disabled = true;
            //document.getElementById("MainContent_rbdDecimales_0").disabled = true;
            //document.getElementById("MainContent_rbdDecimales_1").disabled = true;

            //fuBanco.Enabled = false;
            //fuOctopus.Enabled = false;
            //btnEjecutarProceso.Enabled = false;
            //rbdDecimales.Enabled = false;
            lblNombreArchivoBanco.Visible = false;
            lblNombreArchivoOctopus.Visible = false;
            btnDownload.Visible = false;
           
            }
        

        protected void BtnEjecutarProceso_Click(object sender, EventArgs e)
        {      



            ////Preguntamos si se seleccionó si es decimal americano o europeo
            //if (rbdDecimales.SelectedValue == "")
            //{
            //    Response.Write("<script>alert('Debe seleccionar tipo de decimaless');</script>"); 
            //    return;
            //}

            ////Validamos si los archivos ingresados tienen extension .xlsx
            //if (fuBanco.FileName.exte)



            //Borramos los archivos que existen en la carpeta de proceso
            DirectoryInfo dir = new DirectoryInfo(Server.MapPath("~/Archivos/Input"));

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            bool separadorMilesAmericano;
            if (rbdDecimales.SelectedValue.ToString() == "0")
                separadorMilesAmericano = true;
            else
                separadorMilesAmericano = false;


            string nombreArchivoBanco = string.Empty;
            string nombreArchivoOctopus = string.Empty;

            string destino = "~/Archivos/Input/";//poner la ruta donde quieres que quede el archivo

            //Verificamos si hay seleccionado dos archivos
            //Subimos archivo banco
            //if (!fuBanco.HasFile || !fuOctopus.HasFile)
            //{
            //    Response.Write("<script>alert('No se seleccionaron dos archivos');</script>"); 
            //    return;
            //}

            //Subimos archivo banco          
            string carpetaDestino = Server.MapPath(destino);
            nombreArchivoBanco = System.IO.Path.GetFileName(fuBanco.PostedFile.FileName);
            string SaveLocation = carpetaDestino + nombreArchivoBanco;

            fuBanco.PostedFile.SaveAs(SaveLocation);



            //Subimos archivo octopus      
            nombreArchivoOctopus = System.IO.Path.GetFileName(fuOctopus.PostedFile.FileName);
            SaveLocation = carpetaDestino + nombreArchivoOctopus;
            fuOctopus.PostedFile.SaveAs(SaveLocation);


            Metodos oMetodos = new Metodos();
            XLWorkbook oWorkbook;
            oWorkbook = oMetodos.ProcesarComparacion(Server.MapPath("~/Archivos/Input/") + nombreArchivoBanco, Server.MapPath("~/Archivos/Input/")
                                        + nombreArchivoOctopus, separadorMilesAmericano);

          


            // Microsoft.Office.Interop.Excel.Workbook oWorkbook;// = new   Microsoft.Office.Interop.Excel.Workbook();
            //oWorkbook = oMetodos.ProcesarComparacion(Server.MapPath("~/Archivos/Input/") + nombreArchivoBanco, Server.MapPath("~/Archivos/Input/")
            //                          + nombreArchivoOctopus, separadorMilesAmericano);

            // oWorkbook.SaveAs(@"C:\Desarrollos\ComparadorArchivos\ComparadorWebform\ComparadorWebform\Archivos\Output\ArchivoComparacionExcel.xlsx");

            //Subimos el archivo generado
            destino = "~/Archivos/Output/";//poner la ruta donde quieres que quede el archivo
            string carpetaDestinoGenerado = Server.MapPath(destino);
            // nombreArchivoBanco = System.IO.Path.GetFileName(fuBanco.PostedFile.FileName);
            //  SaveLocation = carpetaDestinoGenerado + "ArchivoComparacionExcel.xlsx";

            //fuBanco.PostedFile.SaveAs(SaveLocation);
            oWorkbook.SaveAs(carpetaDestinoGenerado + "ArchivoComparacionExcel.xlsx");
            //oWorkbook.Close();
            //  Response.Write("<script>alert('PROCESO FINALIZADO');</script>");
            lblNombreArchivoBanco.Text = nombreArchivoBanco;
            lblNombreArchivoOctopus.Text = nombreArchivoOctopus;
            lblNombreArchivoBanco.Visible = true;
            lblNombreArchivoOctopus.Visible = true;   
            btnDownload.Visible = true;


            //fuBanco.Enabled = false;
            //fuOctopus.Enabled = false;
            //rbdDecimales.Enabled = false;



            //string path = "Archivos/Output/ArchivoComparacionExcel.xlsx";
            // ClientScript.RegisterStartupScript(this.GetType(), "open", "window.open('" + path + "','_blank', 'fullscreen=yes');", true);




            //ProcessStartInfo psi = new ProcessStartInfo();
            //psi.UseShellExecute = true;
            //psi.LoadUserProfile = true;
            //psi.WorkingDirectory = Server.MapPath("~/Archivos/Output/");// This line solved my problem
            //psi.FileName = Server.MapPath("~/Archivos/Output/ArchivoComparacionExcel.xlsx");
            //psi.Arguments = "Myargument1 Myargument2";
            // Process.Start(psi);

            //  "window.open('MyPDF.pdf', '_blank', 'fullscreen=yes'); return false;"

            // MessageBox.Show("PROCESO TERMINADO");

            //Process.Start(@"C:\Desarrollos\ComparadorArchivos\ComparadorWebform\ComparadorWebform\Archivos\Output\ArchivoComparacionExcel.xlsx");

            //Process.Start(Server.MapPath("~/Archivos/Output/ArchivoComparacionExcel.xlsx"));

        }

        protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void BtnDownload_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Archivos/Output/ArchivoComparacionExcel.xlsx");
           
        }

        protected void btnNuevaComparacion_Click(object sender, EventArgs e)
        {
            fuBanco.Enabled = true;
            fuOctopus.Enabled = true;
            btnEjecutarProceso.Enabled = true;
            rbdDecimales.Enabled = true;
        }

 
    }

  
}