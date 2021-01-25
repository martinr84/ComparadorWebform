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
          
            lblNombreArchivoBanco.Visible = false;
            lblNombreArchivoOctopus.Visible = false;
            btnDownload.Visible = false;
           
            }
        

        protected void BtnEjecutarProceso_Click(object sender, EventArgs e)
        {      

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
            oWorkbook = oMetodos.ProcesarComparacion(Convert.ToInt32 (ddlBancos.SelectedItem.Value) ,Server.MapPath("~/Archivos/Input/") + nombreArchivoBanco, Server.MapPath("~/Archivos/Input/")
                                        + nombreArchivoOctopus, separadorMilesAmericano);
                     


           
            //Subimos el archivo generado
            destino = "~/Archivos/Output/";//poner la ruta donde quieres que quede el archivo
            string carpetaDestinoGenerado = Server.MapPath(destino);

            oWorkbook.SaveAs(carpetaDestinoGenerado + "ArchivoComparacionExcel.xlsx");
            
            lblNombreArchivoBanco.Text = nombreArchivoBanco;
            lblNombreArchivoOctopus.Text = nombreArchivoOctopus;
            lblNombreArchivoBanco.Visible = true;
            lblNombreArchivoOctopus.Visible = true;   
            btnDownload.Visible = true;



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