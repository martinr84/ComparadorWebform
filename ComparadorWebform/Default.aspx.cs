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
          
          //  lblNombreArchivoBanco.Visible = false;
            //lblNombreArchivoOctopus.Visible = false;
            btnDownload.Visible = false;         
            lblBanco.Visible = false;
            lblNombreArchivoBanco0.Visible = false;
            lblNombreArchivoOctopus.Visible = false;
            lblDecimales.Visible = false;
            lblNombreArchivoBanco.Visible = false;
       
           
            }
        

        protected void BtnEjecutarProceso_Click(object sender, EventArgs e)
        {
                     
            bool separadorMilesAmericano;
            string nombreArchivoBanco = string.Empty;
            string nombreArchivoOctopus = string.Empty;
            string input = "~/Archivos/Input/";//poner la ruta donde quieres que quede el archivo
            string output = "~/Archivos/Output/";//poner la ruta donde quieres que quede el archivo
            string carpetaInput = Server.MapPath(input);
            string carpetaOutput = Server.MapPath(output);
            //Borramos los archivos que existen en la carpeta de proceso
            DirectoryInfo dir = new DirectoryInfo(Server.MapPath("~/Archivos/Input"));
            string carpetaDestino = Server.MapPath(output);
            string SaveLocation = "";
            string fecha = DateTime.Now.ToString("dd_MM_yyyy");
            string nombreBanco = ""; 

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            
            if (rbdDecimales.SelectedValue.ToString() == "0")
                separadorMilesAmericano = true;
            else
                separadorMilesAmericano = false;


       

           
            //Subimos archivo banco          
      
            nombreArchivoBanco = System.IO.Path.GetFileName(fuBanco.PostedFile.FileName);
            SaveLocation = carpetaInput + nombreArchivoBanco; 

            fuBanco.PostedFile.SaveAs(SaveLocation);



            //Subimos archivo octopus      
            nombreArchivoOctopus = System.IO.Path.GetFileName(fuOctopus.PostedFile.FileName);
            SaveLocation = carpetaInput + nombreArchivoOctopus;
            fuOctopus.PostedFile.SaveAs(SaveLocation);


            Metodos oMetodos = new Metodos();
            XLWorkbook oWorkbook;
            oWorkbook = oMetodos.ProcesarComparacion(Convert.ToInt32 (ddlBancos.SelectedItem.Value) ,Server.MapPath("~/Archivos/Input/") + nombreArchivoBanco, Server.MapPath("~/Archivos/Input/")
                                        + nombreArchivoOctopus, separadorMilesAmericano);
                     


           
            //Subimos el archivo generado
            output = "~/Archivos/Output/";//poner la ruta donde quieres que quede el archivo


            switch (Convert.ToInt32(ddlBancos.SelectedItem.Value))
            {
                case Bancos.Santander:
                    nombreBanco = "Santander";
                    break;

                case Bancos.ICBC:
                    nombreBanco = "ICBC";
                    break;

                case Bancos.Galicia:
                    nombreBanco = "Galicia";
                    break;


            }


            if (Convert.ToInt32 (rbdDecimales.SelectedItem.Value) == 0)
                lblDecimales.Text = "Signo decimal es el Punto";
            else
                lblDecimales.Text = "Signo decimal es la Coma";




             oWorkbook.SaveAs(carpetaOutput  + "ArchivoComparacion_" + nombreBanco + "_" + fecha   +  ".xlsx");
            
            lblNombreArchivoBanco.Text = nombreArchivoBanco;
            lblNombreArchivoOctopus.Text = nombreArchivoOctopus;
            lblNombreArchivoBanco.Visible = true;
            lblNombreArchivoOctopus.Visible = true;   
            btnDownload.Visible = true;
            //panelInfoProcesada.Visible = true;
            lblNombreArchivoBanco0.Text = "Banco:" +  nombreBanco;     
            lblBanco.Visible = true;
            lblNombreArchivoBanco0.Visible = true;
            lblNombreArchivoOctopus.Visible = true;
            lblDecimales.Visible = true;

        }

        protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void BtnDownload_Click(object sender, EventArgs e)
        {
            //Response.Redirect("~/Archivos/Output/ArchivoComparacionExcel.xlsx");
            string nombreBanco = "";
            switch (Convert.ToInt32(ddlBancos.SelectedItem.Value))
            {
                case Bancos.Santander:
                    nombreBanco = "Santander";
                    break;

                case Bancos.ICBC:
                    nombreBanco = "ICBC";
                    break;

                case Bancos.Galicia:
                    nombreBanco = "Galicia";
                    break;


            }


            Response.Redirect("~/Archivos/Output/ArchivoComparacion_" + nombreBanco + "_" + DateTime.Now.ToString("dd_MM_yyyy") + ".xlsx");

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