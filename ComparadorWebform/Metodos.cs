using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using ClosedXML.Excel;
namespace ComparadorWebform
{
    public class Metodos
    {

        public bool ValidarDosArchivos()
        {
            DirectoryInfo dir = new System.IO.DirectoryInfo(@"C:\Desarrollos\ComparadorArchivos\ComparadorArchivos\Input");
            int count = dir.GetFiles().Length;

            if (count != 2)
                return false;
            return true;
        }


        public XLWorkbook ProcesarComparacion(int banco,string archivoBanco, string archivoOctopus, bool separadorMilesAmericano)
        {
            string[,] infoBanco = new string[500, 4];
            string[,] infoOctopus = new string[500, 5];
            string[,] infoBancoNoEnOctopus = new string[500, 4];
            string[,] infoOctopusNoEnBanco = new string[500, 4];
            decimal[] sumatoriaCodigosBancoExcluidos = new decimal[6];
            XLWorkbook oWorkbook;
            // var rows = XLWoroWorkbookkbook.worWorksheet(1);
            // bool separadorMilesAmericano = false;

           // decimal sumatoria = 0;

            //FileInfo oFile = new FileInfo(@"C:\Desarrollos\ComparadorArchivos\ComparadorArchivos\Input\ImputBanco2.xls");
            FileInfo oFile = new FileInfo(archivoBanco);




          

            //ProcesarExcelBancoSantander(infoBanco, separadorMilesAmericano, archivoBanco, separadorMilesAmericano);
            ProcesarExcelBancos(banco, infoBanco, separadorMilesAmericano, archivoBanco, separadorMilesAmericano);
            ProcesarExcelOctopus(infoOctopus, archivoOctopus);
            ProcesarDiferenciasBancoNoEnOctopus(banco, infoBanco, infoOctopus, infoBancoNoEnOctopus);
            ProcesarDiferenciasOctopusNoEnBanco(infoBanco, infoOctopus, infoOctopusNoEnBanco);
            //sumatoriaCodigosBancoExcluidos = CalcularSumatoriaCodigosExcluidosBancoSantander(infoBanco);

            sumatoriaCodigosBancoExcluidos = CalcularSumatoriaCodigosExcluidosBancos(banco, infoBanco);

           // GenerarReporteTXT(infoBancoNoEnOctopus, infoOctopusNoEnBanco, sumatoriaCodigosBancoExcluidos);
            oWorkbook = GenerarReporteExcel(banco, infoBancoNoEnOctopus, infoOctopusNoEnBanco, sumatoriaCodigosBancoExcluidos);


            return oWorkbook;

        }

        public void ProcesarExcelBancos(int banco, string[,] infoBanco, bool separadorMilesAmericanostring, string archivoBanco, bool separadorMilesAmericano)
        {
            //Pregunto de que banco es el archivo banco

            switch (banco)
            {
                case Bancos.Santander:
                    ProcesarExcelBancoSantander(infoBanco, separadorMilesAmericano, archivoBanco, separadorMilesAmericano);
                    break;
                case Bancos.ICBC:
                    ProcesarExcelBancoICBC(infoBanco, separadorMilesAmericano, archivoBanco, separadorMilesAmericano);
                    break;

                case Bancos.Galicia:
                    ProcesarExcelBancoGalicia(infoBanco, separadorMilesAmericano, archivoBanco, separadorMilesAmericano);
                    break;

            }

        }


        public string[,] ProcesarExcelBancoSantander(string[,] infoBanco, bool separadorMilesAmericanostring, string archivoBanco, bool separadorMilesAmericano)
        {         

            var oWorkbook = new XLWorkbook(archivoBanco);
            var oWorksheet = oWorkbook.Worksheet(1);          


             int j = 0;
            
            //Si luego de recorrer el excel  encuentra una fila con columna 6 vacía sale
            //ANDANDO PERFECTO CON LOGIN
            // while (!string.IsNullOrEmpty(oWorksheet.Cell(i + 6, 6).Value.ToString()))
            for  (  int i =5; i < 200; i++)
            {

                int test;
                bool esNumerico;
                esNumerico = int.TryParse(oWorksheet.Cell(i, 4).Value.ToString(), out test);


                //if (!string.IsNullOrEmpty(oWorksheet.Cell( i , 4).Value.ToString()))
                if (!string.IsNullOrEmpty(oWorksheet.Cell(i, 4).Value.ToString()) && esNumerico == true)
                { 
                                    
                                    
                                    
                                    
                                    //andaba martin fallando en somee i
                                    //string sDate = oWorksheet.Cell(i + 6, 1).Value.ToString();
                                    //infoBanco[i, 0] = DateTime.Parse(sDate).ToString("dd/MM/yyyy");
                                    
                                    
                                    //ANDANDO PERFECTO CON LOGIN
                                    infoBanco[j, 0] =   oWorksheet.Cell(i , 1).Value.ToString();
                                    
                                    //codigo
                                    
                                    double codigo = Convert.ToDouble(oWorksheet.Cell(i , 4).Value.ToString());
                                    infoBanco[j, 1] = codigo.ToString();
                                    
                                    
                                    //concepto
                                    infoBanco[j, 2] = oWorksheet.Cell(i , 6).Value.ToString();
                                    
                                    //importe
                                    double importeAmericano;
                                    string importeString;
                                    if (separadorMilesAmericano == true)
                                    {
                                        importeAmericano = Convert.ToDouble(oWorksheet.Cell(i , 7).Value.ToString());
                                        infoBanco[j, 3] = importeAmericano.ToString();
                                    }
                                    
                                    
                                    else
                                    {
                                        importeString = oWorksheet.Cell(i , 7).Value.ToString();
                                    
                                        if (importeString.Contains("("))
                                        {
                                            //importeString.Replace("(", "").Replace(")","");
                                            importeString.Replace("(", "");
                                    
                                            importeString = importeString.Replace("(", "").Replace(")", "");
                                            importeString = "-" + importeString;
                                    
                                        }
                                        //Si es importe es formato americano (coma separa miles y punto decimales)
                                        if (separadorMilesAmericano == false)
                                        {
                                            //ESTO ES CLAVE EN IIS NO HAY QUE PONERLO PERO PARA DESARROLLO SI
                                            //importeString = importeString.Replace(".", "").Replace(",", ".");
                                        }
                                    
                                    
                                    
                                        decimal sacarDecimalesVacios;
                                        sacarDecimalesVacios = Convert.ToDecimal(importeString) / 1.00m;
                                    
                                        infoBanco[j, 3] = sacarDecimalesVacios.ToString();
                                    
                                    
                                    }
                    j++;
                }//end if 

               // i=i+1;

            }//end for

            //xlWorkBook.Close(true, null, null);
            //xlApp.Quit();
            return infoBanco;

        }

        
        public string[,] ProcesarExcelBancoICBC(string[,] infoBanco, bool separadorMilesAmericanostring, string archivoBanco, bool separadorMilesAmericano)
        {

            var oWorkbook = new XLWorkbook(archivoBanco);
            var oWorksheet = oWorkbook.Worksheet(1);


            int j = 0;

            //Si luego de recorrer el excel  encuentra una fila con columna 6 vacía sale
            //ANDANDO PERFECTO CON LOGIN
            // while (!string.IsNullOrEmpty(oWorksheet.Cell(i + 6, 6).Value.ToString()))
            for (int i = 4; i < 200; i++)
            {

                int test;
                bool esNumerico;
                esNumerico = int.TryParse(oWorksheet.Cell(i, 2).Value.ToString(), out test);


                //if (!string.IsNullOrEmpty(oWorksheet.Cell( i , 4).Value.ToString()))
                if (!string.IsNullOrEmpty(oWorksheet.Cell(i, 2).Value.ToString()) && esNumerico == true)
                {




                    //andaba martin fallando en somee i
                    //string sDate = oWorksheet.Cell(i + 6, 1).Value.ToString();
                    //infoBanco[i, 0] = DateTime.Parse(sDate).ToString("dd/MM/yyyy");


                    //ANDANDO PERFECTO CON LOGIN
                    infoBanco[j, 0] = oWorksheet.Cell(i, 1).Value.ToString();

                    //codigo

                    double codigo = Convert.ToDouble(oWorksheet.Cell(i, 2).Value.ToString());
                    infoBanco[j, 1] = codigo.ToString();


                    //concepto
                    infoBanco[j, 2] = oWorksheet.Cell(i, 3).Value.ToString();

                    //importe
                    double importeAmericano;
                    string importeString;

                    if (oWorksheet.Cell(i, 4).Value.ToString() == string.Empty)
                        importeString = oWorksheet.Cell(i, 5).Value.ToString();
                    else
                        importeString = oWorksheet.Cell(i, 4).Value.ToString();


                    if (separadorMilesAmericano == true)
                    {
                        importeAmericano = Convert.ToDouble(importeString);
                        infoBanco[j, 3] = importeAmericano.ToString();
                    }


                    else
                    {
                        //importeString = oWorksheet.Cell(i, 7).Value.ToString();

                        if (importeString.Contains("("))
                        {
                            //importeString.Replace("(", "").Replace(")","");
                            importeString.Replace("(", "");

                            importeString = importeString.Replace("(", "").Replace(")", "");
                            importeString = "-" + importeString;

                        }
                        //Si es importe es formato americano (coma separa miles y punto decimales)
                        if (separadorMilesAmericano == false)
                        {
                            //ESTO ES CLAVE EN IIS NO HAY QUE PONERLO PERO PARA DESARROLLO SI
                            //importeString = importeString.Replace(".", "").Replace(",", ".");
                        }



                        decimal sacarDecimalesVacios;
                        sacarDecimalesVacios = Convert.ToDecimal(importeString) / 1.00m;

                        infoBanco[j, 3] = sacarDecimalesVacios.ToString();


                    }
                    j++;
                }//end if 

                // i=i+1;

            }//end for

            //xlWorkBook.Close(true, null, null);
            //xlApp.Quit();
            return infoBanco;

        }



        public string[,] ProcesarExcelBancoGalicia(string[,] infoBanco, bool separadorMilesAmericanostring, string archivoBanco, bool separadorMilesAmericano)
        {

            var oWorkbook = new XLWorkbook(archivoBanco);
            var oWorksheet = oWorkbook.Worksheet(1);

            int j = 0;

         
            for (int i = 2; i < 200; i++)
            {

                //int test;
                //bool esNumerico;
                //esNumerico = int.TryParse(oWorksheet.Cell(i, 2).Value.ToString(), out test);


         
                if (!string.IsNullOrEmpty(oWorksheet.Cell(i, 1).Value.ToString()) )
                {
                    infoBanco[j, 0] = oWorksheet.Cell(i, 1).Value.ToString();

                    //codigo
                    //Galicia no tiene código
                    //double codigo = Convert.ToDouble(oWorksheet.Cell(i, 2).Value.ToString());
                    //infoBanco[j, 1] = codigo.ToString();


                    //concepto
                    infoBanco[j, 2] = oWorksheet.Cell(i, 2).Value.ToString();

                    //importe
                    double importeAmericano;
                    string importeString;

                    if (oWorksheet.Cell(i, 4).Value.ToString() == "0")
                        importeString = oWorksheet.Cell(i, 5).Value.ToString();
                    else
                        importeString = "-" + oWorksheet.Cell(i, 4).Value.ToString();


                    if (separadorMilesAmericano == true)
                    {
                        importeAmericano = Convert.ToDouble(importeString);
                        infoBanco[j, 3] = importeAmericano.ToString();
                    }


                    else
                    {
                        //importeString = oWorksheet.Cell(i, 7).Value.ToString();

                        if (importeString.Contains("("))
                        {
                            //importeString.Replace("(", "").Replace(")","");
                            importeString.Replace("(", "");

                            importeString = importeString.Replace("(", "").Replace(")", "");
                            importeString = "-" + importeString;

                        }
                        //Si es importe es formato americano (coma separa miles y punto decimales)
                        if (separadorMilesAmericano == false)
                        {
                            //ESTO ES CLAVE EN IIS NO HAY QUE PONERLO PERO PARA DESARROLLO SI
                            //importeString = importeString.Replace(".", "").Replace(",", ".");
                        }



                        decimal sacarDecimalesVacios;
                        sacarDecimalesVacios = Convert.ToDecimal(importeString) / 1.00m;

                        infoBanco[j, 3] = sacarDecimalesVacios.ToString();


                    }
                    j++;
                }//end if 

                // i=i+1;

            }//end for

            //xlWorkBook.Close(true, null, null);
            //xlApp.Quit();
            return infoBanco;

        }


        public string[,] ProcesarExcelOctopus(string[,] infoOctopus, string archivoOctopus)
        {
            // Application xlApp;
            //  XLWorkbook oWorkbook;
            //worksheet xlWorkSheet;
            //  Range range;
            double debe, haber;

            // int cantidadFilas = 0;
            // int cl = 0;

            //  xlApp = new Application();
            //xlWorkBook = xlApp.Workbooks.Open(@"C:\Desarrollos\ComparadorArchivos\ComparadorArchivos\Input\ImputOctopus2.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            // xlWorkBook = xlApp.Workbooks.Open(archivoOctopus, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            var oWorkbook = new XLWorkbook(archivoOctopus);
            var oWorksheet = oWorkbook.Worksheet(1);




            // range = xlWorkSheet.UsedRange;
            //cantidadFilas = range.Rows.Count;
            //cl = range.Columns.Count;


            int i = 0;
            //Si luego de recorrer el excel  encuentra una fila con columna 6 vacía sale
            while (!string.IsNullOrEmpty(oWorksheet.Cell(i + 6, 2).Value.ToString()))
            {

                string sDate = oWorksheet.Cell(i + 6, 1).Value.ToString();
                infoOctopus[i, 0] = sDate;


                //concepto
                infoOctopus[i, 1] = oWorksheet.Cell(i + 6, 2).Value.ToString();

                //debe              
                debe = Convert.ToDouble(oWorksheet.Cell(i + 6, 4).Value.ToString());
                //((infoOctopus[i, 3] = debe.ToString();


                infoOctopus[i, 2] = (string.IsNullOrEmpty(oWorksheet.Cell(i + 6, 4).Value.ToString())) ? "0" : debe.ToString();



                //haber              
                haber = Convert.ToDouble(oWorksheet.Cell(i + 6, 5).Value.ToString());
                //infoOctopus[i, 5] = haber.ToString();
                infoOctopus[i, 3] = (string.IsNullOrEmpty(oWorksheet.Cell(i + 6, 5).Value.ToString())) ? "0" : haber.ToString();





                //sumaDebeHaner
                infoOctopus[i, 4] = (debe + haber).ToString();

                //martin prueba decimales
                //decimal haber2;
                //haber2 = Convert.ToDecimal(infoOctopus[i, 4]);



                i++;

            }

            //   xlWorkBook.Close(true, null, null);
            // xlApp.Quit();
            return infoOctopus;

        }


        public void ProcesarDiferenciasBancoNoEnOctopus(int banco, string[,] infoBanco, string[,] infoOctopus, string[,] infoBancoNoEnOctopus)
        {
            switch (banco)
            {
                case Bancos.Santander:
                    ProcesarDiferenciasBancoSantanderNoEnOctopus(banco, infoBanco, infoOctopus, infoBancoNoEnOctopus);
                    break;
                case Bancos.ICBC:
                    ProcesarDiferenciasBancoICBCNoEnOctopus(banco, infoBanco, infoOctopus, infoBancoNoEnOctopus);
                    break;

                case Bancos.Galicia:
                    ProcesarDiferenciasBancoGaliciaNoEnOctopus(banco, infoBanco, infoOctopus, infoBancoNoEnOctopus);
                    break;

            }


          
        }
        public string[,] ProcesarDiferenciasBancoSantanderNoEnOctopus(int banco, string[,] infoBanco, string[,] infoOctopus, string[,] infoBancoNoEnOctopus)
        {
            int i = 0;
            int t = 0;



            while (!string.IsNullOrEmpty(infoBanco[i, 0]))
            {
                string montoActual;
                bool montoEncontrado = false;

                int j = 0;

                while (!string.IsNullOrEmpty(infoOctopus[j, 0]))
                {
                    if (!infoBanco[i, 1].Contains("4633") && !infoBanco[i, 1].Contains("4637") && !infoBanco[i, 1].Contains("3254")
                      || !infoBanco[i, 1].Contains("1924") && !infoBanco[i, 1].Contains("2960"))
                    {
                        if (infoBanco[i, 3] == infoOctopus[j, 4])
                        {
                            montoEncontrado = true;
                            break;
                        }
                    }
                    j++;

                }
                if (montoEncontrado == false
                       && !infoBanco[i, 1].Contains("4633") && !infoBanco[i, 1].Contains("4637") && !infoBanco[i, 1].Contains("3254")
                      && !infoBanco[i, 1].Contains("1924") && !infoBanco[i, 1].Contains("2960"))
                {
                    infoBancoNoEnOctopus[t, 0] = infoBanco[i, 0];
                    infoBancoNoEnOctopus[t, 1] = infoBanco[i, 1];
                    infoBancoNoEnOctopus[t, 2] = infoBanco[i, 2];
                    infoBancoNoEnOctopus[t, 3] = infoBanco[i, 3];
                    t++;
                }
                i++;
            }
            return infoBancoNoEnOctopus;
        }



        public string[,] ProcesarDiferenciasBancoICBCNoEnOctopus(int banco, string[,] infoBanco, string[,] infoOctopus, string[,] infoBancoNoEnOctopus)
        {
            int i = 0;
            int t = 0;



            while (!string.IsNullOrEmpty(infoBanco[i, 0]))
            {
                
                bool montoEncontrado = false;

                int j = 0;



                while (!string.IsNullOrEmpty(infoOctopus[j, 0]))
                {
                    if (!infoBanco[i, 1].Contains("259") && !infoBanco[i, 1].Contains("260") && !infoBanco[i, 1].Contains("212")
                      && !infoBanco[i, 1].Contains("207") && !infoBanco[i, 1].Contains("206") && !infoBanco[i, 1].Contains("515") )
                    {
                        if (infoBanco[i, 3] == infoOctopus[j, 4])
                        {
                            montoEncontrado = true;
                            break;
                        }
                    }
                    j++;

                }
                if (montoEncontrado == false
                       && !infoBanco[i, 1].Contains("259") && !infoBanco[i, 1].Contains("260") && !infoBanco[i, 1].Contains("212")
                      && !infoBanco[i, 1].Contains("207") && !infoBanco[i, 1].Contains("206") && !infoBanco[i, 1].Contains("515") )
                {
                    infoBancoNoEnOctopus[t, 0] = infoBanco[i, 0];
                    infoBancoNoEnOctopus[t, 1] = infoBanco[i, 1];
                    infoBancoNoEnOctopus[t, 2] = infoBanco[i, 2];
                    infoBancoNoEnOctopus[t, 3] = infoBanco[i, 3];
                    t++;
                }
                i++;
            }
            return infoBancoNoEnOctopus;
        }


        public string[,] ProcesarDiferenciasBancoGaliciaNoEnOctopus(int banco, string[,] infoBanco, string[,] infoOctopus, string[,] infoBancoNoEnOctopus)
        {
            int i = 0;
            int t = 0;



            while (!string.IsNullOrEmpty(infoBanco[i, 0]))
            {

                bool montoEncontrado = false;

                int j = 0;



                while (!string.IsNullOrEmpty(infoOctopus[j, 0]))
                {
                    if (!infoBanco[i, 2].Contains("IMP. DEB. LEY 25413 GRAL.")
                        && !infoBanco[i, 2].Contains("IMP. CRE. LEY 25413 GRAL.")
                        && !infoBanco[i, 2].Contains("COMISION MANTENIMIENTO CTA. CTE/CCE")
                      && !infoBanco[i, 2].Contains("IVA")                    )
                    {
                        if (infoBanco[i, 3] == infoOctopus[j, 4])
                        {
                            montoEncontrado = true;
                            break;
                        }
                    }
                    j++;

                }
                if (montoEncontrado == false
                       && !infoBanco[i, 2].Contains("IMP. DEB. LEY 25413 GRAL.")
                        && !infoBanco[i, 2].Contains("IMP. CRE. LEY 25413 GRAL.")
                        && !infoBanco[i, 2].Contains("COMISION MANTENIMIENTO CTA. CTE/CCE")
                      && !infoBanco[i, 2].Contains("IVA"))
                {
                    infoBancoNoEnOctopus[t, 0] = infoBanco[i, 0];
                    infoBancoNoEnOctopus[t, 1] = infoBanco[i, 1];
                    infoBancoNoEnOctopus[t, 2] = infoBanco[i, 2];
                    infoBancoNoEnOctopus[t, 3] = infoBanco[i, 3];
                    t++;
                }
                i++;
            }
            return infoBancoNoEnOctopus;
        }

        public string[,] ProcesarDiferenciasOctopusNoEnBanco(string[,] infoBanco, string[,] infoOctopus, string[,] infoOctopusNoEnBanco)
        {
            int i = 0;
            int t = 0;



            while (!string.IsNullOrEmpty(infoOctopus[i, 0]))
            {
                //  string montoActual = infoOctopus[i, 3];
                bool montoEncontrado = false;

                int j = 0;
                if (i > 43)
                {
                    int a;
                    a = 1;

                }

                while (!string.IsNullOrEmpty(infoBanco[j, 0]))
                {
                    if (j > 44)
                    {
                        int a;
                        a = 1;

                    }

                    if (infoOctopus[i, 4] == infoBanco[j, 3])
                    {
                        montoEncontrado = true;
                        break;
                    }
                    j++;
                }



                if (montoEncontrado == false)
                {
                    infoOctopusNoEnBanco[t, 0] = infoOctopus[i, 0];
                    infoOctopusNoEnBanco[t, 1] = infoOctopus[i, 1];
                    infoOctopusNoEnBanco[t, 2] = infoOctopus[i, 2];
                    infoOctopusNoEnBanco[t, 3] = infoOctopus[i, 3];
                    t++;
                }

                i++;

            }

            return infoOctopusNoEnBanco;
        }



        public decimal[] CalcularSumatoriaCodigosExcluidosBancos(int banco, string[,] infoBanco)
        {
            decimal[] sumaXCodigo = new decimal[10];

            switch (banco)
            {
                case Bancos.Santander:
                    sumaXCodigo = CalcularSumatoriaCodigosExcluidosBancoSantander(infoBanco);
                    break;
                case Bancos.ICBC:
                    sumaXCodigo = CalcularSumatoriaCodigosExcluidosBancoICBC(infoBanco);                   
                    break;

                case Bancos.Galicia:
                    sumaXCodigo = CalcularSumatoriaCodigosExcluidosBancoGalicia(infoBanco);
                    break;


                default:
                    break;
            }

            return sumaXCodigo;
        }

        public decimal[] CalcularSumatoriaCodigosExcluidosBancoSantander(string[,] infoBanco)
        {
            int i = 0;
            decimal importe;
            decimal sumatoria = 0;
            decimal suma4633 = 0;
            decimal suma4637 = 0;
            decimal suma3254 = 0;
            decimal suma1924 = 0;
            decimal suma2960 = 0;
            decimal[] sumaXCodigo = new decimal[6];

            while (!string.IsNullOrEmpty(infoBanco[i, 0]))
            {

                if (infoBanco[i, 1].Contains("4633") || infoBanco[i, 1].Contains("4637") || infoBanco[i, 1].Contains("3254")
                      || infoBanco[i, 1].Contains("1924") || infoBanco[i, 1].Contains("2960"))
                {

                    importe = Convert.ToDecimal(infoBanco[i, 3]);
                    sumatoria = sumatoria + importe;
                }

                switch (infoBanco[i, 1])
                {
                    case "4633":
                        suma4633 = suma4633 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "4637":
                        suma4637 = suma4637 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "3254":
                        suma3254 = suma3254 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "1924":
                        suma1924 = suma1924 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "2960":
                        suma2960 = suma2960 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                }



                i++;
            }


            sumaXCodigo[0] = suma4633;
            sumaXCodigo[1] = suma4637;
            sumaXCodigo[2] = suma3254;
            sumaXCodigo[3] = suma1924;
            sumaXCodigo[4] = suma2960;
            sumaXCodigo[5] = sumatoria;


            return sumaXCodigo;
        }


        public decimal[] CalcularSumatoriaCodigosExcluidosBancoICBC(string[,] infoBanco)
        {
            int i = 0;
            decimal importe;
            decimal sumatoria = 0;
            decimal suma259 = 0;
            decimal suma260 = 0;
            decimal suma212 = 0;
            decimal suma207 = 0;
            decimal suma206 = 0;
            decimal suma515 = 0;






            decimal[] sumaXCodigo = new decimal[10];

            while (!string.IsNullOrEmpty(infoBanco[i, 0]))
            {

                if (infoBanco[i, 1].Contains("259") || infoBanco[i, 1].Contains("260") || infoBanco[i, 1].Contains("212")
                      || infoBanco[i, 1].Contains("207") || infoBanco[i, 1].Contains("206") || infoBanco[i, 1].Contains("515") )
                {

                    importe = Convert.ToDecimal(infoBanco[i, 3]);
                    sumatoria = sumatoria + importe;
                }

                switch (infoBanco[i, 1])
                {
                    case "259":
                        suma259 = suma259 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "260":
                        suma260 = suma260 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "212":
                        suma212 = suma212 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "207":
                        suma207 = suma207 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "206":
                        suma206 = suma206 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "515":
                        suma515 = suma515 + Convert.ToDecimal(infoBanco[i, 3]);
                        break;

                }



                i++;
            }


            sumaXCodigo[0] = suma259;
            sumaXCodigo[1] = suma260;
            sumaXCodigo[2] = suma212;
            sumaXCodigo[3] = suma207;
            sumaXCodigo[4] = suma206;
            sumaXCodigo[5] = suma515;
            sumaXCodigo[6] = sumatoria;


            return sumaXCodigo;
        }



        public decimal[] CalcularSumatoriaCodigosExcluidosBancoGalicia(string[,] infoBanco)
        {
            int i = 0;
            decimal importe;
            decimal sumatoria = 0;
            decimal deb = 0;
            decimal cre = 0;
            decimal mantenimiento = 0;
            decimal iva = 0;







            decimal[] sumaXCodigo = new decimal[10];

            while (!string.IsNullOrEmpty(infoBanco[i, 0]))
            {

                if (infoBanco[i, 2].Contains("IMP. DEB. LEY 25413 GRAL.")
                    || infoBanco[i, 2].Contains("IMP. CRE. LEY 25413 GRAL.") 
                    || infoBanco[i, 2].Contains("COMISION MANTENIMIENTO CTA. CTE/CCE")
                      || infoBanco[i, 2].Contains("IVA") )
                {

                    importe = Convert.ToDecimal(infoBanco[i, 3]);
                    sumatoria = sumatoria + importe;
                }

                switch (infoBanco[i, 2])
                {
                    case "IMP. DEB. LEY 25413 GRAL.":
                        deb = deb + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "IMP. CRE. LEY 25413 GRAL.":
                        cre = cre + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "COMISION MANTENIMIENTO CTA. CTE/CCE":
                        mantenimiento  = mantenimiento + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                    case "IVA":
                        iva = iva + Convert.ToDecimal(infoBanco[i, 3]);
                        break;
                   

                }



                i++;
            }


            sumaXCodigo[0] = deb;
            sumaXCodigo[1] = cre;
            sumaXCodigo[2] = mantenimiento;
            sumaXCodigo[3] = iva;      
            sumaXCodigo[4] = sumatoria;


            return sumaXCodigo;
        }

        public bool GenerarReporteTXT(string[,] infoBancoNoEnOctopus, string[,] infoOctopusNoEnBanco, decimal[] sumatoriaCodigosBancoExcluidos)
        {
            string fileName = @"C:\Desarrollos\ComparadorArchivos\ComparadorArchivos\Output\ReporteComparacion.txt";

            try
            {
                // Check if file already exists. If yes, delete it.     
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                // Create a new file     
                using (StreamWriter sw = File.CreateText(fileName))
                {
                    sw.WriteLine("Conceptos presentes en Banco y no en Octopus");
                    //for (int i = 0; i < 50; i++)
                    int i = 0;
                    while (!string.IsNullOrEmpty(infoBancoNoEnOctopus[i, 0]))
                    {
                        sw.WriteLine(infoBancoNoEnOctopus[i, 0] + ";" + infoBancoNoEnOctopus[i, 1] + ";" + infoBancoNoEnOctopus[i, 2]
                            + ";" + infoBancoNoEnOctopus[i, 3]);
                        i++;
                    }

                    sw.WriteLine("");
                    sw.WriteLine("");

                    i = 0;
                    sw.WriteLine("Conceptos presentes en Octoupus y no en Banco");
                    //for (int i = 0; i < 50; i++)
                    while (!string.IsNullOrEmpty(infoOctopusNoEnBanco[i, 0]))
                    {
                        sw.WriteLine(infoOctopusNoEnBanco[i, 0] + ";" + infoOctopusNoEnBanco[i, 1] + ";" + infoOctopusNoEnBanco[i, 2]
                            + ";" + infoOctopusNoEnBanco[i, 3]);

                        i++;
                    }

                    sw.WriteLine("Sumatoria Còdigos Excluidos Banco = " + sumatoriaCodigosBancoExcluidos);



                }

                // Write file contents on console.     
                using (StreamReader sr = File.OpenText(fileName))
                {
                    string s = "";
                    while ((s = sr.ReadLine()) != null)
                    {
                        Console.WriteLine(s);
                    }
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
            }


            return true;
        }

        public XLWorkbook GenerarReporteExcel(int banco,string[,] infoBancoNoEnOctopus, string[,] infoOctopusNoEnBanco, decimal[] sumatoriaCodigosBancoExcluidos)
        {
            //   Application excel;
            XLWorkbook oWorkbook = new XLWorkbook();// = new Workbook ();
            var oWorksheet = oWorkbook.AddWorksheet("Comparativa");
            int i = 0;


           // string fileName = @"C:\Desarrollos\ComparadorArchivos\ComparadorWebform\ComparadorWebform\Archivos\Output\ArchivoComparacionExcel.xlsx";



            try
            {
                //// Check if file already exists. If yes, delete it.     
                //if (File.Exists(fileName))
                //{
                //    File.Delete(fileName);
                //}

                //excel = new Application();
                //excel.Visible = false;
                //excel.DisplayAlerts = false;
                //worKbooK = excel.Workbooks.Add(Type.Missing);




                oWorksheet.Cell("A1").Value = "COMPARACIÓN DE ARCHIVOS";

                oWorksheet.Cell("A3").Value = "Conceptos presentes en Banco y no en Octopus";





                oWorksheet.Cell("A4").Value = "FECHA";
                oWorksheet.Cell("B4").Value = "CÓDIGO";
                oWorksheet.Cell("C4").Value = "CONCEPTO";
                oWorksheet.Cell("D4").Value = "IMPORTE";

                i = 5;
                int j = 0;
                while (!string.IsNullOrEmpty(infoBancoNoEnOctopus[j, 0]))
                {

                    //  oWorksheet.Cells(i, 1) = infoBancoNoEnOctopus[j, 0].ToString();
                    oWorksheet.Cells("A" + i).Value = infoBancoNoEnOctopus[j, 0].ToString();

                    //Para banco galicia no viene código
                    if (!string.IsNullOrEmpty(infoBancoNoEnOctopus[j, 1]))
                        oWorksheet.Cells("B" + i).Value = infoBancoNoEnOctopus[j, 1].ToString();


                    oWorksheet.Cells("C" + i).Value = infoBancoNoEnOctopus[j, 2].ToString();
                    oWorksheet.Cells("D" + i).Value = Convert.ToDecimal ( infoBancoNoEnOctopus[j, 3].ToString());

                    i++;
                    j++;
                }





                i = i + 1;

                oWorksheet.Cell("A" + i).Value = "Conceptos presentes en Octoupus y no en Banco";


                i = i + 1;
                oWorksheet.Cells("A" + i).Value = "FECHA";
                oWorksheet.Cells("C" + i).Value = "CONCEPTO";
                oWorksheet.Cells("D" + i).Value = "DEBE";
                oWorksheet.Cells("E" + i).Value = "HABER";


                //for (int i = 0; i < 50; i++)
                j = 0;
                i++;
                while (!string.IsNullOrEmpty(infoOctopusNoEnBanco[j, 0]))
                {
                    oWorksheet.Cells("A" + i).Value = infoOctopusNoEnBanco[j, 0];
                    oWorksheet.Cells("C" + i).Value = infoOctopusNoEnBanco[j, 1];
                    oWorksheet.Cells("D" + i).Value =  Convert.ToDecimal (infoOctopusNoEnBanco[j, 2].ToString());
                    oWorksheet.Cells("E" + i).Value = Convert.ToDecimal(infoOctopusNoEnBanco[j, 3].ToString());

                    j++;
                    i++;
                }

                i++;

                oWorksheet.Cell("A" + i).Value = "Sumatoria Códigos Excluidos Banco:";

                i++;


                switch (banco)
                {
                    case Bancos.Santander:
                        oWorksheet.Cell("A" + i).Value = "Sumatoria código 4633 ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[0];
                        i++;
                        oWorksheet.Cell("A" + i).Value = "Sumatoria código 4637 ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[1];
                        i++;
                        oWorksheet.Cell("A" + i).Value = "Sumatoria código 3254 ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[2];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "Sumatoria código 1924 ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[3];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "Sumatoria código 2960 ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[4];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "TOTAL CODIGOS EXCLUIDOS ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[5];
                        break;
                    case Bancos.ICBC:
                        oWorksheet.Cell("A" + i).Value = "Sumatoria código 259 ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[0];
                        i++;
                        oWorksheet.Cell("A" + i).Value = "Sumatoria código 260 ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[1];
                        i++;
                        oWorksheet.Cell("A" + i).Value = "Sumatoria código 212 ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[2];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "Sumatoria código 207 ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[3];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "Sumatoria código 206 ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[4];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "Sumatoria código 515 ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[5];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "TOTAL CODIGOS EXCLUIDOS ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[6];
                        break;

                    case Bancos.Galicia:
                        oWorksheet.Cell("A" + i).Value = "Sumatoria IMP. DEB. LEY 25413 GRAL. ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[0];
                        i++;
                        oWorksheet.Cell("A" + i).Value = "Sumatoria IMP. CRE. LEY 25413 GRAL. ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[1];
                        i++;
                        oWorksheet.Cell("A" + i).Value = "Sumatoria COMISION MANTENIMIENTO CTA. CTE/CCE ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[2];    
                        i++;
                        oWorksheet.Cell("A" + i).Value = "Sumatoria IVA ";
                        oWorksheet.Cell("B" + i).Value = sumatoriaCodigosBancoExcluidos[3];
                        i++;
                        oWorksheet.Cells("A" + i).Value = "TOTAL CODIGOS EXCLUIDOS ";
                        oWorksheet.Cells("B" + i).Value = sumatoriaCodigosBancoExcluidos[4];
                        break;

                }




                return oWorkbook;




                //worKbooK.SaveAs(@"C:\Desarrollos\ComparadorArchivos\ComparadorWebform\ComparadorWebform\Archivos\Output\ArchivoComparacionExcel.xlsx");
                //worKbooK.Close();
                //excel.Quit();


            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);

            }
            finally
            {

                // worKsheeT = null;
                //celLrangE = null;
                //worKbooK = null;
            }

            return null;
        }

        public static string ConvertXLS_XLSX(FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            // var xlsFile = file.FullName;
            var xlsFile = @"C:\Desarrollos\ComparadorArchivos\ComparadorArchivos\Input\ImputBanco2.xls";
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return xlsxFile;
        }
    }
}