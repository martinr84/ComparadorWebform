function validarIngresos() {


    //Validamos que hayan seleccionado tipo de decimales
    var rbPuntoDecimal =  document.getElementById("MainContent_rbdDecimales_0");
    var rbComaDecimal = document.getElementById("MainContent_rbdDecimales_1");
    var ddlBancos = document.getElementById("MainContent_ddlBancos");


    if (ddlBancos.value  == 0) {
        alert('No se seleccionó cuál es el banco');
        return false;
    }



    if (rbPuntoDecimal.checked == false && rbComaDecimal.checked == false)
    {
        alert('Debe seleccionar tipo de decimales');
        return false;
    }

   


    //Validamos archivos subidos de banco y octopus
    var fileInputBanco = document.getElementById('MainContent_fuBanco');
    var fileInputOctopus = document.getElementById('MainContent_fuOctopus');

    var filePathBanco = fileInputBanco.value;
    var filePathOctopus = fileInputOctopus.value;

    // Allowing file type 
    var allowedExtensions =
        /(\.xlsx)$/i;

    if (filePathBanco == "") {
        alert('No se seleccionó archivo de banco a comparar');
        return false;
    }
    else {
        if (!allowedExtensions.exec(filePathBanco)) {
            alert('Archivo de banco con extension Invalida');
            fileInputBanco.value = '';
            return false;
        }
    }



    if (filePathOctopus == "") {
        alert('No se seleccionó archivo de octopus a comparar');
        return false;
    }
    else {
        if (!allowedExtensions.exec(filePathOctopus)) {
            alert('Archivo de octopus con extension Invalida');
            fileInputOctopus.value = '';
            return false;
        }
    }



}


function abrirArchivo() {
        var Excel = new ActiveXObject("Excel.Application");
        Excel.Visible = true; Excel.Workbooks.open("Archivos/Output/ArchivoComparacionExcel.xlsx");
        var excel_sheet = Excel.Worksheets("sheetname_1");
        excel_sheet.activate();
}

window.onload = function () {
    //document.getElementById("MainContent_fuBanco").disabled = true;
    //document.getElementById("MainContent_fuOctopus").disabled = true;
    //document.getElementById("MainContent_btnEjecutarProceso").disabled = true;
    //document.getElementById("MainContent_rbdDecimales_0").disabled = true;
    //document.getElementById("MainContent_rbdDecimales_1").disabled = true;


  
};


function habilitarElementos() {

    //document.getElementById("MainContent_fuBanco").disabled = false;
    //document.getElementById("MainContent_fuOctopus").disabled = false;
    //document.getElementById("MainContent_rbdDecimales").disabled = false;
    //document.getElementById("MainContent_btnEjecutarProceso").disabled = false;
    //document.getElementById("MainContent_rbdDecimales_0").disabled = false;
    //document.getElementById("MainContent_rbdDecimales_1").disabled = false;   
   
 
    
}


 
