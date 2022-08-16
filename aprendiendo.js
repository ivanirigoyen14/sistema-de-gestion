function Recorrer() {
  var servicio = SpreadsheetApp;
  var archivo = servicio.getActiveSpreadsheet();
  var hoja = archivo.getActiveSheet();
  
}

  function Facturar2(){
    var app = SpreadsheetApp.getActive();
    
    var hoja_producto = app.getSheets()[0];
    var hoja_clientes = app.getSheets()[1];
    var hoja_hacerfactura = app.getSheets()[2];
    var hoja_historialfacturas = app.getSheets()[3];
    

    Logger.log(hoja_producto.getName());
    Logger.log(hoja_clientes.getName());
    Logger.log(hoja_hacerfactura.getName());
    Logger.log(hoja_historialfacturas.getName());
    

    /*var hoja_producto = app.openById("1794044215");
    var hoja_clientes = app.openById("1467283655");
    var hoja_hacerfactura = app.openById("883688930");
    var hoja_historialfacturas = app.openById("0");
    var hoja_factura = app.openById("942379042");*/

    app.setActiveSheet(hoja_historialfacturas, true);
    //app.setActiveSheet(app.getSheetByName('Historial Facturas'), true);//ir a la hoja....
    if(app.getActiveRange().getRow()== 1 || app.getActiveCell().isBlank())
    {
      
      Browser.msgBox('Debes seleccionar una Factura valida');
      return;

    }
    Logger.log('la fila es %s',app.getActiveRange().getRow());
    //guardar codigo factura, fecha y cod cliente
    var cod_fact = app.getRange('A'+ app.getActiveRange().getRow()).getValue();
    var fecha = app.getRange('B'+ app.getActiveRange().getRow()).getDisplayValue();
    var cod_cli = app.getRange('C'+ app.getActiveRange().getRow()).getValue();
    Logger.log(cod_fact);
    Logger.log(fecha);
    Logger.log(cod_cli);

   //Obtenemos los datos del cliente a partir del "cod_cli"

    //Vamos a la hoja clientes
    app.setActiveSheet(hoja_clientes, true);
    //buscamos codigos de cliente a partir de la fila 14

    
    var fila_cliente = 14;
    //var idcliente = app.getRange('A'+ fila_cliente).getValue();
    //Logger.log(idcliente);
    Logger.log(app.getRange('A'+ fila_cliente).getDisplayValue());
    Logger.log(app.getRange('A'+ fila_cliente).isBlank() != true);
    //recorremos las filas de la hoja "clientes"
    //var valor = app.getRange('A14').activate().getValue();
    for(var fila_cliente = 13 ;app.getRange('A'+ fila_cliente).getDisplayValue()!= ''; fila_cliente++)
    {
      //SI EL CODIGO DE CLIENTE COINCIDE OBTENEMOS LOS DATOS

      if(app.getRange('A'+ fila_cliente).getValue() == cod_cli)
      {
        var nombre = app.getRange('B'+ fila_cliente).getValue()+' '+app.getRange('C'+fila_cliente).getValue();
        var direccion = app.getRange('D'+ fila_cliente).getValue();
        var dni = app.getRange('F'+ fila_cliente).getValue();
        break;
      }
      fila_cliente++;
      Logger.log(fila_cliente);
      Logger.log('A'+ fila_cliente);
      Logger.log(app.getRange('A'+ fila_cliente).getDisplayValue()!= '');
    }
    Logger.log(nombre);
    Logger.log(direccion);
    Logger.log(dni);
    
    //Vamos a la hoja factura
    var hoja_factura = app.getSheets()[4];
    Logger.log(hoja_factura.getName());
    app.setActiveSheet(hoja_factura, true);
    //app.setActive(app.getSheetByName('Factura'), false);
    //vaciar datos de la factura
    Logger.log('ESTA EN BLANCO LA CELDA? : ', !app.getRange('B17').isBlank());
    while(!app.getRange('B17').isBlank)
    {
      app.deleteRow(17);
    }

    //Insertamos los datos: cod_fact, fecha, nombre, direccion y dni

    app.getRange('F10').setValue(cod_fact);
    app.getRange('C10').setValue(nombre);
    app.getRange('C12').setValue(direccion);
    app.getRange('C11').setValue(dni);
    app.getRange('B8').setValue(fecha);
    
    //Buscamos productos en la hoja de historial factura y las introducimos en la factura
    //Vamos a la hoja Historial Facturas
    app.setActiveSheet(hoja_historialfacturas, true);
    //buscamos a partir de la fila 2

    var fila_factura = 2;
    var celda = 'A'+fila_factura;
    Logger.log('La fila es> %s',fila_factura);

    //Recorremos todas las filas de la hoja
    Logger.log('Codigo de factura %s', app.getRange(celda).getValue())
    while(!app.getRange(celda).isBlank())
    {
      //si el codigo de la factura coincide copiamos los datos
      if(app.getRange(celda).getValue()==cod_fact)
      {
        //Obtener datos
        var nombre_producto = app.getRange('F'+fila_factura).getValue();
        var cantidad = app.getRange('G'+fila_factura).getValue();
        var total_producto = app.getRange('H'+fila_factura).getValue();
        Logger.log('El nombre de producto es> %s', nombre_producto);
        Logger.log('La cantidad es> %s', cantidad);
        Logger.log('El total de producto es> %s', total_producto);
        //vamos a la hoja Factura

        app.setActiveSheet(app.getSheetByName('Factura'), true);

        //insertar una fila en blanco en la fila 17

        app.getActiveSheet().insertRowBefore(17);

        //copiamos los datos

        app.getRange('B17').setValue(nombre_producto);
        app.getRange('E17').setValue(cantidad);
        app.getRange('G17').setValue(total_producto);
        app.getRange('F17').setValue(total_producto/cantidad);

        //Vamos a la hoja historial facturas
        app.setActiveSheet(hoja_historialfacturas, true);
      }
      fila_factura++;
    }
  }
/*var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E14').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Factura '), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Historial Facturas'), true);
  spreadsheet.getRange('A3:H3').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Factura '), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Historial Facturas'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Factura '), true);
  spreadsheet.getRange('B17:D17').activate();*/

 // spreadsheet.getRange('\'Historial Facturas\'!A4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);