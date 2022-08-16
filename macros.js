function Ingreso_prod() {
  var spreadsheet = SpreadsheetApp.getActive();

  //Comprobar Producto
  if(spreadsheet.getRange('B8').isBlank() || spreadsheet.getRange('B9').isBlank() )
  {
    Browser.msgBox('ERROR','Debes introducir Producto', Browser.Buttons.OK);
    return;
  }

  spreadsheet.getRange('8:8').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A8').activate();
  spreadsheet.getRange('B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B8').activate();
  spreadsheet.getRange('B3:B4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
  spreadsheet.getRange('B3:B4').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1').activate();
};

function Ingreso_cliente() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('14:14').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A14').activate();
  spreadsheet.getRange('B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B14').activate();
  spreadsheet.getRange('B3:B11').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
  spreadsheet.getRange('B3:B5'||'B7:B11').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1').activate();
  spreadsheet.getRange('B7:B11').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B3:B5').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1').activate();
};

function InsertarProductoFactura() {
  var spreadsheet = SpreadsheetApp.getActive();

    //Comprobaciones

    //Comprobar Fecha
  
    if(spreadsheet.getRange('B4').isBlank() )
   {
    Browser.msgBox('ERROR','Debes introducir Fecha', Browser.Buttons.OK);
    return;
   }
   //Comprobar cliente
  
    if(spreadsheet.getRange('B3').isBlank() || spreadsheet.getRange('B3').getValue()=='NOMBRE Cliente' )
   {
    Browser.msgBox('ERROR','Debes introducir Cliente', Browser.Buttons.OK);
    return;
    }
   //Comprobar Producto
   if(spreadsheet.getRange('B8').isBlank() || spreadsheet.getRange('B9').isBlank() || spreadsheet.getRange('B8').getValue()   =='Producto' )
   {
    Browser.msgBox('ERROR','Debes introducir Producto', Browser.Buttons.OK);
    return;
   }
    spreadsheet.getRange('13:13').activate();
    spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
    spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    spreadsheet.getActiveRangeList().setBackground("#FFFFFF");
    spreadsheet.getRange('A13').activate();
    spreadsheet.getRange('B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('B13').activate();
    spreadsheet.getRange('B4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('D13').activate()
    spreadsheet.getRange('B3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('C13').activate();
    spreadsheet.getRange('F2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('E13').activate();
    spreadsheet.getRange('F8').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('F13').activate();
    spreadsheet.getRange('B8').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('G13').activate();
    spreadsheet.getRange('B9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('H13').activate();
    spreadsheet.getRange('B10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    //Comprobar Stock
    if(spreadsheet.getRange('B9').getValue() > spreadsheet.getRange('F9').getValue())
    {
      spreadsheet.getRange('A13:H13').activate();
      spreadsheet.getActiveRangeList().setBackground('ACCENT3');
      //Browser.msgBox('ERROR','No HAY STOCK SUFICIENTE', Browser.Buttons.OK);
      return;
    }
   
    
  
    
};

function EliminarProductoFactura() {
  var spreadsheet = SpreadsheetApp.getActive();
  if(spreadsheet.getActiveRange().getRow()>12)
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  else
  {
    Browser.msgBox('ERROR','Selecciona celda valida',Browser.Buttons.OK);
  }
};

function Guardarfactura() {
var spreadsheet = SpreadsheetApp.getActive();

while(!spreadsheet.getRange('A13').isBlank())
{
   Agregarproductofactura();
}

}

function Agregarproductofactura() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('13:13').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Historial Facturas'), true);
  spreadsheet.getRange('2:2').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A2').activate();
  spreadsheet.getRange('\'Hacer Factura\'!13:13').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Hacer Factura'), true);
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};

//Toma datos de la hoja Historial factura y los incerta e la factura
function Ver_factura() {

  //guardar la hoja de calculo en la variable gestor

  var gestor = SpreadsheetApp.getActive();
  
  //Vamos a la hoja clientes
  gestor.setActiveSheet(gestor.getSheetByName('Historial Facturas'));
  
  //Vamos a la hoja Historial Facturas
  //gestor.getActiveSpreadsheet().getSheetByName("Historial Facturas");
  
  //Activamos la hoja Historial Facturas
  //gestor.setActiveSheet.getSheetByName('Historial Facturas');
  //comprobar si los datos son validos!
  var num_fila = gestor.getActiveRange().getRow()
  if((gestor.getActiveRange().getRow()== 1) || (gestor.getActiveCell().isBlank()))
  {
    
    Browser.msgBox('Debes seleccionar una Factura valida');
    return;

  }
  
  //guardar codigo factura, fecha y cod cliente
  var cod_fact = gestor.getRange('A'+gestor.getActiveRange().getRow()).getValue();
  var fecha = gestor.getRange('B'+gestor.getActiveRange().getRow()).getValue();
  var cod_cli = gestor.getRange('C'+gestor.getActiveRange().getRow()).getValue();

  var nombre= '';
  var direccion= '';
  var dni= '';

  //Obtenemos los datos del cliente a partir del "cod_cli"

  //Vamos a la hoja clientes
  gestor.setActiveSheet(gestor.getSheetByName('Clientes'));

  //buscamos codigos de cliente a partir de la fila 14

  var fila_cliente = 14

  //recorremos las filas de la hoja "clientes"

  while(!gestor.getRange('A'+fila_cliente).isBlank())
  {
    //SI EL CODIGO DE CLIENTE COINCIDE OBTENEMOS LOS DATOS

    if(gestor.getRange('A'+fila_cliente).getValue()==cod_cli)
    {
      nombre = gestor.getRange('B'+fila_cliente).getValue()+' '+gestor.getRange('C'+fila_cliente).getValue();
      direccion = gestor.getRange('D'+fila_cliente).getValue();
      dni = gestor.getRange('F'+fila_cliente).getValue();
      break;
    }
    fila_cliente++;
  }
  //Vamos a la hoja de ver factura
  gestor.setActiveSheet().getSheetByName('Factura');

  //Insertamos los datos: cod_fact, fecha, nombre, direccion y dni

   gestor.getRange('F10').setValue(cod_fact);
   gestor.getRange('C10').setValue(nombre);
   gestor.getRange('C12').setValue(direccion);
   gestor.getRange('C11').setValue(dni);
   gestor.getRange('B8').setValue('Fecha: '+ fecha);

  //vaciar datos de la factura
  while(!gestor.getRange('B17').isBlank())
  {
    gestor.deleteRow(17);
  }

  //Buscamos productos en la hoja de historial factura y las introducimos en la factura
  //Vamos a la hoja factura
  gestor.setActiveSheet(gestor.getSheetByName('Historial Facturas'));

  //buscamos a partir de la fila 2

  var fila_factura = 2;

  //Recorremos todas las filas de la hoja

  while(!gestor.getRange('A'+fila_factura).isBlank())
  {
    //si el codigo de la factura coincide copiamos los datos
    if(gestor.getRange('A'+fila_factura).getValue==cod_fact)
    {
      //Obtener datos
      var nombre_producto = gestor.getRange('F'+fila_factura).getValue();
      var cantidad = gestor.getRange('G'+fila_factura).getValue();
      var total_producto = gestor.getRange('H'+fila_factura).getValue();

      //vamos a la hoja Factura

      gestor.setActiveSheet(gestor.getSheetByName('Factura'));

      //insertar una fila en blanco en la fila 17

      gestor.getActiveSheet().insertRowBefore(17);

      //copiamos los datos

      gestor.getRange('B17').setValue(nombre_producto);
      gestor.getRange('E17').setValue(cantidad);
      gestor.getRange('G17').setValue(total_producto);

      //Vamos a la hoja hiistorial facturas
       gestor.setActiveSheet(gestor.getSheetByName('Historial Facturas'));
    }
    fila_factura++;
  }
    //mostramos la factura 

    gestor.setActiveSheet(gestor.getSheetByName('Factura'));
}



function Prueba_pestania() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E14').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Factura '), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Historial Facturas'), true);
  spreadsheet.getRange('A3:H3').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Factura '), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Historial Facturas'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Factura '), true);
  spreadsheet.getRange('B17:D17').activate();
  spreadsheet.getRange('\'Historial Facturas\'!A4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};