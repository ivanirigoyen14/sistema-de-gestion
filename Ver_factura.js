function Facturar() {
  //guardar la hoja de calculo en la variable gestor

  var gestor = SpreadsheetApp.getActive();
  gestor.getActiveSheet();
 // var num_fila = gestor.getActiveRange().getRow()
 Logger.log(gestor.getActiveRange().getRow());
  if(gestor.getActiveRange().getRow()== 1 || gestor.getActiveCell().isBlank())
  {
    
    Browser.msgBox('Debes seleccionar una Factura valida');
    return;

  }
  Logger.log(gestor.getActiveRange().getRow());
  //guardar codigo factura, fecha y cod cliente
  var cod_fact = gestor.getRange('A'+ gestor.getActiveRange().getRow()).getValue();
  var fecha = gestor.getRange('B'+ gestor.getActiveRange().getRow()).getValue();
  var cod_cli = gestor.getRange('C'+ gestor.getActiveRange().getRow()).getValue();

  var nombre= '';
  var direccion= '';
  var dni= '';

  //Obtenemos los datos del cliente a partir del "cod_cli"

  //Vamos a la hoja clientes
  gestor.setActiveSheet(gestor.getSheetByName('Clientes'));

  //buscamos codigos de cliente a partir de la fila 14

  var fila_cliente = 14;
  var col_fila = 'A'+fila_cliente;
  //recorremos las filas de la hoja "clientes"
  var valor = gestor.getRange('A14').activate().getValue();
  /*while(!gestor.getRange('A'+fila_cliente).isBlank())
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
  gestor.setActiveSheet(gestor.getSheetByName('Factura'));

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
  }*/
}
