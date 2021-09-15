
# SAP
  
With this module you can automate SAP R3 and Hana. Perform all the actions that you can record with the tracker shared 
by Stefan Schnell in the official SAP forum  
  
![banner](https://raw.githubusercontent.com/rocketbot-cl/SAPObjetos/master/img/Banner_SAPObjetos.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.
## Como usar este module
  
Eiusmod veniam ut nisi minim in. Do et deserunt eiusmod veniam sint aliqua nulla adipisicing laboris voluptate fugiat 
ullamco elit do. Sint amet cillum fugiat excepteur mollit voluptate reprehenderit nisi commodo sint minim.
## Descripción de los comandos

### LoginSap
  
Abre la aplicación de sap, establece la conexión y realiza el login
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta SAP|Ruta de la aplicación de SAP|C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe|
|Nombre Conexión|Nombre de la conexión donde quieres logearte|Desarrollo (Directo)|
|ID Usuario|Identificador del campo de usuario de SAP|wnd[0]/usr/txtRSYST-BNAME|
|Usuario|Nombre de usuario que ingresas en el campo usuario al hacer login|Usuario1|
|ID Password|Identificador del campo de contraseña de SAP|wnd[0]/usr/pwdRSYST-BCODE|
|Contraseña|Contraseña que ingresas en el campo contraseña al hacer login|Contraseña|

### Click en Objeto
  
Click en objeto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Valor|Valor mostrado en el tracker después de un igual (=) (Ej. id.atributo = valor), o como entre parentesis (Ex. id.atributo(valor)|QMNUM|
|Opción|Propiedad indicada en el tracker después de findById('algun id'). Ej session.findById('wnd[0]/tbar[0]/okcd').propiedad|Opción|

### Extraer Texto
  
Extraer texto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Marca/Desmarca Check
  
Marca/Desmarca Check
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Opción|Seleccionar si quieres marcar o desmarcar el checkbox|Check|
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|

### Extraer Celda
  
Extraer texto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Número de fila|Número de la fila donde se quiere extraer el texto. En el tracker aparecerá como selectedRows|0|
|Nombre de la columna|Nombre de la columna donde se quiere extraer el texto. En el tracker aparecerá como currentCellColumn|TYPE_DOC|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Click en Celda
  
Click en una celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Número de fila|Número de la fila donde se quiere extraer el texto. En el tracker aparecerá como selectedRows|0|
|Nombre de la columna|Nombre de la columna donde se quiere extraer el texto. En el tracker aparecerá como currentCellColumn|TYPE_DOC|
|Tipo de click|Propiedad indicada en el tracker después de findById('algun id'). Ej session.findById('wnd[0]/tbar[0]/okcd').propiedad|Opción|

### Ejecutar Script
  
Ejecuta un script VBS grabador con SAP
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta VBA Script|Ruta del archivo VBs|C:/ruta/al/script.vbs|
|Script predefinido||Opción|

### Obtener estado de checkbox
  
Devuelve True si el checkbox está seleccionado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Enviar tecla
  
Replica el evento enviar tecla
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Tecla|Tecla o combinación que se desea enviar|{'en': 'F1'}|

### Obtener propiedad
  
Obtiene una propiedad del objeto SAP especificado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Propiedad|Propiedad que se quiere obtener. Puedes ver todas las propiedades acá https//help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.03/en-US/a2e9357389334dc89eecc1fb13999ee3.html|Opción|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|
