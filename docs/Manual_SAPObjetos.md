



# SAP
  
Con este módulo podras automatizar SAP R3 y Hana. Realizar acciones que puede grabar con el tracker compartido por Stefan Schnell en el foro oficial de SAP.
  
![banner](imgs/Banner_SAPObjetos.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.  




## Cómo usar este módulo
Para usar este módulo, tienes que conectarte a tu cuenta de SAP.

Primeramente hay que activar los scripts en SAP GUI.

En SAP GUI, tienes que ir a la transacción RZ11; en el nombre del parámetro, escribe "sapgui/user_scripting" y haz click en Mostrar (display).

![sap1](imgs/sap1.png)
![sap2](imgs/sap2.png)

Haz clicn en "Cambiar valor" y en "Nuevo valor" selecciona "TRUE". Guárdalo.

![sap3](imgs/sap3.png)
![sap4](imgs/sap4.png)

La siguiente imagen indica que GUI Scripting esta habilitado en RZ11.

![sap5](imgs/sap5.png)

En "Accesibilidad y Scripting", ve a "Scripting" y activa la opción "Enable scripting". Guárdalo.

![sap6](imgs/sap6.png)

Como trabajar correctamente

Abre Tracker.exe

![sap7](imgs/sap7.png)

Abrirá una ventana como estas

![sap8](imgs/sap8.png)

Abre SAP y logueate. Una vez loqueado, dale click al ícono y sincronizará la sesión activa de SAP.

![sap9](imgs/sap9.png)

Ve a "Recorder" y dale click en el ícono de Python.

![sap10](imgs/sap10.png)


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
  
Hace un click en SAP. Puede ser mediante diferentes propiedades, como setFocus, text, etc. También puedes modificar el 
texto de un elemento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Valor|Valor mostrado en el tracker después de un igual (=) (Ej. id.atributo = valor), o como entre parentesis (Ex. id.atributo(valor)|QMNUM|
|Opción|Propiedad indicada en el tracker después de findById('algun id'). Ej session.findById('wnd[0]/tbar[0]/okcd').propiedad|Opción|

### Extraer Texto
  
Extrae el texto de un objeto en sap, mediante la propiedad text
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Marca/Desmarca Check
  
Marca o desmarca un objeto en SAP
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Opción|Seleccionar si quieres marcar o desmarcar el checkbox|Check|
|Marcar individual con absoluteRow|Opción de marcarlos con columna absoluta|Checkbox|
|Marcar todos|Opción de marcarlos a todos|Checkbox|
|Columna absoluta|Columna con la que se desea interactuar|3|
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|

### Extraer Celda
  
Extrae el texto en una celda, puede ser con GetItemText o GetCellValue
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Número de fila|Número de la fila donde se quiere extraer el texto. En el tracker aparecerá como selectedRows|0|
|Nombre de la columna|Nombre de la columna donde se quiere extraer el texto. En el tracker aparecerá como currentCellColumn|TYPE_DOC|
|Opción||Opción|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Click en Celda
  
Realiza un click en una celda. Puede ser con clickCurrentCell, setCurrentCell o doubleClickCurrentCell
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
|Tecla|Tecla o combinación que se desea enviar|F1|

### Obtener propiedad
  
Obtiene una propiedad del objeto SAP especificado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Propiedad|Propiedad que se quiere obtener. Puedes ver todas las propiedades acá https//help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.03/en-US/a2e9357389334dc89eecc1fb13999ee3.html|Opción|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|
