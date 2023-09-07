



# SAP
  
Con este módulo puedes automatizar SAP R3 y Hana. Realiza todas las acciones que puedas grabar con el rastreador creado por Stefan Schnell.  

*Read this in other languages: [English](Manual_SAPObjetos.md), [Português](Manual_SAPObjetos.pr.md), [Español](Manual_SAPObjetos.es.md)*
  
![banner](imgs/Banner_SAPObjetos.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  



## Cómo usar este módulo
Para usar este módulo, tienes que conectarte a tu cuenta de SAP.

Primeramente hay que activar los scripts en SAP GUI.

En SAP GUI, tiene sque ir a la transacción RZ11; en el nombre del parámetro, escribe "sapgui/user_scripting" y haz click en Mostrar (display).

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
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|
|Guardar resultado de conexión|Variable donde se guardará el resultado de la conexión|variable|

### Conectar
  
Conéctese a una aplicación SAP abierta
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre Conexión|Nombre de la conexión donde quieres logearte|Desarrollo (Directo)|
|Guardar resultado de conexión|Variable donde se guardará el resultado de la conexión|variable|

### Ejecutar una acción
  
Ejecutar una acción en SAP como seleccionar, hacer foco o modificar el texto de un elemento mediante diferentes propiedades (setFocus, text, etc.)
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Número de fila (Opcional)|Número de la fila donde se quiere ejecutar la acción. En el tracker aparecerá como selectedRows|0|
|Nombre de la columna (Opcional)|Nombre de la columna donde se quiere ejecutar la acción. En el tracker aparecerá como currentCellColumn|TYPE_DOC|
|Valor|Valor mostrado en el tracker después de un igual (=) (Ej. id.atributo = valor), o como entre parentesis (Ex. id.atributo(valor)|QMNUM|
|Opción|Propiedad indicada en el tracker después de findById('algun id'). Ej session.findById('wnd[0]/tbar[0]/okcd').propiedad|Opción|
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|

### Extraer Texto
  
Extrae el texto de un objeto en sap, mediante la propiedad text
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|
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
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|

### Extraer Celda
  
Extrae el texto en una celda, puede ser con GetItemText o GetCellValue
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Número de fila|Número de la fila donde se quiere extraer el texto. En el tracker aparecerá como selectedRows|0|
|Nombre de la columna|Nombre de la columna donde se quiere extraer el texto. En el tracker aparecerá como currentCellColumn|TYPE_DOC|
|Opción||Opción|
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Click en Celda
  
Realiza un click en una celda. Puede ser con clickCurrentCell, setCurrentCell o doubleClickCurrentCell
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Número de fila|Número de la fila donde se quiere extraer el texto. En el tracker aparecerá como selectedRows|0|
|Nombre de la columna|Nombre de la columna donde se quiere extraer el texto. En el tracker aparecerá como currentCellColumn|TYPE_DOC|
|Tipo de click|Propiedad indicada en el tracker después de findById('algun id'). Ej session.findById('wnd[0]/tbar[0]/okcd').propiedad|Opción|
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|

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
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Enviar tecla
  
Replica el evento enviar tecla
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Número de fila (Opcional)|Número de la fila donde se quiere enviar tecla. En el tracker aparecerá como selectedRows|0|
|Nombre de la columna (Opcional)|Nombre de la columna donde se quiere enviar tecla. En el tracker aparecerá como currentCellColumn|TYPE_DOC|
|Tecla|Tecla o combinación que se desea enviar|F1|
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|

### Obtener propiedad
  
Obtiene una propiedad del objeto SAP especificado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Propiedad|Propiedad que se quiere obtener. Puedes ver todas las propiedades acá https//help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.03/en-US/a2e9357389334dc89eecc1fb13999ee3.html|Opción|
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|

### Esperar objeto
  
Espera a que un objeto sea visible
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Objeto|Identificador obtenido en el tracker|wnd[0]/tbar[0]/okcd|
|Timeout|Lapso de tiempo (en segundos) para buscar el elemento|10|
|Asignar a Variable|Nombre de variable donde guardar el resultado|variable|
