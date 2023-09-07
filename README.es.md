



# SAP
  
Con este módulo puedes automatizar SAP R3 y Hana. Realiza todas las acciones que puedas grabar con el rastreador creado por Stefan Schnell.  

*Read this in other languages: [English](README.md), [Português](README.pr.md), [Español](README.es.md)*

## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Overview


1. LoginSap  
Abre la aplicación de sap, establece la conexión y realiza el login

2. Conectar  
Conéctese a una aplicación SAP abierta

3. Ejecutar una acción  
Ejecutar una acción en SAP como seleccionar, hacer foco o modificar el texto de un elemento mediante diferentes propiedades (setFocus, text, etc.)

4. Extraer Texto  
Extrae el texto de un objeto en sap, mediante la propiedad text

5. Marca/Desmarca Check  
Marca o desmarca un objeto en SAP

6. Extraer Celda  
Extrae el texto en una celda, puede ser con GetItemText o GetCellValue

7. Click en Celda  
Realiza un click en una celda. Puede ser con clickCurrentCell, setCurrentCell o doubleClickCurrentCell

8. Ejecutar Script  
Ejecuta un script VBS grabador con SAP

9. Obtener estado de checkbox  
Devuelve True si el checkbox está seleccionado

10. Enviar tecla  
Replica el evento enviar tecla

11. Obtener propiedad  
Obtiene una propiedad del objeto SAP especificado

12. Esperar objeto  
Espera a que un objeto sea visible  




----
### OS

- windows
- docker

### Dependencies

### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)