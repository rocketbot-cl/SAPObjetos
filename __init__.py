# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
from win32com import client
import sys
import subprocess
import time

"""
    Obtengo el modulo que fue invocado
"""
module = GetParams("module")
global SAPObject
global id_object



"""
    Realiza Login
"""
if module == "LoginSap":
    # global conx
    # global id_object

    id_user = GetParams('id_user')
    user = GetParams('user')
    id_pass = GetParams('id_pass')
    password = GetParams('password')
    path = GetParams('path')
    conn = GetParams('conn')
    id_button = GetParams('id_btn')

    """
        validaciones
    """
    if not id_user or not id_pass or not conn or not path:
        raise Exception('No ha ingresado todos los datos')

    if id_user:
        try:
            if not path:
                path = "C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe"
            subprocess.Popen(path)
            time.sleep(10)

            print("Get Object 'SAPGUI'")
            SapGuiAuto = client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            print("Open connection")
            connection = application.OpenConnection(conn, True)
            print("session")
            session = connection.Children(0)

            if user and password:
                session.findById(id_user).SetFocus()
                session.findById(id_user).text = user
                session.findById(id_pass).text = password

            if connection:
                SAPObject = connection

        except:
            PrintException()
            print(sys.exc_info()[0])
            SAPObject = None

try:
    if module == "ClickObjeto":
        # global id_object
        id_object = GetParams('id_object')
        input_ = GetParams('input_')
        tipo = GetParams('tipo')

        #agregar espacios hasta completar 10 caracteres
        #"{:>10}".format(input_)

        if not SAPObject:
            raise Exception("Debe iniciar sesión en SAP")

        connection = SAPObject
        session = SAPObject.Children(0)
        session.findById("wnd[0]").maximize()


        if input_:
            if input_.startswith('"') and "," not in input_:
                input_ = eval(input_)

        if id_object:

            if tipo == "text":
                try:
                    session.findById(id_object).text()
                except:
                    session.findById(id_object).text = input_

            elif tipo == "press":
                session.findById(id_object).press()

            elif tipo == "setFocus":
                session.findById(id_object).setFocus()

            elif tipo == "select":
                session.findById(id_object).select()

            elif tipo == "close":
                session.findById(id_object).close()

            elif tipo == "contextMenu":
                session.findById(id_object).contextMenu()
                time.sleep(2)

            elif tipo == "createSession":
                session.createSession()

            elif tipo == "selectColumn":
                session.findById(id_object).selectColumn(input_)

            elif tipo == "pressContextButton":
                if input_:
                    session.findById(id_object).pressContextButton(input_)
                else:
                    session.findById(id_object).pressContextButton()

            elif tipo == "pressToolbarContextButton":
                if input_:
                    session.findById(id_object).pressToolbarContextButton(input_)
                else:
                    session.findById(id_object).pressToolbarContextButton()

            elif tipo == "pressButton":
                if input_:
                    session.findById(id_object).pressButton(input_)
                else:
                    session.findById(id_object).pressButton()

            elif tipo == "pressToolbarButton":
                session.findById(id_object).pressToolbarButton(input_)

            elif tipo == "currentCellColumn":
                session.findById(id_object).currentCellColumn(input_)

            elif tipo == "selectContextMenuItem":
                session.findById(id_object).selectContextMenuItem(input_)

            elif tipo == "selectedNode":
                session.findById(id_object).selectedNode(input_)

            elif tipo == "selectNode":
                session.findById(id_object).selectNode(input_)

            elif tipo == "selectedRows":
                session.findById(id_object).selectedRows = input_

            elif tipo == "verticalScrollbar":
                try:
                    session.findById(id_object).verticalScrollbar(input_)
                except:
                    session.findById(id_object).verticalScrollbar.position = input_

            elif tipo == "key":
                session.findById(id_object).key = input_
            
            elif tipo == "caretPosition":
                try:
                    session.findById(id_object).caretPosition(input_)
                except:
                    session.findById(id_object).caretPosition = input_
            elif tipo == "expandNode":
                try:
                    session.findById(id_object).expandNode(input_)
                except: 
                    session.findById(id_object).expandNode = input_
            elif tipo == "selectItem":
                try:
                    split_input = input_.split(',')
                    param1 = split_input[0]
                    param2 = split_input[1]
                    if param1.startswith('"'):
                        param1 = eval(param1)
                    if param2.startswith('"'):
                        param2 = eval(param2)
                    session.findById(id_object).selectItem(param1,param2)
                except: 
                     session.findById(id_object).selectItem(input_)
            elif tipo == "ensureVisibleHorizontalItem":
                try:
                    split_input = input_.split(',')
                    param1 = split_input[0]
                    param2 = split_input[1]
                    if param1.startswith('"'):
                        param1 = eval(param1)
                    if param2.startswith('"'):
                        param2 = eval(param2)

                    session.findById(id_object).ensureVisibleHorizontalItem(param1, param2)
                except:
                    PrintException()
                    session.findById(id_object).ensureVisibleHorizontalItem = input_
            elif tipo == "topNode":
                try:
                    session.findById(id_object).topNode(input_)
                except:
                    session.findById(id_object).topNode = input_
            elif tipo == "getAbsoluteRow":
                if not input_:
                    input_ = GetParams("absoluteRow")
                session.findById(id_object).getAbsoluteRow(int(input_)).selected = -1

    if module == "ExtraerTexto":
        id_object = GetParams('id_object')
        var = GetParams('var')

        connection = SAPObject
        session = connection.Children(0)

        if id_object:
            val = session.findById(id_object).text
            SetVar(var,val)


    if module == "click_check":
        id_object = GetParams('id_object')
        tipo = GetParams('tipo')
        todo = GetParams('todo')
        absolute = GetParams('absolute')
        absoluteRow = GetParams('absoluteRow')
        if todo is not None:
            todo = eval(todo)
        if absolute is not None:
            absolute = eval(absolute)
        # print(id_object)

        if not SAPObject:
            raise Exception("Debe iniciar sesión en SAP")

        connection = SAPObject
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()
        
        if id_object and tipo == "marca_":
            if todo == True:
                session.findById(id_object).SelectAllColumns()
            else:
                if absolute == True:
                    session.findById(id_object).getAbsoluteRow(absoluteRow).selected = -1
                else:
                    session.findById(id_object).selected = -1

        if id_object and tipo == "desmarca_":
            if todo == True:
                session.findById(id_object).DeselectAllColumns()
            else:
                if absolute == True:
                    session.findById(id_object).getAbsoluteRow(absoluteRow).selected = 0
                else:
                    session.findById(id_object).selected = -0
            
        

        if not tipo:
            raise Exception("Debe seleccionar una opción")


    if module == "ExtractCell":
        id_object = GetParams('id_object')
        column = GetParams('column')
        row = GetParams('row')
        var = GetParams('var')
        tipo = GetParams('tipo')

        connection = SAPObject
        session = connection.Children(0)
        
        if id_object:
            if tipo == "GetItemText":
                if row.startswith('"'):
                    row = eval(row)
                if column.startswith('"'):
                    column = eval(column)

                val = session.findById(id_object).getitemtext(row, column)
                SetVar(var, val)
            if tipo == "GetCellValue":
                val = session.findById(id_object).getcellvalue(int(row), column)
                SetVar(var, val)


    if module == "ClickCell":
        id_object = GetParams('id_object')
        column = GetParams('column')
        row = GetParams('row')
        click_type = GetParams("type")

        connection = SAPObject
        session = connection.Children(0)

        if not id_object:
            raise Exception("Field 'Id object' is empty")

        if not click_type or click_type == "clickCurrentCell":
            session.findById(id_object).currentCellColumn = column
            session.findById(id_object).selectedRows = row
            session.findById(id_object).clickCurrentCell()
        if click_type == "setCurrentCell":
            session.findById(id_object).setCurrentCell(row, column)
        if click_type == "doubleClickCurrentCell":
            session.findById(id_object).currentCellColumn = column
            session.findById(id_object).selectedRows = row
            session.findById(id_object).doubleClickCurrentCell()
        if click_type == "doubleClickNode":
            session.findById(id_object).doubleClickNode(row)

    if module == "runVBA":
   
        import subprocess

        path = GetParams("path")
        example = GetParams("example")

        if path:
            path = path.replace("/", os.sep)
        if example:
            base_path = tmp_global_obj["basepath"]
            cur_path = base_path + 'modules' + os.sep + 'SAPObjetos' + os.sep + 'libs' + os.sep
            path = cur_path + example

        subprocess.call(r"cscript " + path)


    if module == "checkbox":
        id_object = GetParams('id_object')
        result = GetParams('var')
    
        connection = SAPObject
        session = connection.Children(0)

        if not id_object:
            raise Exception("Field 'Id object' is empty")

        val = session.findById(id_object).selected
        SetVar(result, val)

    if module == "sendKey":
        id_object = GetParams('id_object')
        key = GetParams('key')

        if not SAPObject:
            raise Exception("Debe iniciar sesión en SAP")

        connection = SAPObject
        session = connection.Children(0)

        if not id_object:
            raise Exception("Field 'Id object' is empty")

        if key.isdigit():
            key = int(key)

        session.findById(id_object).sendVKey(key)


    if module == "GetProperty":
        
        def GetProperty(object, property):
            properties = {
                "Height": object.Height, 
                "Highlighted": object.Highlighted,
                "Name": object.Name,
                "Required": object.Required,
                "Text": object.Text,
                "Width": object.Width,
              }
            return properties[property]

        id_object = GetParams('id_object')
        property = GetParams('property')
        result = GetParams('result')

        connection = SAPObject
        session = connection.Children(0)
        res = GetProperty(session.findById(id_object), property)
        SetVar(result, res)

    
except Exception as e:
    PrintException()
    raise e
