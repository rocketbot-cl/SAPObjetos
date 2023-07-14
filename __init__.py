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
from time import time, sleep
import os
import traceback

GetParams = GetParams #type:ignore
tmp_global_obj = tmp_global_obj #type:ignore
PrintException = PrintException #type:ignore
SetVar = SetVar #type:ignore
GetGlobals = GetGlobals #type:ignore

# Add modules libraries to Rocektbot
# -----------------------------------
base_path = tmp_global_obj["basepath"]
cur_path = os.path.join(base_path, 'modules', 'SAPObjetos', 'libs')

if cur_path not in sys.path:
    sys.path.append(cur_path)

try:
    from threading import Thread
    from queue import Queue
except ImportError:
    from Queue import Queue
    
"""
    Obtengo el modulo que fue invocado
"""
module = GetParams("module")

global SAPObject
global id_object
global SAPObjetos_mod
global SAP_session

def open_sap(path):
    global subprocess
    subprocess.Popen(path)
    
def waitForObject(session, id, timeout=10):
    global time, sleep
    inicio = time()
    while True:
        try:
            return session.findById(id)
        except Exception as e:
            pass
        sleep(1)
        fin = time()
        total = fin - inicio
        if total > timeout:
            return session.findById(id)


timeout = GetParams("timeout")
if not timeout:
    timeout = 10
else:
    timeout = int(timeout)
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
    # result = GetParams('result')
    """
        validaciones
    """
    if not id_user or not id_pass or not conn or not path:
        raise Exception('No ha ingresado todos los datos')

    if id_user:
        try:
            if not path:
                path = "C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe"

            q = Queue()
            t = Thread(target=open_sap, args=(path,))
            t.start()

            inicio = time()
            while True:
                try:
                    # print("Get Object 'SAPGUI'")
                    SapGuiAuto = client.GetObject('SAPGUI')
                    application = SapGuiAuto.GetScriptingEngine
                    # print("Open connection")
                    SAPObjetos_mod = application.OpenConnection(conn, True)
                    break
                except Exception as e:
                    pass

                sleep(1)
                fin = time()
                total = fin - inicio
                if total > 60:
                    # SetVar(result, False)
                    raise Exception("Timeout: SAP cannot be opened")
                
            SAP_session = None
            try:
                SAP_session = SAPObjetos_mod.Children(0)
            except:
                SAP_session = SAPObjetos_mod.Children(1)

            if user and password:
                SelectedObj = waitForObject(SAP_session, id_user, timeout)
                SelectedObj.SetFocus()
                SelectedObj.text = user
                SelectedObj = waitForObject(SAP_session, id_pass, timeout)
                SelectedObj.text = password

            if SAPObjetos_mod:
                SAPObject = SAPObjetos_mod

        except Exception as e:
            traceback.print_exc()
            PrintException()
            print(sys.exc_info()[0])
            SAPObject = None
            raise e

try:
    id_object = GetParams('id_object')
    if module != "LoginSap" and module != "wait_object":
        waitForObject(SAP_session, "wnd[0]").maximize()
        SelectedObj = waitForObject(SAP_session, id_object, timeout)
    
    if module == "ClickObjeto":
        # id_object = GetParams('id_object')
        column = GetParams('column')
        row = GetParams('row')
        input_ = GetParams('input_')
        tipo = GetParams('tipo')

        #agregar espacios hasta completar 10 caracteres
        #"{:>10}".format(input_)

        if not SAPObject:
            raise Exception("Debe iniciar sesión en SAP")

        #waitForObject(SAP_session, "wnd[0]").maximize()

        if column not in ["", None] and row not in ["", None]:
            try:
                SelectedObj.currentCellColumn = column
                SelectedObj.selectedRows = row
            except:
                pass

        if input_:
            if input_.startswith('"') and "," not in input_:
                input_ = eval(input_)

        if id_object:

            if tipo == "text":
                try:
                    SelectedObj.text()
                except:
                    SelectedObj.text = input_

            elif tipo == "press":
                SelectedObj.press()

            elif tipo == "setFocus":
                SelectedObj.setFocus()

            elif tipo == "select":
                SelectedObj.select()

            elif tipo == "close":
                SelectedObj.close()

            elif tipo == "contextMenu":
                SelectedObj.contextMenu()
                sleep(2)

            elif tipo == "createSession":
                SAP_session.createSession()

            elif tipo == "selectColumn":
                SelectedObj.selectColumn(input_)

            elif tipo == "pressContextButton":
                if input_:
                    SelectedObj.pressContextButton(input_)
                else:
                    SelectedObj.pressContextButton()

            elif tipo == "pressToolbarContextButton":
                if input_:
                    SelectedObj.pressToolbarContextButton(input_)
                else:
                    SelectedObj.pressToolbarContextButton()

            elif tipo == "pressButton":
                if input_:
                    SelectedObj.pressButton(input_)
                else:
                    SelectedObj.pressButton()

            elif tipo == "pressToolbarButton":
                SelectedObj.pressToolbarButton(input_)

            elif tipo == "currentCellColumn":
                SelectedObj.currentCellColumn(input_)

            elif tipo == "selectContextMenuItem":
                SelectedObj.selectContextMenuItem(input_)

            elif tipo == "selectedNode":
                SelectedObj.selectedNode(input_)

            elif tipo == "selectNode":
                SelectedObj.selectNode(input_)

            elif tipo == "selectedRows":
                SelectedObj.selectedRows = input_

            elif tipo == "verticalScrollbar":
                try:
                    SelectedObj.verticalScrollbar(input_)
                except:
                    SelectedObj.verticalScrollbar.position = input_

            elif tipo == "key":
                SelectedObj.key = input_
            
            elif tipo == "caretPosition":
                try:
                    SelectedObj.caretPosition(input_)
                except:
                    SelectedObj.caretPosition = input_
            elif tipo == "expandNode":
                try:
                    SelectedObj.expandNode(input_)
                except: 
                    SelectedObj.expandNode = input_
            elif tipo == "selectItem":
                try:
                    split_input = input_.split(',')
                    param1 = split_input[0]
                    param2 = split_input[1]
                    if param1.startswith('"'):
                        param1 = eval(param1)
                    if param2.startswith('"'):
                        param2 = eval(param2)
                    SelectedObj.selectItem(param1,param2)
                except: 
                     SelectedObj.selectItem(input_)
            elif tipo == "ensureVisibleHorizontalItem":
                try:
                    split_input = input_.split(',')
                    param1 = split_input[0]
                    param2 = split_input[1]
                    if param1.startswith('"'):
                        param1 = eval(param1)
                    if param2.startswith('"'):
                        param2 = eval(param2)

                    SelectedObj.ensureVisibleHorizontalItem(param1, param2)
                except:
                    PrintException()
                    SelectedObj.ensureVisibleHorizontalItem = input_
            elif tipo == "topNode":
                try:
                    SelectedObj.topNode(input_)
                except:
                    SelectedObj.topNode = input_
            elif tipo == "getAbsoluteRow":
                if not input_:
                    input_ = GetParams("absoluteRow")
                SelectedObj.getAbsoluteRow(int(input_)).selected = -1
            elif tipo == "doubleClickItem":
                SelectedObj.doubleClickItem(row, column)

    if module == "ExtraerTexto":
        # id_object = GetParams('id_object')
        var = GetParams('var')

        if id_object:
            val = SelectedObj.text
            SetVar(var,val)

    if module == "click_check":
        # id_object = GetParams('id_object')
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
        
        # waitForObject(SAP_session, "wnd[0]").maximize()
        
        if id_object and tipo == "marca_":
            if todo == True:
                SelectedObj.SelectAllColumns()
            else:
                if absolute == True:
                    SelectedObj.getAbsoluteRow(absoluteRow).selected = -1
                else:
                    SelectedObj.selected = -1

        if id_object and tipo == "desmarca_":
            if todo == True:
                SelectedObj.DeselectAllColumns()
            else:
                if absolute == True:
                    SelectedObj.getAbsoluteRow(absoluteRow).selected = 0
                else:
                    SelectedObj.selected = -0   

        if not tipo:
            raise Exception("Debe seleccionar una opción")

    if module == "ExtractCell":
        # id_object = GetParams('id_object')
        column = GetParams('column')
        row = GetParams('row')
        var = GetParams('var')
        tipo = GetParams('tipo')

        if id_object:
            if tipo == "GetItemText":
                if row.startswith('"'):
                    row = eval(row)
                if column.startswith('"'):
                    column = eval(column)

                val = SelectedObj.getitemtext(row, column)
                SetVar(var, val)
            if tipo == "GetCellValue":
                val = SelectedObj.getcellvalue(int(row), column)
                SetVar(var, val)

    if module == "ClickCell":
        # id_object = GetParams('id_object')
        column = GetParams('column')
        row = GetParams('row')
        click_type = GetParams("type")

        if not id_object:
            raise Exception("Field 'Id object' is empty")

        if not click_type or click_type == "clickCurrentCell":
            SelectedObj.currentCellColumn = column
            SelectedObj.selectedRows = row
            SelectedObj.clickCurrentCell()
        if click_type == "setCurrentCell":
            SelectedObj.setCurrentCell(row, column)
        if click_type == "doubleClickCurrentCell":
            SelectedObj.currentCellColumn = column
            SelectedObj.selectedRows = row
            SelectedObj.doubleClickCurrentCell()
        if click_type == "doubleClickNode":
            SelectedObj.doubleClickNode(row)

    if module == "runVBA":

        path = GetParams("path")
        example = GetParams("example")

        if path:
            path = "\"" + path.replace("/", os.sep) + "\""
        if example:
            base_path = tmp_global_obj["basepath"]
            cur_path = base_path + 'modules' + os.sep + 'SAPObjetos' + os.sep + 'libs' + os.sep
            path = cur_path + example

        p = subprocess.Popen([r"C:\Windows\System32\cscript.exe ", path], shell=True)
        sleep(5)
        print(p.communicate())
        
    if module == "wait_object":
        # ProcessTime = time.perf_counter  #this returns nearly 0 when first call it if python version <= 3.6
        # timer = ProcessTime()
        # actual = timer
        item = GetParams('item')
        var_ = GetParams('result') 

        try:
            result = waitForObject(SAP_session, item, timeout) 
            if result:
                SetVar(var_, True)
        except Exception as e:
            # print(e)
            SetVar(var_, False)

    if module == "checkbox":
        # id_object = GetParams('id_object')
        result = GetParams('var')

        if not id_object:
            raise Exception("Field 'Id object' is empty")

        val = SelectedObj.selected
        SetVar(result, val)

    if module == "sendKey":
        # id_object = GetParams('id_object')
        column = GetParams('column')
        row = GetParams('row')
        key = GetParams('key')

        if not SAPObject:
            raise Exception("Debe iniciar sesión en SAP")

        if not id_object:
            raise Exception("Field 'Id object' is empty")

        if key.isdigit():
            key = int(key)

        if column not in ["", None] and row not in ["", None]:
            SelectedObj.currentCellColumn = column
            SelectedObj.selectedRows = row
        
        SelectedObj.sendVKey(key)

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

        # id_object = GetParams('id_object')
        property = GetParams('property')
        result = GetParams('result')

        res = GetProperty(SelectedObj, property)
        SetVar(result, res)
    
except Exception as e:
    traceback.print_exc()
    PrintException()
    raise e
