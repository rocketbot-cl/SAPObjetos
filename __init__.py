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
            path = path
            subprocess.Popen(path)
            time.sleep(10)

            SapGuiAuto = client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.OpenConnection(conn, True)
            session = connection.Children(0)

            session.findById(id_user).text = user
            session.findById(id_pass).text = password

            if connection:
                SAPObject = connection

        except:
            PrintException()
            print(sys.exc_info()[0])
            SAPObject = None


if module == "ClickObjeto":
    # global id_object
    id_object = GetParams('id_object')
    input_ = GetParams('input_')
    tipo = GetParams('tipo')

    #agregar espacios hasta completar 10 caracteres
    #"{:>10}".format(input_)

    if not SAPObject:
        raise Exception("Debe iniciar sesión en SAP")

    try:

        connection = SAPObject
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()

        if input_:
            if input_.startswith('"'):
                input_ = eval(input_)

        if id_object:

            if tipo == "text":
                try:
                    session.findById(id_object).text()
                except:
                    session.findById(id_object).text = input_

            if tipo == "press":
                session.findById(id_object).press()

            if tipo == "setFocus":
                session.findById(id_object).setFocus()

            if tipo == "select":
                session.findById(id_object).select()

            if tipo == "close":
                session.findById(id_object).close()

            if tipo == "contextMenu":
                session.findById(id_object).contextMenu()
                time.sleep(2)

            if tipo == "createSession":
                session.createSession()

            if tipo == "selectColumn":
                session.findById(id_object).selectColumn(input_)

            if tipo == "pressContextButton":
                if input_:
                    session.findById(id_object).pressContextButton(input_)
                else:
                    session.findById(id_object).pressContextButton()

            if tipo == "pressToolbarContextButton":
                if input_:
                    session.findById(id_object).pressToolbarContextButton(input_)
                else:
                    session.findById(id_object).pressToolbarContextButton()

            if tipo == "pressButton":
                if input_:
                    session.findById(id_object).pressButton(input_)
                else:
                    session.findById(id_object).pressButton()

            if tipo == "pressToolbarButton":
                session.findById(id_object).pressToolbarButton(input_)

            if tipo == "currentCellColumn":
                session.findById(id_object).currentCellColumn(input_)

            if tipo == "selectContextMenuItem":
                session.findById(id_object).selectContextMenuItem(input_)

            if tipo == "selectedNode":
                session.findById(id_object).selectedNode(input_)

            if tipo == "selectNode":
                session.findById(id_object).selectNode(input_)

            if tipo == "selectedRows":
                session.findById(id_object).selectedRows = input_

            if tipo == "verticalScrollbar":
                session.findById(id_object).verticalScrollbar(input_)
            if tipo == "key":
                if input_:
                    session.findById(id_object).key = input_
                else:
                    session.findById(id_object).key()

            if tipo == "key":
                session.findById(id_object).key = input_

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e


if module == "ExtraerTexto":
    id_object = GetParams('id_object')
    var = GetParams('var')

    try:
        connection = SAPObject
        session = connection.Children(0)

        if id_object:
            val = session.findById(id_object).text
            SetVar(var,val)


    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "click_check":
    id_object = GetParams('id_object')
    tipo = GetParams('tipo')
    # print(id_object)

    if not SAPObject:
        raise Exception("Debe iniciar sesión en SAP")

    try:
        connection = SAPObject
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()

        if id_object and tipo == "marca_":
            session.findById(id_object).selected = -1

        if id_object and tipo == "desmarca_":
            session.findById(id_object).selected = 0

        if not tipo:
            raise Exception("Debe seleccionar una opción")

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "ExtractCell":
    id_object = GetParams('id_object')
    column = GetParams('column')
    row = GetParams('row')
    var = GetParams('var')

    try:
        connection = SAPObject
        session = connection.Children(0)

        if id_object:
            val = session.findById(id_object).getcellvalue(int(row), column)
            SetVar(var, val)


    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "ClickCell":
    id_object = GetParams('id_object')
    column = GetParams('column')
    row = GetParams('row')
    click_type = GetParams("type")

    try:
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

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "runVBA":
    try:
        import subprocess

        path = GetParams("path")
        example = GetParams("example")

        if example:
            base_path = tmp_global_obj["basepath"]
            cur_path = base_path + 'modules' + os.sep + 'SAPObjetos' + os.sep + 'libs' + os.sep
            path = cur_path + example

        subprocess.call(r"cscript " + path)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "checkbox":
    id_object = GetParams('id_object')
    result = GetParams('var')
    try:
        connection = SAPObject
        session = connection.Children(0)

        if not id_object:
            raise Exception("Field 'Id object' is empty")

        val = session.findById(id_object).selected
        SetVar(result, val)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
