VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ENTRADA_DATA 
   Caption         =   "Ingresar Datos"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "F_ENTRADA_DATA.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "F_ENTRADA_DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Evento que se dispara cuando cambia el valor del SpinButton de Cantidad (BOTON_AUMENTO).
' Sincroniza el valor num�rico del SpinButton con el texto mostrado en el TextBox CANTIDAD.
Private Sub BOTON_AUMENTO_Change()
    Me.CANTIDAD.Value = CStr(Me.BOTON_AUMENTO.Value)
End Sub

' Bot�n para AGREGAR un NUEVO REGISTRO a la siguiente fila disponible.
' No se usa para actualizar registros existentes.
Private Sub GUARDAR_Click()
    Dim ws As Worksheet
    Dim ultimaFilaConDatos As Long
    Dim filaParaGuardar As Long
    Dim estadoEntrega As String

    Set ws = Worksheets("BOTELLAS")

    ' Validaci�n b�sica: Asegura que el ID no est� vac�o antes de guardar.
    If Trim(Me.ID_DATO.Value) = "" Then
        MsgBox "El ID del registro no puede estar vac�o.", vbExclamation
        Me.ID_DATO.SetFocus
        Exit Sub
    End If

    ' Determina la pr�xima fila libre para guardar.
    ultimaFilaConDatos = ws.Cells(Rows.Count, 1).End(xlUp).Row
    filaParaGuardar = ultimaFilaConDatos + 1

    ' Obtiene el estado de la casilla 'Entregada'.
    If Me.ENTREGADA.Value = True Then
        estadoEntrega = "SI"
    Else
        estadoEntrega = "NO"
    End If

    ' --- Guarda los datos de los controles en la nueva fila ---
    ' Aseg�rate de que las columnas B a H coincidan con tus campos (Persona, Tipo, Cantidad, Notas, Fecha, Entregada, Fecha_Entrega).
    ws.Range("A" & filaParaGuardar).Value = Me.ID_DATO.Value
    ws.Range("B" & filaParaGuardar).Value = Me.cbxPersona.Value ' Usa el ComboBox Persona
    ws.Range("C" & filaParaGuardar).Value = Me.TIPO_BOTELLA.Value
    ws.Range("D" & filaParaGuardar).Value = Me.CANTIDAD.Value
    ws.Range("E" & filaParaGuardar).Value = Me.NOTAS.Value
    ws.Range("F" & filaParaGuardar).Value = Me.FECHA.Value
    ws.Range("G" & filaParaGuardar).Value = estadoEntrega
    ws.Range("H" & filaParaGuardar).Value = Me.FECHA_ENTREGA.Value

    MsgBox "Registro guardado correctamente en la fila " & filaParaGuardar, vbInformation

    ' --- Despu�s de guardar un nuevo registro ---
    ' Actualiza los l�mites del SpinButton de navegaci�n para incluir la nueva fila.
    ' Posiciona el SpinButton en la fila reci�n guardada (esto activar� spnRegistro_Change para cargarla).
    ultimaFilaConDatos = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim primeraFilaDatos As Long
    primeraFilaDatos = 2

    If ultimaFilaConDatos >= primeraFilaDatos Then
        With Me.spnRegistro
            .Min = primeraFilaDatos
            .Max = ultimaFilaConDatos
            .Value = ultimaFilaConDatos ' Carga el registro guardado (llama a spnRegistro_Change)
            .Enabled = True             ' Habilita la navegaci�n si hay datos
        End With
    Else
         ' Este caso es poco probable despu�s de guardar, pero maneja si la hoja queda vac�a.
         Me.spnRegistro.Enabled = False
         Call LimpiarYPrepararParaNuevoRegistro
    End If

    ' El estado de los botones se ajusta dentro de spnRegistro_Change.
    ' El formulario ahora muestra el registro reci�n guardado.

End Sub


' Prepara el formulario para la entrada de un NUEVO REGISTRO.
' Calcula el pr�ximo ID, limpia campos y gestiona el estado de los botones.
Sub LimpiarYPrepararParaNuevoRegistro()
    Dim ws As Worksheet
    Dim ultimaFilaConDatos As Long

    Set ws = Worksheets("BOTELLAS")

    ' Calcula el pr�ximo ID disponible.
    ultimaFilaConDatos = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim primeraFilaDatos As Long
    primeraFilaDatos = 2

    If ultimaFilaConDatos >= primeraFilaDatos And IsNumeric(ws.Cells(ultimaFilaConDatos, "A").Value) Then
        Me.ID_DATO.Value = ws.Cells(ultimaFilaConDatos, "A").Value + 1
    Else
        Me.ID_DATO.Value = 1
    End If

    ' --- Limpia o establece valores por defecto para los campos de entrada ---
    Me.cbxPersona.Value = ""      ' Limpia el ComboBox Persona
    Me.TIPO_BOTELLA.Value = ""     ' Limpia el ComboBox Tipo
    Me.NOTAS.Value = ""            ' Limpia el TextBox Notas

    ' Los campos de fecha quedan vac�os por defecto al preparar nuevo.
    ' La fecha de hoy est� en el desplegable (configurado en Initialize).
    Me.FECHA.Value = ""
    Me.FECHA_ENTREGA.Value = ""

    Me.ENTREGADA.Value = False     ' Desmarca la casilla Entregada por defecto

    ' Establece la cantidad por defecto a 1 y sincroniza el SpinButton.
    Me.BOTON_AUMENTO.Value = 1
    Me.CANTIDAD.Value = "1"

    ' --- Gestionar estado de los botones para el modo "Nuevo" ---
    Me.spnRegistro.Enabled = False   ' Deshabilita navegaci�n
    Me.btxActualizar.Enabled = False ' Deshabilita Actualizar
    Me.GUARDAR.Enabled = True        ' Habilita Guardar (Nuevo)
    ' --- Fin Gesti�n estado de botones ---

    ' Pone el cursor en el primer campo de entrada.
    Me.cbxPersona.SetFocus

End Sub


' Carga los datos del registro correspondiente al valor actual del SpinButton de navegaci�n (spnRegistro).
' Gestiona el estado de los botones.
Private Sub spnRegistro_Change()
    Dim ws As Worksheet
    Dim filaActual As Long

    Set ws = Worksheets("BOTELLAS")

    ' Obtiene el n�mero de fila desde el SpinButton.
    filaActual = Me.spnRegistro.Value

    ' --- Carga los datos de la filaActual en los controles ---
    ' Usa On Error Resume Next para evitar errores si alguna celda est� vac�a o tiene tipo incorrecto.
    ' Considera validaciones de tipo m�s espec�ficas si es necesario.
    On Error Resume Next

    ' Mapeo de Columnas: A=ID, B=PERSONA, C=Tipo, D=Cantidad, E=NOTAS, F=Fecha, G=Entregada(SI/NO), H=Fecha_Entrega
    Me.ID_DATO.Value = ws.Cells(filaActual, "A").Value
    Me.cbxPersona.Value = ws.Cells(filaActual, "B").Value    ' Carga en ComboBox Persona
    Me.TIPO_BOTELLA.Value = ws.Cells(filaActual, "C").Value

    ' Cantidad: Carga y sincroniza el SpinButton de Cantidad.
    Me.CANTIDAD.Value = ws.Cells(filaActual, "D").Value
    If IsNumeric(Me.CANTIDAD.Value) Then
         Me.BOTON_AUMENTO.Value = CDbl(Me.CANTIDAD.Value)
    Else
         Me.BOTON_AUMENTO.Value = Me.BOTON_AUMENTO.Min ' Si no es num�rico, sincroniza a Min (1)
    End If

    Me.NOTAS.Value = ws.Cells(filaActual, "E").Value

    ' Fecha: Carga si es una fecha v�lida.
    If IsDate(ws.Cells(filaActual, "F").Value) Then
        Me.FECHA.Value = ws.Cells(filaActual, "F").Value
    Else
        Me.FECHA.Value = ""
    End If

    ' Entregada: Carga el estado de la casilla desde "SI"/"NO".
    If UCase(Trim(ws.Cells(filaActual, "G").Value)) = "SI" Then
        Me.ENTREGADA.Value = True
    Else
        Me.ENTREGADA.Value = False
    End If

    ' Fecha Entrega: Carga si es una fecha v�lida.
     If IsDate(ws.Cells(filaActual, "H").Value) Then
        Me.FECHA_ENTREGA.Value = ws.Cells(filaActual, "H").Value
    Else
        Me.FECHA_ENTREGA.Value = ""
    End If

    On Error GoTo 0

    ' Opcional: Muestra la fila actual/total.
    ' Me.lblEstadoRegistro.Caption = "Registro " & filaActual & " de " & Me.spnRegistro.Max

    ' --- Gestionar estado de los botones para el modo "Navegaci�n/Edici�n" ---
    Me.spnRegistro.Enabled = True   ' Navegaci�n habilitada
    Me.btxActualizar.Enabled = True ' Actualizar habilitado (hay registro cargado)
    Me.GUARDAR.Enabled = False      ' Guardar (Nuevo) deshabilitado
    ' --- Fin Gesti�n estado de botones ---

End Sub


' Bot�n para GUARDAR CAMBIOS en el registro actualmente mostrado por spnRegistro.
' Requiere que un registro est� cargado (spnRegistro habilitado).
Private Sub btxActualizar_Click()
    Dim ws As Worksheet
    Dim filaActual As Long
    Dim estadoEntrega As String

    Set ws = Worksheets("BOTELLAS")

    ' --- Validar si hay un registro cargado para actualizar ---
    Dim primeraFilaDatos As Long
    primeraFilaDatos = 2
    If Me.spnRegistro.Enabled = False Or Me.spnRegistro.Value < primeraFilaDatos Then
         MsgBox "No hay un registro existente v�lido seleccionado para actualizar.", vbExclamation
         Exit Sub
    End If

    ' La fila a actualizar es la que indica el SpinButton de navegaci�n.
    filaActual = Me.spnRegistro.Value

    ' --- Validar datos de entrada antes de guardar ---
    If Trim(Me.cbxPersona.Value) = "" Then
         MsgBox "El campo Persona no puede estar vac�o.", vbExclamation
         Me.cbxPersona.SetFocus
         Exit Sub
    End If
     If Trim(Me.CANTIDAD.Value) = "" Or Not IsNumeric(Me.CANTIDAD.Value) Then
         MsgBox "El campo Cantidad debe ser un n�mero v�lido.", vbExclamation
         Me.CANTIDAD.SetFocus
         Exit Sub
     End If
    ' A�ade validaciones para otros campos (fechas si son obligatorias, etc.).

    ' Obtiene el estado de la casilla de verificaci�n.
    If Me.ENTREGADA.Value = True Then
        estadoEntrega = "SI"
    Else
        estadoEntrega = "NO"
    End If

    ' --- Guarda los datos de los controles SOBRE la filaActual en la hoja ---
    ' NO cambies el ID (Columna A) al actualizar un registro existente.
    ' ws.Range("A" & filaActual).Value = Me.ID_DATO.Value ' <-- No hacer esto

    ' Actualiza las otras columnas.
    ws.Range("B" & filaActual).Value = Me.cbxPersona.Value
    ws.Range("C" & filaActual).Value = Me.TIPO_BOTELLA.Value
    ws.Range("D" & filaActual).Value = Me.CANTIDAD.Value
    ws.Range("E" & filaActual).Value = Me.NOTAS.Value
    ws.Range("F" & filaActual).Value = Me.FECHA.Value
    ws.Range("G" & filaActual).Value = estadoEntrega
    ws.Range("H" & filaActual).Value = Me.FECHA_ENTREGA.Value

    MsgBox "Registro en fila " & filaActual & " actualizado correctamente.", vbInformation

    ' Opcional: Recarga el registro para reflejar cualquier formateo o conversi�n de guardado.
    ' Call spnRegistro_Change

End Sub


' Inicia el modo de entrada de un NUEVO REGISTRO.
Private Sub btnNuevo_Click() ' Suponiendo que este es el nombre del bot�n "Nuevo".
    ' Llama al sub para limpiar campos y preparar el formulario.
    Call LimpiarYPrepararParaNuevoRegistro
    ' El estado de los botones se gestiona dentro de LimpiarYPrepararParaNuevoRegistro.
End Sub


' Configuraci�n inicial del formulario al abrirse.
' Llena las listas, configura SpinButtons y carga el �ltimo registro (o prepara para nuevo).
' Gestiona el estado inicial de los botones.
Private Sub UserForm_Initialize()
    Dim fecha_hoy As Date
    Dim ws As Worksheet
    Dim ultimaFilaConDatos As Long ' �ltima fila con datos (generalmente Col A)
    Dim primeraFilaDatos As Long   ' Primera fila donde empiezan los datos (ej. 2)
    Dim ultimaFilaB As Long        ' �ltima fila con datos en Col B (para lista Persona)
    Dim dictUnicos As Object       ' Para nombres �nicos en Col B
    Dim arrUnicos As Variant       ' Array de nombres �nicos para ComboBox

    Set ws = Worksheets("BOTELLAS")
    fecha_hoy = Date

    ' --- Llenar ComboBox Persona con valores �nicos de la Columna B ---
    ' Requiere habilitar la referencia: Herramientas > Referencias... > Microsoft Scripting Runtime.
    On Error Resume Next
    Set dictUnicos = CreateObject("Scripting.Dictionary")
    On Error GoTo 0

    ' Contin�a solo si el objeto Dictionary se cre� correctamente.
    If Not dictUnicos Is Nothing Then
        primeraFilaDatos = 2 ' Asumimos la fila 1 es encabezado

        ' Encuentra la �ltima fila en Col B.
        ultimaFilaB = ws.Cells(Rows.Count, "B").End(xlUp).Row

        ' Recorre Col B (saltando el encabezado) y a�ade nombres �nicos al diccionario.
        If ultimaFilaB >= primeraFilaDatos Then
            Dim cell As Range
            For Each cell In ws.Range("B" & primeraFilaDatos & ":B" & ultimaFilaB)
                If Trim(cell.Value) <> "" Then
                    If Not dictUnicos.Exists(cell.Value) Then
                        dictUnicos.Add Key:=cell.Value, Item:=cell.Value
                    End If
                End If
            Next cell
        End If

        ' Asigna la lista de nombres �nicos al ComboBox Persona.
        If dictUnicos.Count > 0 Then
            arrUnicos = dictUnicos.Keys
            ' Opcional: Si quieres ordenar arrUnicos, llama a una sub de ordenamiento aqu�.
            Me.cbxPersona.List = arrUnicos
        Else
            Me.cbxPersona.List = Array() ' Asigna una lista vac�a si no hay nombres.
        End If

        Set dictUnicos = Nothing ' Libera la memoria.
    Else
        ' Muestra error si la referencia no est� habilitada y el diccionario no se pudo crear.
        MsgBox "Error: La referencia 'Microsoft Scripting Runtime' no est� habilitada.", vbCritical
        ' Considera salir del sub o deshabilitar el ComboBox si es cr�tico.
        ' Me.cbxPersona.Enabled = False
    End If
    ' --- Fin Llenar ComboBox Persona ---


    ' --- Configuraci�n del SpinButton de Cantidad (BOTON_AUMENTO) ---
    With Me.BOTON_AUMENTO
        .Min = 1
        .Max = 24 ' Ajusta este m�ximo.
        .SmallChange = 1
        ' El valor inicial se establece al cargar o preparar nuevo.
    End With

    ' --- Configuraci�n del SpinButton de Navegaci�n (spnRegistro) ---
    ' Encuentra la �ltima fila general (usando Col A como referencia).
    ultimaFilaConDatos = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Decide si cargar el �ltimo registro o preparar para uno nuevo.
    If ultimaFilaConDatos < primeraFilaDatos Then
        ' No hay datos para navegar.
        Me.spnRegistro.Enabled = False
        ' Prepara el formulario para el primer registro.
        Call LimpiarYPrepararParaNuevoRegistro ' Esto tambi�n gestiona el estado de los botones.
        MsgBox "No hay registros existentes para mostrar. Preparado para a�adir el primer registro.", vbInformation
    Else
        ' Hay datos. Configura el SpinButton de navegaci�n.
        Me.spnRegistro.Enabled = True
        With Me.spnRegistro
            .Min = primeraFilaDatos
            .Max = ultimaFilaConDatos
            .SmallChange = 1
            .Value = .Max ' Carga el �LTIMO registro al abrir (llama a spnRegistro_Change).
        End With
        ' El estado de los botones se gestiona dentro de spnRegistro_Change en este caso inicial.
    End If
    ' --- Fin Configuraci�n spnRegistro ---

    ' --- Llenar Listas para ComboBoxes de Fecha ---
    ' Esto hace que la fecha de hoy sea seleccionable en el desplegable.
    ' El valor mostrado en el ComboBox es independiente.
    Me.FECHA.List = Array(fecha_hoy)
    Me.FECHA_ENTREGA.List = Array(fecha_hoy)
    ' --- Fin Llenar Listas Fecha ---

    ' --- Llenar otras ListBoxes ---
    Me.TIPO_BOTELLA.List = Array("COCA 1 1/4 VIDRIO", "CORONA FAMILIRAR", "VICTORIA 355ML VIDRIO")

    ' El estado inicial final de los botones ya fue establecido por LimpiarYPrepararParaNuevoRegistro
    ' o por la llamada a spnRegistro_Change.

End Sub