' =====================================================================
' ARCHIVO      : depurar.vbs
' UBICACION    : C:\pruebas\code\depurar.vbs   (raiz del codigo)
' DESTINO      : C:\excel\poblacion_latam.xlsx (archivo Excel generado)
' PROPOSITO    : Crear un libro de Excel con los paises de Latinoamerica
'                (columna A) y su poblacion (columna B). Si el archivo ya
'                existe, lo abre y lo ACTUALIZA sin borrarlo ni recrearlo.
' RESPONSABILIDAD: Resolver rutas relativas, abrir/crear el libro,
'                escribir datos, dar formato, guardar y liberar Excel.
' ALCANCE      : Fase 1 (solo paises + poblacion). No lee fuentes externas;
'                los datos estan embebidos (aprox. ONU WPP 2024, redondeados).
' RESTRICCIONES: Solo componentes nativos de Windows (WSH + VBScript +
'                Excel COM). Sin instalar nada, sin librerias de terceros.
' NOTA         : Los comentarios van en ASCII (sin tildes) porque WSH lee
'                los .vbs como ANSI; las tildes de los datos se arman con
'                ChrW() para evitar caracteres corruptos.
' EJECUCION    : doble clic, o:  cscript //nologo depurar.vbs
' =====================================================================
'
' ---------------------------------------------------------------------
' JUSTIFICACION TECNICA (por que este enfoque)
' ---------------------------------------------------------------------
' - Automatizacion COM de Excel (CreateObject("Excel.Application")):
'     * Forma tradicional: escribir un .csv. Problema: no es un .xlsx real,
'       pierde formato, abre con separadores segun la region (; o ,) y
'       corrompe tildes segun la codificacion.
'     * Forma actual: COM crea/actualiza un .xlsx real con formato numerico.
'       Funciona con Office 365 ya instalado y no requiere permisos de admin.
' - Rutas relativas a WScript.ScriptFullName:
'     * Forma tradicional: rutas absolutas "C:\..." fijas. Problema: se
'       rompen al mover la carpeta o usar otro equipo/unidad.
'     * Forma actual: se calcula la ruta desde la ubicacion del propio
'       script (..\..\excel), portable mientras se conserve la estructura.
' - Option Explicit: obliga a declarar variables; evita errores por typos
'   (VBScript crea variables nuevas en silencio sin esta directiva).
' - Patron "try/finally" con On Error Resume Next + procedimiento aparte:
'   garantiza cerrar Excel aunque falle algo. Sin esto quedan procesos
'   EXCEL.EXE huerfanos que bloquean el archivo.
' - Actualizar = abrir y sobrescribir celdas (no borrar el archivo): cumple
'   la regla de no eliminar archivos y conserva lo que el usuario agregue
'   fuera del rango A1:B21.
'
' ---------------------------------------------------------------------
' FUTURAS FUNCIONALIDADES / ESTADO DEL ECOSISTEMA
' ---------------------------------------------------------------------
' - Deprecacion de VBScript: Microsoft la anuncio (2024). Pasa a ser
'   "Feature on Demand" en Windows 11 24H2 y se planea desactivarla por
'   defecto y luego retirarla en versiones futuras (sin fecha cerrada
'   publica; verificar la hoja de ruta oficial). Alternativas nativas que
'   cubren este caso: PowerShell (ya incluido en Windows) y Office Scripts
'   (Excel web / Power Automate). Si la politica corporativa lo permite,
'   esta logica deberia migrar a PowerShell.
' - Excel JavaScript API / Office Scripts: resuelven automatizacion de Excel
'   sin COM; estan disponibles hoy (GA), sujetas a licencia y politicas.
' ---------------------------------------------------------------------

Option Explicit                                         ' Obliga a declarar todas las variables

Const xlOpenXMLWorkbook = 51                            ' Codigo de formato .xlsx (en late binding no hay constantes de Excel)

Dim objFso                                              ' Objeto FileSystemObject para manejar rutas y carpetas
Dim objExcel                                            ' Objeto aplicacion Excel (global para poder cerrarlo en el cierre)
Dim objLibro                                            ' Objeto libro de trabajo (global por la misma razon)

Set objFso = CreateObject("Scripting.FileSystemObject") ' Crea el objeto de sistema de archivos nativo de Windows
Set objExcel = Nothing                                  ' Inicializa Excel como vacio para saber luego si se llego a crear
Set objLibro = Nothing                                  ' Inicializa el libro como vacio por el mismo motivo

On Error Resume Next                                    ' Desde aqui los errores no detienen el script; se revisan manualmente
Call Ejecutar()                                         ' Ejecuta la logica; si falla, el error sube hasta aqui sin abortar
Dim strError                                            ' Variable para guardar el texto del error (si lo hay)
strError = ""                                           ' Se asume que no hubo error
If Err.Number <> 0 Then                                 ' Si la ejecucion dejo un error registrado...
    strError = "Error " & Err.Number & ": " & Err.Description ' ...se guarda su detalle para informarlo
    Err.Clear                                           ' Se limpia el error para que no afecte el cierre de recursos
End If                                                  ' Fin de la revision del error

' --- CIERRE GARANTIZADO (equivale a "finally") ---
If Not objLibro Is Nothing Then objLibro.Close False    ' Cierra el libro sin preguntar (ya fue guardado en Ejecutar)
If Not objExcel Is Nothing Then objExcel.Quit           ' Cierra Excel para no dejar EXCEL.EXE huerfano
Set objLibro = Nothing                                  ' Libera la referencia al libro
Set objExcel = Nothing                                  ' Libera la referencia a Excel
Set objFso = Nothing                                    ' Libera la referencia al sistema de archivos
On Error GoTo 0                                         ' Restaura el manejo normal de errores

If strError <> "" Then                                  ' Si hubo error durante el proceso...
    WScript.Echo "FALLO: " & strError                   ' ...lo muestra al usuario
    WScript.Quit 1                                      ' Termina con codigo 1 (convencion: distinto de 0 = error)
Else                                                    ' Si todo salio bien...
    WScript.Echo "OK: archivo Excel creado/actualizado." ' ...confirma el resultado
    WScript.Quit 0                                      ' Termina con codigo 0 (exito)
End If                                                  ' Fin del mensaje final

' =====================================================================
' Sub Ejecutar : logica principal (separada para el patron try/finally)
' =====================================================================
Sub Ejecutar()                                          ' Inicio del procedimiento principal
    Dim strCarpetaCodigo                                ' Carpeta donde vive este script (C:\pruebas\code)
    Dim strCarpetaExcel                                 ' Carpeta destino del Excel (C:\excel)
    Dim strRutaExcel                                    ' Ruta completa del archivo .xlsx
    Dim arrPaises                                       ' Lista de paises
    Dim arrPoblacion                                    ' Lista de poblaciones (misma posicion que el pais)
    Dim objHoja                                         ' Hoja activa donde se escribe
    Dim i                                               ' Contador del ciclo de escritura

    strCarpetaCodigo = objFso.GetParentFolderName(WScript.ScriptFullName) ' Obtiene la carpeta del script (raiz del codigo)
    strCarpetaExcel = objFso.GetAbsolutePathName(strCarpetaCodigo & "\..\..\excel") ' Sube 2 niveles (C:\pruebas\code -> C:\) y entra a \excel
    strRutaExcel = strCarpetaExcel & "\poblacion_latam.xlsx" ' Arma la ruta completa del archivo de salida

    If Not objFso.FolderExists(strCarpetaExcel) Then    ' Si la carpeta destino no existe...
        objFso.CreateFolder strCarpetaExcel             ' ...la crea (accion no destructiva) para poder guardar
    End If                                              ' Fin de la verificacion de carpeta

    arrPaises = ObtenerPaises()                         ' Carga la lista de paises desde la funcion de datos
    arrPoblacion = ObtenerPoblacion()                   ' Carga la lista de poblaciones desde la funcion de datos

    If UBound(arrPaises) <> UBound(arrPoblacion) Then   ' Valida que ambas listas tengan igual tamano
        Err.Raise vbObjectError + 1, "depurar", "Las listas de paises y poblacion no coinciden." ' Aborta con error propio si no coinciden
    End If                                              ' Fin de la validacion

    Set objExcel = CreateObject("Excel.Application")    ' Inicia una instancia de Excel por COM
    objExcel.Visible = False                            ' Oculta Excel: es un proceso por lotes, no interactivo
    objExcel.DisplayAlerts = False                      ' Evita cuadros de dialogo que bloquearian el script

    If objFso.FileExists(strRutaExcel) Then             ' Si el archivo ya existe...
        Set objLibro = objExcel.Workbooks.Open(strRutaExcel) ' ...lo abre para ACTUALIZARLO (no se borra ni recrea)
        If objLibro.ReadOnly Then                       ' Si se abrio solo lectura (otro usuario/proceso lo tiene abierto)...
            Err.Raise vbObjectError + 2, "depurar", "El archivo esta en uso o es de solo lectura: " & strRutaExcel ' ...detiene para no perder cambios
        End If                                          ' Fin de la verificacion de solo lectura
    Else                                                ' Si el archivo no existe...
        Set objLibro = objExcel.Workbooks.Add()         ' ...crea un libro nuevo en blanco
    End If                                              ' Fin de abrir o crear

    Set objHoja = objLibro.Worksheets(1)                ' Toma la primera hoja del libro como hoja de trabajo
    objHoja.Name = "Poblacion"                          ' Nombra la hoja (idempotente: no cambia si ya se llama asi)

    objHoja.Cells(1, 1).Value = "Pais"                  ' Encabezado de la columna A
    objHoja.Cells(1, 2).Value = "Poblacion"             ' Encabezado de la columna B

    For i = 0 To UBound(arrPaises)                      ' Recorre cada pais de la lista (base 0)
        objHoja.Cells(i + 2, 1).Value = arrPaises(i)    ' Escribe el pais en columna A (fila +2: se salta el encabezado y la base 0)
        objHoja.Cells(i + 2, 2).Value = arrPoblacion(i) ' Escribe la poblacion en columna B de la misma fila
    Next                                                ' Siguiente pais

    objHoja.Range("A1:B1").Font.Bold = True             ' Pone los encabezados en negrita para distinguirlos
    objHoja.Range("B2:B" & (UBound(arrPaises) + 2)).NumberFormat = "#,##0" ' Formato numerico con separador de miles (el dato sigue siendo numero)
    objHoja.Columns("A:B").AutoFit                      ' Ajusta el ancho de columnas al contenido

    If objFso.FileExists(strRutaExcel) Then             ' Si el archivo ya existia...
        objLibro.Save                                   ' ...guarda sobre el mismo archivo (actualizacion)
    Else                                                ' Si era nuevo...
        objLibro.SaveAs strRutaExcel, xlOpenXMLWorkbook ' ...lo guarda por primera vez en formato .xlsx
    End If                                              ' Fin del guardado
End Sub                                                 ' Fin de Ejecutar

' =====================================================================
' Function ObtenerPaises : devuelve el arreglo de paises
' Las tildes se construyen con ChrW (233=e', 237=i', 225=a', 250=u')
' =====================================================================
Function ObtenerPaises()                                ' Inicio de la funcion de paises
    ObtenerPaises = Array( _
        "Argentina", _
        "Bolivia", _
        "Brasil", _
        "Chile", _
        "Colombia", _
        "Costa Rica", _
        "Cuba", _
        "Ecuador", _
        "El Salvador", _
        "Guatemala", _
        "Hait" & ChrW(237), _
        "Honduras", _
        "M" & ChrW(233) & "xico", _
        "Nicaragua", _
        "Panam" & ChrW(225), _
        "Paraguay", _
        "Per" & ChrW(250), _
        "Rep" & ChrW(250) & "blica Dominicana", _
        "Uruguay", _
        "Venezuela")                                    ' Devuelve 20 paises (indices 0 a 19)
End Function                                            ' Fin de ObtenerPaises

' =====================================================================
' Function ObtenerPoblacion : devuelve el arreglo de habitantes
' Mismo orden que ObtenerPaises. Valores aproximados (ONU WPP 2024).
' =====================================================================
Function ObtenerPoblacion()                             ' Inicio de la funcion de poblacion
    ObtenerPoblacion = Array( _
        45696000, _
        12414000, _
        211999000, _
        19764000, _
        52886000, _
        5213000, _
        11194000, _
        18001000, _
        6365000, _
        18406000, _
        11725000, _
        10656000, _
        130861000, _
        6917000, _
        4515000, _
        6845000, _
        34217000, _
        11332000, _
        3423000, _
        28405000)                                       ' Devuelve 20 valores (indices 0 a 19)
End Function                                            ' Fin de ObtenerPoblacion

' =====================================================================
' PRUEBA DE ESCRITORIO
' =====================================================================
' Supuesto: script en C:\pruebas\code\depurar.vbs, Excel instalado,
'           C:\excel\poblacion_latam.xlsx NO existe (primera ejecucion).
'
' Paso | Instruccion / flujo                     | Variable -> valor
' -----|-----------------------------------------|---------------------------------
'  1   | Set objFso = CreateObject(...)          | objFso = FileSystemObject
'  2   | Call Ejecutar()                         | entra al procedimiento
'  3   | GetParentFolderName(ScriptFullName)     | strCarpetaCodigo = C:\pruebas\code
'  4   | GetAbsolutePathName(...\..\..\excel)    | strCarpetaExcel  = C:\excel
'  5   | strRutaExcel = ...\poblacion_latam.xlsx | C:\excel\poblacion_latam.xlsx
'  6   | FolderExists(C:\excel)                  | True -> no crea carpeta
'  7   | ObtenerPaises / ObtenerPoblacion        | UBound = 19 en ambos -> valida OK
'  8   | CreateObject("Excel.Application")       | objExcel = Excel oculto
'  9   | FileExists(strRutaExcel)                | False -> Workbooks.Add()
' 10   | Worksheets(1).Name = "Poblacion"        | objHoja = hoja "Poblacion"
' 11   | Cells(1,1)/(1,2)                        | A1="Pais", B1="Poblacion"
' 12   | For i = 0                               | A2="Argentina", B2=45696000
' 13   | i = 1                                   | A3="Bolivia",   B3=12414000
' 14   | i = 2 ... i = 19                        | ... A21="Venezuela", B21=28405000
' 15   | Formato                                 | B2:B21 = "#,##0"; A1:B1 negrita
' 16   | FileExists -> False -> SaveAs (51)      | se crea el .xlsx
' 17   | Retorno a nivel principal               | Err.Number = 0 -> strError = ""
' 18   | Cierre                                  | libro cerrado, Excel.Quit, objetos liberados
' 19   | Mensaje final                           | "OK: archivo Excel creado/actualizado." (exit 0)
'
' Segunda ejecucion: en el paso 9 FileExists = True -> Workbooks.Open;
'   se sobrescriben A1:B21 con los mismos valores; en el paso 16 se
'   ejecuta Save (no SaveAs). Resultado: archivo actualizado, no duplicado.
'
' Caso de falla: archivo abierto por otro usuario -> ReadOnly = True ->
'   Err.Raise 2 -> sube a nivel principal -> strError = "Error ...: El
'   archivo esta en uso..." -> se cierra Excel -> "FALLO: ..." (exit 1).
'
' Resultado esperado: C:\excel\poblacion_latam.xlsx con 21 filas
'   (1 encabezado + 20 paises), poblacion numerica con miles formateados.
' =====================================================================
