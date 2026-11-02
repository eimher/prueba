# =====================================================================
# ARCHIVO      : depurar.ps1
# UBICACION    : C:\pruebas\code\depurar.ps1   (raiz del codigo)
# DESTINO      : C:\excel\poblacion_latam.xlsx (archivo Excel generado)
# PROPOSITO    : Crear un libro de Excel con los paises de Latinoamerica
#                (columna A) y su poblacion (columna B). Si el archivo ya
#                existe, lo abre y lo ACTUALIZA sin borrarlo ni recrearlo.
# RESPONSABILIDAD: Resolver rutas relativas, abrir/crear el libro,
#                escribir datos, dar formato, guardar y liberar Excel.
# ALCANCE      : Fase 1 (solo paises + poblacion). Datos embebidos
#                (aprox. ONU WPP 2024, redondeados, no verificados).
# REQUISITOS   : PowerShell 7+ en Windows y Excel (Office 365) instalado.
#                Sin modulos ni librerias de terceros.
# EJECUCION    : pwsh -ExecutionPolicy Bypass -File .\depurar.ps1
#                (Bypass aplica solo a ese proceso; no requiere admin).
# =====================================================================
#
# ---------------------------------------------------------------------
# JUSTIFICACION TECNICA (por que este enfoque)
# ---------------------------------------------------------------------
# - PowerShell 7 en lugar de VBScript:
#     * VBScript esta en deprecacion por Microsoft y sera retirado de
#       Windows; PowerShell 7 recibe soporte activo.
#     * VBScript leia los archivos como ANSI (tildes corruptas, ChrW).
#       PowerShell 7 usa UTF-8 por defecto: los nombres llevan tildes
#       directas.
#     * Tiene try/catch/finally real; VBScript solo tenia
#       "On Error Resume Next", que oculta errores y es fragil.
# - $PSScriptRoot + Join-Path en lugar de rutas "C:\..." fijas:
#     * Las rutas absolutas se rompen al mover carpeta, cambiar de unidad
#       o de equipo. Con $PSScriptRoot la ruta se calcula desde el propio
#       script y Join-Path usa el separador correcto.
# - Set-StrictMode + $ErrorActionPreference = 'Stop':
#     * Por defecto PowerShell continua tras errores no terminantes y
#       acepta variables inexistentes (valen $null en silencio). Con
#       ambas directivas cada error detiene el flujo y cae en catch.
# - COM de Excel (New-Object -ComObject) en lugar de CSV:
#     * El CSV no es un .xlsx real, pierde formato y depende de la region
#       (separador ; o ,) y de la codificacion.
# - Liberar COM con ReleaseComObject en finally:
#     * Sin liberar, queda EXCEL.EXE huerfano en memoria que bloquea el
#       archivo. El recolector de basura no es determinista con COM.
# - Actualizar = abrir y sobrescribir celdas (nunca borrar el archivo):
#     * Cumple la regla de no eliminar archivos y conserva lo que el
#       usuario agregue fuera de A1:B21.
#
# ---------------------------------------------------------------------
# FUTURAS FUNCIONALIDADES / ESTADO DEL ECOSISTEMA
# ---------------------------------------------------------------------
# - Escribir .xlsx sin Excel instalado: hoy se logra con System.IO.Packaging
#   / System.IO.Compression (nativos, pero manual: se arma el XML Open XML)
#   o con el modulo ImportExcel (tercero; prohibido por tus reglas). No
#   conozco una API nativa de .NET anunciada para esto; verificar la hoja
#   de ruta de .NET.
# - Office Scripts / Excel JavaScript API: automatizacion sin COM, ya
#   disponibles (GA) en Excel web, sujetas a licencia y politica de TI.
# - PowerShell 7.x sigue evolucionando (ciclo anual alineado a .NET LTS);
#   revisar notas de version para cambios en soporte COM.
# ---------------------------------------------------------------------

#Requires -Version 7.0
# Aborta de inmediato si el interprete es anterior a PowerShell 7.0

Set-StrictMode -Version Latest
# Activa modo estricto: variables no declaradas y propiedades inexistentes generan error

$ErrorActionPreference = 'Stop'
# Convierte todo error no terminante en terminante para que lo capture catch

# --- DATOS (orden paralelo: mismo indice = mismo pais) ---------------
$datos = [ordered]@{
    'Argentina'            = 45696000
    'Bolivia'              = 12414000
    'Brasil'               = 211999000
    'Chile'                = 19764000
    'Colombia'             = 52886000
    'Costa Rica'           = 5213000
    'Cuba'                 = 11194000
    'Ecuador'              = 18001000
    'El Salvador'          = 6365000
    'Guatemala'            = 18406000
    'Haití'                = 11725000
    'Honduras'             = 10656000
    'México'               = 130861000
    'Nicaragua'            = 6917000
    'Panamá'               = 4515000
    'Paraguay'             = 6845000
    'Perú'                 = 34217000
    'República Dominicana' = 11332000
    'Uruguay'              = 3423000
    'Venezuela'            = 28405000
}
# Diccionario ordenado pais -> habitantes; [ordered] conserva el orden y evita listas paralelas desalineadas

$xlOpenXMLWorkbook = 51
# Codigo de formato .xlsx (con COM tardio no existen las constantes de Excel)

$excel = $null
# Referencia a la aplicacion Excel; null para saber en finally si llego a crearse

$libro = $null
# Referencia al libro; null por el mismo motivo

$hoja = $null
# Referencia a la hoja; null por el mismo motivo

$codigoSalida = 0
# Codigo de salida del proceso: 0 = exito, 1 = error

try {
    # Inicio del bloque protegido: cualquier error salta a catch

    $carpetaCodigo = $PSScriptRoot
    # Carpeta donde vive este script (C:\pruebas\code)

    $carpetaExcel = [System.IO.Path]::GetFullPath((Join-Path $carpetaCodigo '..\..\excel'))
    # Sube 2 niveles (C:\pruebas\code -> C:\) y entra a \excel; GetFullPath normaliza los ".."

    $rutaExcel = Join-Path $carpetaExcel 'poblacion_latam.xlsx'
    # Ruta completa del archivo de salida

    if (-not (Test-Path -LiteralPath $carpetaExcel -PathType Container)) {
        # Si la carpeta destino no existe...

        $null = New-Item -ItemType Directory -Path $carpetaExcel
        # ...la crea (accion no destructiva); $null evita imprimir el objeto creado
    }
    # Fin de la verificacion de carpeta

    $existe = Test-Path -LiteralPath $rutaExcel -PathType Leaf
    # Guarda si el archivo ya existe; se reutiliza al abrir y al guardar

    $excel = New-Object -ComObject Excel.Application
    # Inicia una instancia de Excel por COM

    $excel.Visible = $false
    # Oculta Excel: es un proceso por lotes, no interactivo

    $excel.DisplayAlerts = $false
    # Evita cuadros de dialogo que bloquearian el script

    if ($existe) {
        # Si el archivo ya existe...

        $libro = $excel.Workbooks.Open($rutaExcel)
        # ...lo abre para ACTUALIZARLO (no se borra ni recrea)

        if ($libro.ReadOnly) {
            # Si se abrio en solo lectura (otro usuario/proceso lo tiene abierto)...

            throw "El archivo está en uso o es de solo lectura: $rutaExcel"
            # ...detiene el flujo para no perder cambios; lo atrapa catch
        }
        # Fin de la verificacion de solo lectura
    }
    else {
        # Si el archivo no existe...

        $libro = $excel.Workbooks.Add()
        # ...crea un libro nuevo en blanco
    }
    # Fin de abrir o crear

    $hoja = $libro.Worksheets.Item(1)
    # Toma la primera hoja como hoja de trabajo

    $hoja.Name = 'Poblacion'
    # Nombra la hoja (idempotente: no cambia si ya se llama asi)

    $hoja.Cells.Item(1, 1).Value2 = 'País'
    # Encabezado de la columna A

    $hoja.Cells.Item(1, 2).Value2 = 'Población'
    # Encabezado de la columna B

    $fila = 2
    # Primera fila de datos (la 1 es el encabezado)

    foreach ($par in $datos.GetEnumerator()) {
        # Recorre cada par pais/poblacion en el orden definido

        $hoja.Cells.Item($fila, 1).Value2 = $par.Key
        # Escribe el pais en la columna A de la fila actual

        $hoja.Cells.Item($fila, 2).Value2 = [double]$par.Value
        # Escribe la poblacion en la columna B como numero (double es el tipo numerico que Excel maneja por COM)

        $fila++
        # Avanza a la siguiente fila
    }
    # Fin del ciclo de escritura

    $ultimaFila = $fila - 1
    # Ultima fila escrita (fila ya apunta a la siguiente libre)

    $hoja.Range('A1:B1').Font.Bold = $true
    # Pone los encabezados en negrita para distinguirlos

    $hoja.Range("B2:B$ultimaFila").NumberFormat = '#,##0'
    # Formato numerico con separador de miles; el dato sigue siendo numero

    $null = $hoja.Columns('A:B').AutoFit()
    # Ajusta el ancho de columnas al contenido; $null descarta el valor devuelto

    if ($existe) {
        # Si el archivo ya existia...

        $libro.Save()
        # ...guarda sobre el mismo archivo (actualizacion)
    }
    else {
        # Si era nuevo...

        $libro.SaveAs($rutaExcel, $xlOpenXMLWorkbook)
        # ...lo guarda por primera vez en formato .xlsx
    }
    # Fin del guardado

    Write-Host "OK: archivo Excel creado/actualizado -> $rutaExcel"
    # Confirma el resultado al usuario con la ruta final
}
catch {
    # Cualquier error del bloque try cae aqui

    Write-Host "FALLO: $($_.Exception.Message)" -ForegroundColor Red
    # Muestra el detalle del error en rojo

    $codigoSalida = 1
    # Marca el codigo de salida como error
}
finally {
    # Se ejecuta SIEMPRE, haya o no error (cierre garantizado)

    if ($null -ne $libro) {
        # Si el libro llego a abrirse...

        $libro.Close($false)
        # ...lo cierra sin preguntar (ya fue guardado en try)

        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($libro)
        # Libera la referencia COM del libro
    }
    # Fin del cierre del libro

    if ($null -ne $hoja) {
        # Si se obtuvo la hoja...

        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hoja)
        # ...libera su referencia COM
    }
    # Fin de la liberacion de la hoja

    if ($null -ne $excel) {
        # Si Excel llego a iniciarse...

        $excel.Quit()
        # ...lo cierra para no dejar EXCEL.EXE huerfano

        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        # Libera la referencia COM de la aplicacion
    }
    # Fin del cierre de Excel

    [System.GC]::Collect()
    # Fuerza recoleccion de basura para soltar cualquier RCW pendiente

    [System.GC]::WaitForPendingFinalizers()
    # Espera a que terminen los finalizadores para que Excel salga de memoria de verdad
}
# Fin de try/catch/finally

exit $codigoSalida
# Termina con el codigo calculado (0 exito, 1 error) para uso en CMD/tareas programadas

# =====================================================================
# PRUEBA DE ESCRITORIO
# =====================================================================
# Supuesto: script en C:\pruebas\code\depurar.ps1, Excel instalado,
#           C:\excel\poblacion_latam.xlsx NO existe (primera ejecucion).
#
# Paso | Instruccion / flujo                    | Variable -> valor
# -----|----------------------------------------|-----------------------------------
#  1   | #Requires -Version 7.0                 | OK si es pwsh 7+
#  2   | $datos = [ordered]@{...}               | 20 pares pais -> habitantes
#  3   | $excel/$libro/$hoja = $null            | nada abierto aun; $codigoSalida=0
#  4   | $carpetaCodigo = $PSScriptRoot         | C:\pruebas\code
#  5   | GetFullPath(Join-Path ...\..\..\excel) | $carpetaExcel = C:\excel
#  6   | $rutaExcel                             | C:\excel\poblacion_latam.xlsx
#  7   | Test-Path carpeta                      | True -> no crea carpeta
#  8   | $existe = Test-Path archivo            | False
#  9   | New-Object -ComObject Excel.Application| $excel = Excel oculto
# 10   | if ($existe) -> else                   | Workbooks.Add() -> $libro nuevo
# 11   | $hoja = Worksheets.Item(1); Name       | hoja "Poblacion"
# 12   | Cells(1,1), Cells(1,2)                 | A1="País", B1="Población"
# 13   | foreach, $fila=2                       | A2=Argentina, B2=45696000
# 14   | $fila=3                                | A3=Bolivia,   B3=12414000
# 15   | ... hasta el ultimo par                | A21=Venezuela, B21=28405000; $fila=22
# 16   | $ultimaFila = $fila - 1                | 21
# 17   | Formato                                | B2:B21="#,##0"; A1:B1 negrita; AutoFit
# 18   | $existe False -> SaveAs(ruta, 51)      | se crea el .xlsx
# 19   | Write-Host                             | "OK: archivo Excel creado/actualizado -> C:\excel\poblacion_latam.xlsx"
# 20   | finally                                | libro cerrado, Excel.Quit, COM liberado
# 21   | exit $codigoSalida                     | 0
#
# Segunda ejecucion: en el paso 8 $existe = True -> Workbooks.Open;
#   se sobrescribe A1:B21 con los mismos valores; en el paso 18 se
#   ejecuta Save() (no SaveAs). Resultado: archivo actualizado, no duplicado.
#
# Caso de falla: archivo abierto por otro usuario -> ReadOnly = True ->
#   throw -> catch muestra "FALLO: El archivo está en uso..." en rojo,
#   $codigoSalida = 1 -> finally cierra Excel -> exit 1.
#
# Resultado esperado: C:\excel\poblacion_latam.xlsx con 21 filas
#   (1 encabezado + 20 paises), poblacion numerica con miles formateados.
# =====================================================================
