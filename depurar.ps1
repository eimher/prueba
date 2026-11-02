# =====================================================================
# ARCHIVO      : depurar.ps1
# UBICACION    : C:\pruebas\code\depurar.ps1   (raiz del codigo)
# DESTINO      : C:\excel\poblacion_latam.xlsx (archivo Excel generado)
# PROPOSITO    : Crear un libro .xlsx con los paises de Latinoamerica
#                (columna A) y su poblacion (columna B). Si el archivo ya
#                existe, lo ACTUALIZA (se reemplaza por la version nueva).
# RESPONSABILIDAD: Resolver rutas relativas, armar el paquete Open XML
#                (.xlsx = ZIP de XML), escribirlo de forma segura y
#                reportar el resultado.
# ALCANCE      : Fase 1 (solo paises + poblacion). Datos embebidos
#                (aprox. ONU WPP 2024, redondeados, no verificados).
# REQUISITOS   : PowerShell 7+. NO requiere Excel instalado, ni COM, ni
#                modulos de terceros: solo .NET (System.IO.Compression).
# EJECUCION    : pwsh -ExecutionPolicy Bypass -File .\depurar.ps1
#                (Bypass aplica solo a ese proceso; no requiere admin).
# =====================================================================
#
# ---------------------------------------------------------------------
# JUSTIFICACION TECNICA (por que este enfoque)
# ---------------------------------------------------------------------
# - Generar el .xlsx directo (Open XML) en lugar de automatizar Excel por COM:
#     * COM exige Excel instalado y con licencia, abre un proceso EXCEL.EXE
#       que puede quedar huerfano, no funciona en Linux/servidores y es
#       imposible de probar sin Windows + Office.
#     * Un .xlsx es un ZIP con XML. Con System.IO.Compression (incluido en
#       .NET) se escribe sin Excel: mas rapido, determinista y verificable
#       en cualquier sistema donde corra PowerShell 7.
# - Un .xlsx real en lugar de CSV:
#     * El CSV no guarda formato ni tipos y depende de la region
#       (separador ; o ,) y de la codificacion (tildes).
# - Escritura segura (archivo temporal + Move-Item):
#     * Escribir directo sobre el destino deja un archivo corrupto si el
#       proceso falla a la mitad. Con temporal, el destino solo cambia si
#       el .xlsx nuevo quedo completo. Si Excel tiene el archivo abierto,
#       Move-Item falla con mensaje claro y el original no se toca.
# - Cadenas "inline" (t="inlineStr") en lugar de sharedStrings.xml:
#     * Menos partes en el paquete y menos indices que mantener; Excel las
#       lee sin problema y las convierte al guardar.
# - Set-StrictMode + $ErrorActionPreference = 'Stop':
#     * Por defecto PowerShell continua tras errores no terminantes y
#       acepta variables inexistentes ($null en silencio). Con ambas
#       directivas cada error detiene el flujo y cae en catch.
# - Join-Path con segmentos separados y $PSScriptRoot:
#     * Las rutas absolutas "C:\..." se rompen al mover la carpeta; con
#       segmentos separados el separador correcto lo pone .NET por sistema.
#
# ---------------------------------------------------------------------
# LIMITACION CONOCIDA (actualizar = regenerar)
# ---------------------------------------------------------------------
# Al "actualizar" se regenera el libro completo con los datos del script.
# Si el usuario agrego otras hojas, formulas o celdas manualmente en el
# .xlsx, se pierden. Fase siguiente (si se requiere): editar solo
# xl/worksheets/sheet1.xml dentro del ZIP existente (ZipArchiveMode.Update).
#
# ---------------------------------------------------------------------
# FUTURAS FUNCIONALIDADES / ESTADO DEL ECOSISTEMA
# ---------------------------------------------------------------------
# - No conozco una API nativa de .NET para leer/escribir .xlsx de alto
#   nivel anunciada; hoy se resuelve con Open XML SDK (paquete de
#   Microsoft, pero de terceros para tus reglas) o a mano como aqui.
# - Office Scripts / Excel JavaScript API: automatizacion sin COM,
#   disponible (GA) en Excel web; sujeta a licencia y politica de TI.
# - PowerShell 7.x evoluciona anualmente con .NET; revisar notas de version.
# ---------------------------------------------------------------------

#Requires -Version 7.0
# Aborta de inmediato si el interprete es anterior a PowerShell 7.0

Set-StrictMode -Version Latest
# Activa modo estricto: variables no declaradas y propiedades inexistentes generan error

$ErrorActionPreference = 'Stop'
# Convierte todo error no terminante en terminante para que lo capture catch

# --- DATOS (diccionario ordenado: pais -> habitantes) -----------------
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
# Conserva el orden de insercion y evita dos listas paralelas que se desalineen

function Add-EntradaZip {
    # Funcion auxiliar: escribe un archivo de texto dentro del ZIP
    param(
        [System.IO.Compression.ZipArchive] $Zip,
        # Paquete ZIP abierto en modo escritura

        [string] $Nombre,
        # Ruta interna de la parte (siempre con "/")

        [string] $Contenido
        # Texto XML a guardar en esa parte
    )
    $entrada = $Zip.CreateEntry($Nombre)
    # Crea la entrada vacia dentro del ZIP

    $flujo = $entrada.Open()
    # Abre el flujo de escritura de la entrada

    try {
        # Bloque protegido para garantizar el cierre del flujo

        $escritor = [System.IO.StreamWriter]::new($flujo, [System.Text.UTF8Encoding]::new($false))
        # Escritor UTF-8 SIN BOM (Excel rechaza algunos XML con BOM en partes del paquete)

        try {
            # Bloque protegido para cerrar el escritor

            $escritor.Write($Contenido)
            # Escribe el XML completo en la entrada
        }
        finally {
            # Siempre se ejecuta

            $escritor.Dispose()
            # Vacia y cierra el escritor (cierra tambien el flujo)
        }
    }
    finally {
        # Siempre se ejecuta

        $flujo.Dispose()
        # Garantiza el cierre aunque el escritor no se haya creado
    }
}
# Fin de Add-EntradaZip

$codigoSalida = 0
# Codigo de salida del proceso: 0 = exito, 1 = error

$rutaTemporal = $null
# Ruta del temporal; null para saber en catch si llego a definirse

try {
    # Inicio del bloque protegido: cualquier error salta a catch

    $carpetaExcel = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..' '..' 'excel'))
    # Sube 2 niveles desde C:\pruebas\code hasta C:\ y entra a \excel; GetFullPath normaliza los ".."

    $rutaExcel = Join-Path $carpetaExcel 'poblacion_latam.xlsx'
    # Ruta completa del archivo de salida

    $rutaTemporal = Join-Path $carpetaExcel 'poblacion_latam.xlsx.tmp'
    # Ruta del temporal donde se arma el paquete antes de publicarlo

    if (-not (Test-Path -LiteralPath $carpetaExcel -PathType Container)) {
        # Si la carpeta destino no existe...

        $null = New-Item -ItemType Directory -Path $carpetaExcel
        # ...la crea (accion no destructiva); $null evita imprimir el objeto
    }
    # Fin de la verificacion de carpeta

    $existia = Test-Path -LiteralPath $rutaExcel -PathType Leaf
    # Guarda si el archivo ya existia (solo para informar el mensaje final)

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    # Cultura invariante: los numeros en XML siempre llevan punto decimal, sin importar la region

    $xmlEnc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    # Declaracion XML comun a todas las partes del paquete

    $ultimaFila = $datos.Count + 1
    # Ultima fila con datos: N paises + 1 de encabezado

    $filas = [System.Text.StringBuilder]::new()
    # Acumulador del XML de las filas (evita concatenar strings en bucle, que es O(n^2))

    $null = $filas.Append('<row r="1">')
    # Abre la fila 1 (encabezados)

    $null = $filas.Append('<c r="A1" s="1" t="inlineStr"><is><t>País</t></is></c>')
    # Celda A1: texto "País" con estilo 1 (negrita)

    $null = $filas.Append('<c r="B1" s="1" t="inlineStr"><is><t>Población</t></is></c>')
    # Celda B1: texto "Población" con estilo 1 (negrita)

    $null = $filas.Append('</row>')
    # Cierra la fila 1

    $fila = 2
    # Primera fila de datos

    foreach ($par in $datos.GetEnumerator()) {
        # Recorre cada par pais/poblacion en el orden definido

        $pais = [System.Security.SecurityElement]::Escape([string]$par.Key)
        # Escapa &, <, > y comillas para que el XML siempre sea valido

        $poblacion = ([long]$par.Value).ToString($inv)
        # Convierte la poblacion a texto numerico invariante (ej. 45696000)

        $null = $filas.Append("<row r=`"$fila`">")
        # Abre la fila actual

        $null = $filas.Append("<c r=`"A$fila`" t=`"inlineStr`"><is><t>$pais</t></is></c>")
        # Celda A: el pais como texto

        $null = $filas.Append("<c r=`"B$fila`" s=`"2`"><v>$poblacion</v></c>")
        # Celda B: la poblacion como NUMERO (sin t=) con estilo 2 (formato #,##0)

        $null = $filas.Append('</row>')
        # Cierra la fila actual

        $fila++
        # Avanza a la siguiente fila
    }
    # Fin del ciclo de filas

    $hojaXml = $xmlEnc +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        "<dimension ref=`"A1:B$ultimaFila`"/>" +
        '<cols><col min="1" max="1" width="26" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/></cols>' +
        '<sheetData>' + $filas.ToString() + '</sheetData></worksheet>'
    # Hoja completa: rango usado, anchos de columna (A=26, B=16) y las filas

    $tiposXml = $xmlEnc +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>'
    # [Content_Types].xml: declara el tipo de cada parte para que Excel sepa interpretarla

    $relsRaizXml = $xmlEnc +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
    # _rels/.rels: indica cual es el documento principal (el workbook)

    $libroXml = $xmlEnc +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Poblacion" sheetId="1" r:id="rId1"/></sheets></workbook>'
    # xl/workbook.xml: una sola hoja llamada "Poblacion"

    $relsLibroXml = $xmlEnc +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'
    # xl/_rels/workbook.xml.rels: enlaza el workbook con la hoja y los estilos

    $estilosXml = $xmlEnc +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="3">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
        '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    # xl/styles.xml: estilo 0 normal, 1 negrita (encabezados), 2 formato integrado #,##0 (id 3); cellStyles declara el estilo "Normal" que Excel espera

    $flujoArchivo = [System.IO.File]::Open($rutaTemporal, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    # Abre (o crea) el temporal; Create reemplaza solo un .tmp propio de ejecuciones previas

    try {
        # Bloque protegido para cerrar el archivo temporal siempre

        $zip = [System.IO.Compression.ZipArchive]::new($flujoArchivo, [System.IO.Compression.ZipArchiveMode]::Create)
        # Crea el ZIP en modo escritura sobre el temporal

        try {
            # Bloque protegido para cerrar el ZIP siempre (cierra y escribe el directorio central)

            Add-EntradaZip $zip '[Content_Types].xml' $tiposXml
            # Parte 1: tipos de contenido

            Add-EntradaZip $zip '_rels/.rels' $relsRaizXml
            # Parte 2: relacion raiz

            Add-EntradaZip $zip 'xl/workbook.xml' $libroXml
            # Parte 3: libro

            Add-EntradaZip $zip 'xl/_rels/workbook.xml.rels' $relsLibroXml
            # Parte 4: relaciones del libro

            Add-EntradaZip $zip 'xl/styles.xml' $estilosXml
            # Parte 5: estilos

            Add-EntradaZip $zip 'xl/worksheets/sheet1.xml' $hojaXml
            # Parte 6: hoja con los datos
        }
        finally {
            # Siempre se ejecuta

            $zip.Dispose()
            # Cierra el ZIP: sin esto el archivo queda incompleto
        }
    }
    finally {
        # Siempre se ejecuta

        $flujoArchivo.Dispose()
        # Libera el archivo temporal para poder moverlo
    }

    Move-Item -LiteralPath $rutaTemporal -Destination $rutaExcel -Force
    # Publica el temporal como archivo final (reemplaza el existente); falla si esta abierto en Excel

    $accion = if ($existia) { 'actualizado' } else { 'creado' }
    # Texto del mensaje segun si el archivo ya existia

    Write-Host "OK: archivo Excel $accion -> $rutaExcel ($($datos.Count) paises)"
    # Confirma el resultado con la ruta final y la cantidad de paises
}
catch {
    # Cualquier error del bloque try cae aqui

    Write-Host "FALLO: $($_.Exception.Message)" -ForegroundColor Red
    # Muestra el detalle del error en rojo

    if ($rutaTemporal -and (Test-Path -LiteralPath $rutaTemporal)) {
        # Si quedo un temporal a medias (archivo propio de este script)...

        Remove-Item -LiteralPath $rutaTemporal -ErrorAction SilentlyContinue
        # ...lo retira para no dejar basura; el archivo final nunca se toca aqui
    }
    # Fin de la limpieza del temporal

    $codigoSalida = 1
    # Marca el codigo de salida como error
}

exit $codigoSalida
# Termina con el codigo calculado (0 exito, 1 error) para uso en CMD/tareas programadas

# =====================================================================
# PRUEBA DE ESCRITORIO
# =====================================================================
# Supuesto: script en C:\pruebas\code\depurar.ps1; C:\excel\ no tiene aun
#           poblacion_latam.xlsx (primera ejecucion).
#
# Paso | Instruccion / flujo                    | Variable -> valor
# -----|----------------------------------------|-----------------------------------
#  1   | #Requires -Version 7.0                 | OK si es pwsh 7+
#  2   | $datos = [ordered]@{...}               | 20 pares pais -> habitantes
#  3   | $codigoSalida = 0                      | 0
#  4   | GetFullPath(Join-Path $PSScriptRoot    | $carpetaExcel = C:\excel
#      |   '..' '..' 'excel')                   |
#  5   | $rutaExcel / $rutaTemporal             | ...\poblacion_latam.xlsx / ...xlsx.tmp
#  6   | Test-Path carpeta                      | True -> no crea carpeta
#  7   | $existia = Test-Path archivo           | False
#  8   | $ultimaFila = Count + 1                | 21
#  9   | Fila 1 (encabezados)                   | A1="País", B1="Población" (estilo 1)
# 10   | foreach, $fila=2                       | A2=Argentina, B2=45696000
# 11   | $fila=3                                | A3=Bolivia,   B3=12414000
# 12   | ... ultimo par                         | A21=Venezuela, B21=28405000; $fila=22
# 13   | Arma 6 partes XML                      | hoja, tipos, rels, libro, rels libro, estilos
# 14   | File.Open(temporal) + ZipArchive       | temporal con 6 entradas
# 15   | Dispose zip y flujo                    | temporal completo y liberado
# 16   | Move-Item temporal -> final -Force     | C:\excel\poblacion_latam.xlsx
# 17   | $accion = 'creado'                     | mensaje "OK: archivo Excel creado -> ... (20 paises)"
# 18   | exit 0                                 | exito
#
# Segunda ejecucion: $existia = True; en el paso 16 Move-Item -Force
#   reemplaza el anterior; mensaje "actualizado". Sin duplicados.
#
# Caso de falla: archivo abierto en Excel -> Move-Item lanza error ->
#   catch muestra "FALLO: ..." en rojo, retira el .tmp, $codigoSalida=1,
#   exit 1. El .xlsx original queda intacto.
#
# Resultado esperado: C:\excel\poblacion_latam.xlsx con 21 filas
#   (1 encabezado + 20 paises), poblacion numerica con separador de miles.
# =====================================================================
