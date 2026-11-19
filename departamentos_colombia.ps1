# =====================================================================
# ARCHIVO      : departamentos_colombia.ps1
# UBICACION    : C:\informes\egp\code\departamentos_colombia.ps1
# DESTINO      : C:\informes\egp\output\output.xlsx  (resuelto en relativo:
#                <carpeta del script>\..\output\output.xlsx)
# PROPOSITO    : Crear un libro Excel con 3 columnas:
#                  A = Departamento, B = Capital, C = Poblacion.
#                Si output.xlsx existe lo ACTUALIZA; si no existe lo CREA
#                (y crea la carpeta \output si falta).
# RESPONSABILIDAD: Resolver rutas relativas al script, armar el paquete
#                Open XML (.xlsx = ZIP de XML), publicarlo de forma
#                segura (temporal + reemplazo) y reportar el resultado.
# ALCANCE      : Fase 1. Datos embebidos: 32 departamentos + Bogota D.C.
#                Poblacion: proyeccion DANE 2025 (base CNPV 2018). Suma de
#                los 33 registros = 53.057.212 = total nacional (verificado).
#                DANE publico una serie actualizada el 08-ago-2025: confirmar
#                si el informe exige esa version.
# REQUISITOS   : Windows PowerShell 5.1 (nativo en Windows 10/11) o
#                PowerShell 7+. NO requiere Excel, COM, ni modulos externos.
# EJECUCION    : powershell -NoProfile -ExecutionPolicy Bypass -File .\departamentos_colombia.ps1
#                (Bypass aplica solo a ese proceso; no requiere admin).
# CODIFICACION : Guardar en UTF-8 CON BOM para que PowerShell 5.1 lea bien
#                las tildes (sin BOM las interpreta como ANSI -> "BogotÃ¡").
# =====================================================================
#
# ---------------------------------------------------------------------
# JUSTIFICACION TECNICA
# ---------------------------------------------------------------------
# - Open XML directo en vez de COM (New-Object -ComObject Excel.Application):
#     * COM deja procesos EXCEL.EXE huerfanos si el script falla, depende de
#       la version/licencia de Office y bloquea si hay dialogos ocultos.
#     * .NET (System.IO.Compression) viene en Windows: sin instalar nada,
#       compatible con equipos corporativos con permisos restringidos.
# - Rutas relativas con $PSScriptRoot (no con el directorio actual):
#     * El directorio actual cambia segun desde donde se invoque (CMD,
#       tarea programada, doble clic). $PSScriptRoot siempre es la carpeta
#       del .ps1, asi "..\output" siempre apunta al mismo lugar.
# - Escritura atomica (temporal + Move-Item):
#     * Si falla a mitad, el output.xlsx anterior queda intacto.
# - Compatible con PowerShell 5.1:
#     * Join-Path con 2 argumentos (el 3er+ segmento solo existe en PS 7).
#     * Add-Type de System.IO.Compression (en PS 5.1 no se carga solo).
#
# ---------------------------------------------------------------------
# LIMITACION CONOCIDA
# ---------------------------------------------------------------------
# "Actualizar" = regenerar el libro completo. Cambios manuales hechos en
# output.xlsx (otras hojas, formulas) se pierden en cada ejecucion.
#
# ---------------------------------------------------------------------
# FUTURAS FUNCIONALIDADES / ESTADO DEL ECOSISTEMA
# ---------------------------------------------------------------------
# - DANE publica datos abiertos (datos.gov.co, API Socrata): una fase
#   siguiente podria descargar la poblacion con Invoke-RestMethod si la
#   red corporativa lo permite, en vez de datos embebidos.
# - Windows PowerShell 5.1 no recibe nuevas funciones (solo seguridad);
#   PowerShell 7.x es la linea activa, pero requiere instalacion.
# ---------------------------------------------------------------------

#Requires -Version 5.1
# Aborta si el interprete es anterior a 5.1 (necesario para ::new() y [ordered])

Set-StrictMode -Version Latest
# Modo estricto: variables no declaradas o propiedades inexistentes generan error

$ErrorActionPreference = 'Stop'
# Todo error no terminante pasa a terminante para que lo capture catch

Add-Type -AssemblyName System.IO.Compression
# Carga ZipArchive (en PS 5.1 no viene cargado por defecto)

# --- DATOS: departamento -> @(capital, poblacion DANE 2025) ---------
$datos = [ordered]@{
    'Amazonas'                  = @('Leticia',        85057)
    'Antioquia'                 = @('Medellín',       6928372)
    'Arauca'                    = @('Arauca',         279191)
    'Atlántico'                 = @('Barranquilla',   2865034)
    'Bogotá D.C.'               = @('Bogotá D.C.',    7942867)
    'Bolívar'                   = @('Cartagena',      2241282)
    'Boyacá'                    = @('Tunja',          1290393)
    'Caldas'                    = @('Manizales',      1054450)
    'Caquetá'                   = @('Florencia',      429041)
    'Casanare'                  = @('Yopal',          473165)
    'Cauca'                     = @('Popayán',        1599148)
    'Cesar'                     = @('Valledupar',     1469159)
    'Chocó'                     = @('Quibdó',         593106)
    'Córdoba'                   = @('Montería',       1992907)
    'Cundinamarca'              = @('Bogotá D.C.',    3535067)
    'Guainía'                   = @('Inírida',        59706)
    'Guaviare'                  = @('San José del Guaviare', 84696)
    'Huila'                     = @('Neiva',          1208728)
    'La Guajira'                = @('Riohacha',       1066679)
    'Magdalena'                 = @('Santa Marta',    1544507)
    'Meta'                      = @('Villavicencio',  1156405)
    'Nariño'                    = @('Pasto',          1713586)
    'Norte de Santander'        = @('Cúcuta',         1709289)
    'Putumayo'                  = @('Mocoa',          390742)
    'Quindío'                   = @('Armenia',        557884)
    'Risaralda'                 = @('Pereira',        1003225)
    'San Andrés y Providencia'  = @('San Andrés',     63438)
    'Santander'                 = @('Bucaramanga',    2398303)
    'Sucre'                     = @('Sincelejo',      1034102)
    'Tolima'                    = @('Ibagué',         1386410)
    'Valle del Cauca'           = @('Cali',           4708393)
    'Vaupés'                    = @('Mitú',           44142)
    'Vichada'                   = @('Puerto Carreño', 148738)
}
# Diccionario ordenado: una sola estructura evita listas paralelas desalineadas

function Add-EntradaZip {
    # Escribe un archivo de texto (parte XML) dentro del ZIP
    param(
        [System.IO.Compression.ZipArchive] $Zip,
        # Paquete ZIP abierto en modo Create

        [string] $Nombre,
        # Ruta interna de la parte (siempre con "/")

        [string] $Contenido
        # XML a guardar en esa parte
    )
    $entrada = $Zip.CreateEntry($Nombre)
    # Crea la entrada vacia dentro del ZIP

    $escritor = New-Object System.IO.StreamWriter($entrada.Open(), (New-Object System.Text.UTF8Encoding($false)))
    # Escritor UTF-8 SIN BOM sobre el flujo de la entrada (partes Open XML sin BOM)

    try {
        # Bloque protegido para cerrar siempre el escritor

        $escritor.Write($Contenido)
        # Escribe el XML completo
    }
    finally {
        # Se ejecuta aun si Write falla

        $escritor.Dispose()
        # Vacia buffer y cierra escritor + flujo de la entrada
    }
}
# Fin de Add-EntradaZip

function ConvertTo-CeldaTexto {
    # Devuelve el XML de una celda de texto inline, con el texto escapado
    param(
        [string] $Ref,
        # Referencia de celda (ej. "A2")

        [string] $Texto,
        # Texto a mostrar

        [int] $Estilo = 0
        # Indice de estilo (0 normal, 1 encabezado)
    )
    $seguro = [System.Security.SecurityElement]::Escape($Texto)
    # Escapa &, <, >, comillas: evita XML invalido (ej. "San Andrés & ...")

    return "<c r=`"$Ref`" s=`"$Estilo`" t=`"inlineStr`"><is><t>$seguro</t></is></c>"
    # Celda tipo inlineStr: no requiere sharedStrings.xml
}
# Fin de ConvertTo-CeldaTexto

$codigoSalida = 0
# 0 = exito, 1 = error (util para CMD / tareas programadas)

$rutaTemporal = $null
# Null hasta definirse: en catch indica si hay temporal que retirar

try {
    # Bloque protegido: cualquier error salta a catch

    $carpetaSalida = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\output'))
    # Sube 1 nivel desde \code y entra a \output; GetFullPath normaliza ".." -> C:\informes\egp\output

    $rutaSalida = Join-Path $carpetaSalida 'output.xlsx'
    # Ruta final del archivo Excel

    $rutaTemporal = Join-Path $carpetaSalida 'output.xlsx.tmp'
    # Temporal donde se arma el paquete antes de reemplazar el final

    if (-not (Test-Path -LiteralPath $carpetaSalida -PathType Container)) {
        # Si \output no existe...

        $null = New-Item -ItemType Directory -Path $carpetaSalida
        # ...la crea (no destructivo); $null suprime la salida del objeto
    }
    # Fin verificacion de carpeta

    $existia = Test-Path -LiteralPath $rutaSalida -PathType Leaf
    # Solo para el mensaje final: creado vs actualizado

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    # Numeros en XML sin separadores regionales (Colombia usa "." de miles)

    $xmlEnc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    # Declaracion XML comun a todas las partes

    $ultimaFila = $datos.Count + 1
    # N registros + 1 fila de encabezado

    $filas = New-Object System.Text.StringBuilder
    # Acumulador eficiente (concatenar strings en bucle es O(n^2))

    $null = $filas.Append('<row r="1">')
    # Abre fila de encabezados

    $null = $filas.Append((ConvertTo-CeldaTexto 'A1' 'Departamento' 1))
    # A1 en negrita

    $null = $filas.Append((ConvertTo-CeldaTexto 'B1' 'Capital' 1))
    # B1 en negrita

    $null = $filas.Append((ConvertTo-CeldaTexto 'C1' 'Población' 1))
    # C1 en negrita

    $null = $filas.Append('</row>')
    # Cierra fila 1

    $fila = 2
    # Primera fila de datos

    foreach ($par in $datos.GetEnumerator()) {
        # Recorre cada departamento en el orden definido

        $capital = [string]$par.Value[0]
        # Capital: primer elemento del arreglo

        $poblacion = ([long]$par.Value[1]).ToString($inv)
        # Poblacion como texto numerico invariante (ej. 4620000)

        $null = $filas.Append("<row r=`"$fila`">")
        # Abre la fila actual

        $null = $filas.Append((ConvertTo-CeldaTexto "A$fila" ([string]$par.Key)))
        # Columna A: departamento

        $null = $filas.Append((ConvertTo-CeldaTexto "B$fila" $capital))
        # Columna B: capital

        $null = $filas.Append("<c r=`"C$fila`" s=`"2`"><v>$poblacion</v></c>")
        # Columna C: NUMERO real (permite sumar/filtrar) con formato #,##0

        $null = $filas.Append('</row>')
        # Cierra la fila

        $fila++
        # Siguiente fila
    }
    # Fin del ciclo

    $hojaXml = $xmlEnc +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        "<dimension ref=`"A1:C$ultimaFila`"/>" +
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        '<cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/></cols>' +
        '<sheetData>' + $filas.ToString() + '</sheetData>' +
        "<autoFilter ref=`"A1:C$ultimaFila`"/>" +
        '</worksheet>'
    # Hoja: rango, encabezado inmovilizado, anchos A/B/C, filas y autofiltro

    $tiposXml = $xmlEnc +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>'
    # [Content_Types].xml: tipo de cada parte del paquete

    $relsRaizXml = $xmlEnc +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
    # _rels/.rels: documento principal = workbook

    $libroXml = $xmlEnc +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Departamentos" sheetId="1" r:id="rId1"/></sheets>' +
        '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Departamentos!$A$1:$C$' + $ultimaFila + '</definedName></definedNames>' +
        '</workbook>'
    # xl/workbook.xml: hoja "Departamentos" + nombre oculto que Excel exige para el autofiltro

    $relsLibroXml = $xmlEnc +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'
    # xl/_rels/workbook.xml.rels: enlaza libro con hoja y estilos

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
    # xl/styles.xml: 0 normal, 1 negrita, 2 formato integrado #,##0 (id 3)

    $flujoArchivo = [System.IO.File]::Open($rutaTemporal, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    # Crea/reemplaza solo el .tmp propio del script

    try {
        # Garantiza cerrar el archivo temporal

        $zip = New-Object System.IO.Compression.ZipArchive($flujoArchivo, [System.IO.Compression.ZipArchiveMode]::Create)
        # ZIP en modo escritura sobre el temporal

        try {
            # Garantiza cerrar el ZIP (escribe el directorio central)

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
            # Cierra el ZIP; sin esto el archivo queda corrupto
        }
    }
    finally {
        # Siempre se ejecuta

        $flujoArchivo.Dispose()
        # Libera el temporal para poder moverlo
    }

    Move-Item -LiteralPath $rutaTemporal -Destination $rutaSalida -Force
    # Publica el temporal como output.xlsx; falla si Excel lo tiene abierto

    $accion = if ($existia) { 'actualizado' } else { 'creado' }
    # Texto segun estado previo

    Write-Host "OK: archivo $accion -> $rutaSalida ($($datos.Count) registros)"
    # Confirma ruta final y cantidad de registros
}
catch {
    # Captura cualquier error del bloque try

    Write-Host "FALLO: $($_.Exception.Message)" -ForegroundColor Red
    # Muestra el error en rojo

    if ($rutaTemporal -and (Test-Path -LiteralPath $rutaTemporal)) {
        # Si quedo un temporal a medias (propio del script)...

        Remove-Item -LiteralPath $rutaTemporal -ErrorAction SilentlyContinue
        # ...lo retira; output.xlsx nunca se toca aqui
    }
    # Fin limpieza del temporal

    $codigoSalida = 1
    # Marca error
}

exit $codigoSalida
# 0 exito / 1 error, legible desde CMD con %ERRORLEVEL%

# =====================================================================
# PRUEBA DE ESCRITORIO
# =====================================================================
# Supuesto: script en C:\informes\egp\code\; C:\informes\egp\output\ NO existe.
#
# Paso | Instruccion                              | Variable -> valor
# -----|------------------------------------------|--------------------------------------
#  1   | #Requires 5.1 / StrictMode / Add-Type    | entorno listo
#  2   | $datos = [ordered]@{...}                 | 33 registros (32 dptos + Bogota D.C.)
#  3   | $codigoSalida = 0                        | 0
#  4   | GetFullPath(Join-Path $PSScriptRoot      | $carpetaSalida = C:\informes\egp\output
#      |   '..\output')                           |
#  5   | $rutaSalida / $rutaTemporal              | ...\output\output.xlsx / ...\output.xlsx.tmp
#  6   | Test-Path carpeta -> False               | New-Item crea C:\informes\egp\output
#  7   | $existia                                 | False
#  8   | $ultimaFila = 33 + 1                     | 34
#  9   | Fila 1                                   | A1=Departamento B1=Capital C1=Población
# 10   | foreach, $fila=2                         | A2=Amazonas  B2=Leticia  C2=85057
# 11   | $fila=3                                  | A3=Antioquia B3=Medellín C3=6928372
# 12   | ... ultimo ($fila=34)                    | A34=Vichada B34=Puerto Carreño C34=148738; $fila=35
# 13   | Arma 6 partes XML                        | hoja (A1:C34), tipos, rels, libro, rels, estilos
# 14   | File.Open + ZipArchive + Dispose         | output.xlsx.tmp completo
# 15   | Move-Item -Force                         | C:\informes\egp\output\output.xlsx
# 16   | $accion                                  | "OK: archivo creado -> ...\output.xlsx (33 registros)"
# 17   | exit 0                                   | %ERRORLEVEL% = 0
#
# 2da ejecucion: paso 6 no crea carpeta; $existia=True; Move-Item reemplaza;
#   mensaje "actualizado".
# Falla (output.xlsx abierto en Excel): Move-Item lanza error -> catch,
#   "FALLO: ..." en rojo, se retira el .tmp, exit 1, output.xlsx intacto.
#
# Resultado: hoja "Departamentos", 34 filas (1 encabezado + 33), encabezado
#   en negrita e inmovilizado, autofiltro, poblacion numerica con #,##0.
# =====================================================================
