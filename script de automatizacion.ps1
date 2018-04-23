clear-host

####----------PROPIEDADES QUE MEJORAN LA VISUALIZACION------#####

$ie = New-Object -ComObject InternetExplorer.Application

$ie.navigate2("https://eu1.mobileiron.com/index.html#!/")

$ie.visible = $true

$ie.Width = 1920
$ie.Height = 1080

$ie.Left = 0
$ie.top = 0

$ie.ToolBar = 0

#####---------CONVERSION DEL ARCHIVO DE CSV A XLSX----------#####



$csv = "C:\Users\ortiga\Desktop\table.csv" #archivo origen
$xlsx = "C:\Users\ortiga\Desktop\table2.xlsx" #archivo destino
$delimitador = "," #especificamos el delimitador

# creamos una nueva hoja vacia y la seleccionamos
$excel = New-Object -ComObject excel.application 
$documento = $excel.Workbooks.Add(1)
$hoja = $documento.worksheets.Item(1)




$conecta = ("TEXT;" + $csv)
$Conector = $hoja.QueryTables.add($conecta,$hoja.Range("A1"))
$query = $hoja.QueryTables.item($Conector.name)
$query.TextFileOtherDelimiter = $delimitador
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,1 * $hoja.Cells.Columns.Count
$query.AdjustColumnWidth = 1


$query.Refresh()
$query.Delete()

#si no existe creamos el  documento

if(!(test-path $xlsx)){

$documento.SaveAs($xlsx,51)

}




####------EMPEZAMOS A TRABAJAR CON EL EXCEL----------####



$excel = New-Object -ComObject excel.application

$documento = $excel.Workbooks.Open($xlsx)

$hoja = $excel.WorkSheets.item(1)

[INT]$contador = 1



#CONTAMOS LAS FILAS

[INT]$filas = $hoja.UsedRange.Rows.Count



do{

$contador++


$columna2 = "A" + $contador

$dispositivo = $hoja.Cells.Columns.Range($columna2).Text


$columna = "BE" + $contador


if($usuario -contains '*javiermiras*'){


$contador++


}

$usuario = $hoja.Cells.Columns.Range($columna).Text



##----------AUTOMATIZACION INTERNET EXPLORER------------------##




Start-Sleep -s 5


 
$element = $ie.Document.getElementsByTagName("a")[2].click()


$ie.Document.getElementsByTagName("button")[0].click()

$ie.Document.getElementsByTagName("ul")[6].getElementsByTagName("li")[1].getElementsByTagName("a")[0].click()

$ie.Document.getElementsByName("email")[0].value = $usuario
$ie.Document.getElementsByName("uid")[0].value = $usuario









}until($contador -eq  $filas)





Get-Process -Name Excel | Stop-Process