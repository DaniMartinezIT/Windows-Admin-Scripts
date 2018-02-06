#############################################################################################
########### Extrae eventos del sistema de un Windows XP y los carga en una excel   ##########
###########                                                                        ##########
########### Sólo para XP (Para Windows 7 migrar Get-eventLog a Get-WinEvent)       ##########
###########                                                                        ##########
#############################################################################################
########### Intrucciones:                                                          ##########                  
########### - Entrar en el equipo remoto a cualquier unidad que pida contraseña    ##########
########### Por ejemplo: \\ip_equipo\c$                                            ##########
########### - Ejecutar éste script                                                 ##########
########### - Sentarse a esperar que abra la excel                                 ##########
#############################################################################################
##################################################### Daniel Martínez Rodríguez    ##########
#############################################################################################
#############################################################################################

#Funcion formateador de excel
Function Sheet_Formater ($Sheet) {
    $Sheet.Range('A1').cells="Indice"
    $Sheet.Range('B1').cells = "ID Instancia"
    $Sheet.Range('C1').cells="Fuente" 
    $Sheet.Range('D1').cells="ReplacementStrings" 
    $Sheet.Range('E1').cells="Hora"
    $Sheet.Range('F1').cells="Mensaje"
    $Sheet.Range('B2').columnWidth = 13
    $Sheet.Range('C2').columnWidth = 30
    $Sheet.Range('D2').columnWidth = 40
    $Sheet.Range('E2').columnWidth = 17
    $Sheet.Range('F2').columnWidth = 120
        
    $Sheet.Range('A1').Font.Bold=$true
    $Sheet.Range('B1').Font.Bold=$true
    $Sheet.Range('C1').Font.Bold=$true
    $Sheet.Range('D1').Font.Bold=$true
    $Sheet.Range('E1').Font.Bold=$true
    $Sheet.Range('F1').Font.Bold=$true

   
}

#Creando excel
$Excel = new-object -comobject Excel.Application      
$Workbook = $Excel.workbooks.add()     
$Hoja1 = $Workbook.sheets | where {$_.name -eq 'Hoja1'} 
$Hoja2 = $Workbook.sheets | where {$_.name -eq "Hoja2"} 
$Hoja1.name = "Applications"
$Hoja2.name = "System"

#Formateando dos hojas, una para log de Applications y otra para system (se descarta Security, ver último comentario)
Sheet_Formater $Hoja1
Sheet_Formater $Hoja2


#Recogida de variables ####### pendiente añadir opción de pasar por parametros
$equipo = Read-Host "¿Dirección Ip?"
$dias = Read-Host "¿De cuantos dias?"

#Recogida de Logs ####### pendiente comprobar que equipo es accesible para evitar un intento que va a dar error
$intRow1 = 1
Get-EventLog -LogName Application -ComputerName $equipo -entryType Error -After ((Get-Date).Date.AddDays(-$dias)) |
    ForEach-Object {
        #Guardado de datos en hoja1    
        $intRow1++   
        $Hoja1.cells.Item($intRow1,1) = $_.Index
        $Hoja1.cells.Item($intRow1,2) = $_.InstanceID
        $Hoja1.cells.Item($intRow1,3) = $_.Source
        $Hoja1.cells.Item($intRow1,4) = $_.ReplacementStrings
        $Hoja1.cells.Item($intRow1,5) = $_.TimeGenerated
        $Hoja1.cells.Item($intRow1,6) = $_.Message
        Write-Host "Exportando Log de aplicaciones. Se ha encontrado $intRow1 errores"
    }
$intRow2 = 1
Get-EventLog -LogName System -ComputerName $equipo -entryType Error -After ((Get-Date).Date.AddDays(-$dias)) |
    ForEach-Object {
        #Guardado de datos en hoja2
        $intRow2++   
        $Hoja2.cells.Item($intRow1,1) = $_.Index
        $Hoja2.cells.Item($intRow1,2) = $_.InstanceID
        $Hoja2.cells.Item($intRow1,3) = $_.Source
        $Hoja2.cells.Item($intRow1,4) = $_.ReplacementStrings
        $Hoja2.cells.Item($intRow1,5) = $_.TimeGenerated
        $Hoja2.cells.Item($intRow1,6) = $_.Message
        Write-Host "Sistema: Analizando $intRow2 de $_.count"
    }

    

# Mostrando resultados
Write-Host "Finalizado"
$Excel.visible = $true 


# Se podría añadir un Logname nuevo de Security, que contendría intentos de sesión acertados o fallidos. 
# Se descarta por no estar muy convencido de ser útiles. Si lees éste código y te parece necesario es poner "-LogName Security"
