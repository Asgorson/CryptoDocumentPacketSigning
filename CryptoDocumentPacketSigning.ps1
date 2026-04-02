$pathToUnsigned = "~\Desktop\AutoSigning\Unsigned\"
$pathToSigned = "~\Desktop\AutoSigning\Signed\"
$thumbprint = "EC592c195771aaf224d3044d156828430bf53eab" # отпечаток Виги

# переименование неподписанного документа
function changeUnsignedFileName($pathToUnsigned){
    $i = -1
    Get-ChildItem -Path $pathToUnsigned | foreach-object{
        $i++
        [String[]]$actualFileNameArray += $_.name                               # заполнение массива исходными названиями файлов 
        $extension = (Get-Item $_.Fullname).Extension                           # получить расширение файла
        Rename-Item -Path $_.FullName -newName ("Unsigned" + $i + $extension)   # переименование неподписанных документов. Необходимо, потому что крипто про не принимает кирилицу, если работать через скрипт
        [String[]]$changedFilenameArray += ("Unsigned" + $i + $extension)       # заполнение массива измененными названиями файлов
    }
    $data = @{
        actualFileNameArray = $actualFileNameArray
        changedFilenameArray = $changedFileNameArray
    }
    return $data
}

# возврат актуального имени документа тест
function returnActualName($pathToFolder, $changedFilenameArray, $actualFileNameArray, $key){
    if($key -eq "signed"){
        for($i=0; $i -lt $actualFileNameArray.length; $i++){
            $fullFilePath = $pathToFolder + $changedFilenameArray[$i] + ".sig"
            try{
                Rename-Item -path $fullFilePath -newName ($actualFileNameArray[$i] + ".sig") -ErrorAction Stop
            }catch{
                if($_.Exception.Message -like "*не существует*"){
                    Write-Host "Файл не найден"
                }else{
                    Write-Error "$_.ExceptionMessage"
                }
            }
        }
    }elseif($key -eq "unsigned"){
        for($i=0; $i -lt $actualFileNameArray.length; $i++){
            $fullFilePath = $pathToFolder + $changedFilenameArray[$i]
            try{
                Rename-Item -path $fullFilePath -newName ($actualFileNameArray[$i]) -ErrorAction Stop
            }catch{
                if($_.Exception.Message -like "*не существует*"){
                    Write-Host "Файл не найден"
                }else{
                    Write-Error "$_.Exception.Message"
                }
            }
        }
    }
    
}

function getThumbprint(){
    $bossName = (Read-Host "`nВведите фамилию руководителя, на которого оформлена подпись").Trim()
    $certArray = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where {$_.Subject -like "*$bossName*"}
    $currentDate = Get-Date
    foreach($cert in $certArray){
        if($currentDate -lt $cert.notAfter){
            Write-Host "`n-----------------------------------------------------------------------------"
            Write-Host "`nСертификат: `n$($cert.subject) `n`nОтпечаток: `n$($cert.thumbprint) `n`nИстекает: `n$($cert.notAfter)"
            Write-Host "`n-----------------------------------------------------------------------------"
        }
    }
    $thumbPrint = Read-Host ("`nСкопируйте нужный отпечаток сертификата и вствьте сюда")
    return $thumbPrint
}

# перед выполнением скрипта нужно выбрать английскую раскладку, иначе иммитация нажатия клавиш будет некорректной
$thumbprint = getThumbprint
Set-WinUserLanguageList -LanguageList en-US, ru-ru -force
$process = Start-Process -passThru -FilePath "C:\Program Files (x86)\Common Files\Crypto Pro\Shared\cptools.exe"   # запуск программы             # ожидание запуска программы
$wshell = New-Object -ComObject WScript.Shell
while(!$wshell.AppActivate($process.ID)){
    Start-Sleep -s 0.2              # ожидание запуска программы
}              
$wshell = New-Object -ComObject WScript.Shell
Start-Sleep -Milliseconds 500       # ожидание выполнения команды
$wshell.SendKeys("{TAB}")           # переход в поле вкладок
Start-Sleep -Milliseconds 200
$wshell.SendKeys("{DOWN 4}")        # переход во вкладку "Создание подиси"
Start-Sleep -Milliseconds 200
$wshell.SendKeys("{TAB 2}")         # переход в поле "Поиск сертификата"
Start-Sleep -Milliseconds 200    
$wshell.SendKeys("$thumbprint")     # ввод отпечатка сертификата
Start-Sleep -Milliseconds 200  
$wshell.SendKeys("{TAB}")           # переход в поле выбора пути для неподписанного файла
Start-Sleep -Milliseconds 200
$data = changeUnsignedFileName $pathToUnsigned
$pathToUnsigned = Resolve-Path $pathToUnsigned
$pathToSigned = Resolve-Path $pathToSigned
for ($j=0; $j -lt $data.changedFilenameArray.length; $j++){
    $wshell.SendKeys("$($pathToUnsigned.path + $data.changedFilenameArray[$j])") # ввод пути до документа для подписи
    Start-Sleep -Milliseconds 200
    $wshell.SendKeys("{TAB}")           # переход в поле выбора пути для подписанного файла
    Start-Sleep -Milliseconds 200
    $wshell.SendKeys("$($pathToSigned.path + $data.changedFilenameArray[$j]).sig")   # ввод пути для подписанного документа
    Start-Sleep -Milliseconds 200
    $wshell.SendKeys("{ENTER}")         # Подписание файла
    Start-Sleep -Milliseconds 2000
    $wshell.SendKeys("+{TAB}")          # переход обратно в поле ввода пути до документа для подписи
    Start-Sleep -Milliseconds 1000
}
Stop-process -name "cptools"
Start-Sleep -Milliseconds 1000
returnActualName $pathToUnsigned.path $data.changedFilenameArray $data.actualFileNameArray "unsigned"   # Возврат актуального имени неподписанного файла 
Start-Sleep -Milliseconds 1000
returnActualName $pathToSigned.path $data.changedFilenameArray $data.actualFileNameArray "signed"       # Возврат актуального имени подписанного файла
Set-WinUserLanguageList -LanguageList ru-ru, en-US -Force

#1 tab - Перейти в поле вкладок 
#4 down - Перейти во вкладку "Создание подиси"
#2 tab - Пеерйти в поле "Поиск сертификата"
#1 tab - Перейти в поле выбора пути для неподписанного файла
#1 tab - Перейти в поле выбора пути для подписанного файла
#1 enter - Подписать
#1 shift + tab - Вернуться в поле выбора пути неподписанного файла
