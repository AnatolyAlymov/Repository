###############################################################################
#
# Copyright (C) Anatoly Alymov, 2021
#
# Requres PowerShell v5
# Usage:
# This script must be run with administrator rights which can run on the required servers
#
###############################################################################

# Функция логирования
function Set-MsgLog {
    param (
        # Parameter help description
        [Parameter(Mandatory=$true)][string]$pathLog,
        [Parameter(Mandatory=$true)]$Msg
    )

    Write-Output "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss"): $Msg" *>> $pathLog
    #Write-Output "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss"): $Msg"

}

# Функция поиска сервера в массиве
function Set-FoundServer {
    param (
        [Parameter(Mandatory=$true)]$array1, # подаем первый массив
        [Parameter(Mandatory=$true)]$array2  # подаем второй массив
    )
    # Проходимся значениями из первого массива по второму если находим то возвращаем 1
    $result = 0
    foreach ($item in $array1){
             
        Write-Host "server name: "  $item
        
        if ($item -eq $array2) {
            
            Write-Host "Ooh, server found"
            $result = 1
            break

        }   
        
    }

    return $result

}

# Функция перезаписи xml файла
function Update-SwitchComboBox {
    param (
        [Parameter(Mandatory=$true)]$id,
        [Parameter(Mandatory=$true)]$xml,
        [Parameter(Mandatory=$true)]$NameReplace,
        [Parameter(Mandatory=$true)]$pathRepositoryReplace,
        [Parameter(Mandatory=$true)]$pathLocalReplace,
        [Parameter(Mandatory=$true)]$pathLogsReplace,
        [Parameter(Mandatory=$true)]$DescriptReplace,
        [Parameter(Mandatory=$true)]$xmlSave
        )
    # Заменяем элементы в xml    
    $xml.SelectSingleNode("//Task[@id='$id']").Name = $NameReplace
    $xml.SelectSingleNode("//Task[@id='$id']").PathRepository = $pathRepositoryReplace
    $xml.SelectSingleNode("//Task[@id='$id']").PathLocal = $pathLocalReplace
    $xml.SelectSingleNode("//Task[@id='$id']").PathLogs = $pathLogsReplace
    $xml.SelectSingleNode("//Task[@id='$id']").Descript = $DescriptReplace
    
    # Сохраняем изменения
    $xml.Save($xmlSave)

}

# Функция переключение созданных заданий
function Get-SwitchComboBox {
    param (
        [Parameter(Mandatory=$true)]$id,
        [Parameter(Mandatory=$true)]$xml
    )
    # Меняем значения
    $name = $xml.SelectSingleNode("//Task[@id='$id']").Name
    $pathRepository = $xml.SelectSingleNode("//Task[@id='$id']").PathRepository
    $pathLocal = $xml.SelectSingleNode("//Task[@id='$id']").PathLocal
    $pathLogs = $xml.SelectSingleNode("//Task[@id='$id']").PathLogs
    $description = $xml.SelectSingleNode("//Task[@id='$id']").Descript
    
    # Возвращам измененые значения
    return @{"xmlName"=$name;"xmlPathRepository"=$pathRepository;"xmlPathLocal"=$pathLocal;"xmlPathLogs"=$pathLogs;"xmlDescript"=$description}

}

# Функция создания новый заданий в документе xml
function New-createXmlTask {
    
    param (
        [Parameter(Mandatory=$true)]$id,
        [Parameter(Mandatory=$true)]$xml,
        [Parameter(Mandatory=$true)]$Name,
        [Parameter(Mandatory=$true)]$pathRepository,
        [Parameter(Mandatory=$true)]$pathLocal,
        [Parameter(Mandatory=$true)]$pathLogs,
        [Parameter(Mandatory=$true)]$Descript,
        [Parameter(Mandatory=$true)]$xmlSave
        )
    # Создаем новую ветку в paramSchedTask
    $newnode = $xml.settings.paramSchedTask.AppendChild($xml.CreateElement("Task"))
    $newnode.SetAttribute("id","$id")
    $newNAME = $newnode.AppendChild($xml.CreateElement("Name"))
    $newNAME.AppendChild($xml.CreateTextNode("$Name"))
    $newPathRepo = $newnode.AppendChild($xml.CreateElement("PathRepository"))
    $newPathRepo.AppendChild($xml.CreateTextNode("$pathRepository"))
    $newPathLocal = $newnode.AppendChild($xml.CreateElement("PathLocal"))
    $newPathLocal.AppendChild($xml.CreateTextNode("$pathLocal"))
    $newPathLogs = $newnode.AppendChild($xml.CreateElement("PathLogs"))
    $newPathLogs.AppendChild($xml.CreateTextNode("$pathLogs"))
    $newDescript = $newnode.AppendChild($xml.CreateElement("Descript"))
    $newDescript.AppendChild($xml.CreateTextNode("$Descript"))

    # Сохраняем изменения
    $xml.Save($xmlSave)

}

# Фукция удаления раздела задания из документа xml
function Remove-SwitchXmlTask {
    param (
        [Parameter(Mandatory=$true)]$id,
        [Parameter(Mandatory=$true)]$xml,
        [Parameter(Mandatory=$true)]$xmlSave
    )
    # Выбираем ветку и удаляем ее из xml файла
    $n = $xml.SelectSingleNode("//settings/paramSchedTask/Task[@id='$id']")
    $n.ParentNode.RemoveChild($n)
    
    # Сохраняем изменения
    $xml.Save($xmlSave)

}

# Функция заполнения treeViews
function Get-TreeViewTask {
        param ()
    # Загружаем имена из файла xml
    $DataMenu = $Global:xmlConfig.SelectNodes('//settings/servers/s').Name
    Write-Host $DataMenu

    # Заданий может быть много надо проходиться по списку и добавлять в каждую ноду
    foreach ($i in $DataMenu){         
        Write-Host "Server: "$i

        $t = $treeView_Menu.Nodes.Add($i)
        $DataTask = $Global:xmlConfig.SelectSingleNode("//settings/servers/s[@id='$i']").Task
        if ($null -ne $DataTask) {
            
            $t.Nodes.AddRange($DataTask)
            Write-Host "Needed add to TreeViews"$DataTask

        }   
        
    }

}

# Функция создания сервера в документе xml
function New-createXmlServer {
    
    param (
        [Parameter(Mandatory=$true)]$id,
        [Parameter(Mandatory=$true)]$xml,
        [Parameter(Mandatory=$true)]$name,
        [Parameter(Mandatory=$true)]$task,
        [Parameter(Mandatory=$true)]$xmlSave
        )
    # Создаем новую ветку в paramSchedTask
    $newnode = $xml.settings.servers.AppendChild($xml.CreateElement("s"))
    $newnode.SetAttribute("id","$id")
    $newNAME = $newnode.AppendChild($xml.CreateElement("name"))
    $newNAME.AppendChild($xml.CreateTextNode("$name"))
    $newPathRepo = $newnode.AppendChild($xml.CreateElement("task"))
    $newPathRepo.AppendChild($xml.CreateTextNode("$task"))
    
    # Сохраняем изменения
    $xml.Save($xmlSave)

}

# Фукция удаления раздела задания из документа xml
function Remove-ServerXml {
    param (
        [Parameter(Mandatory=$true)]$id,
        [Parameter(Mandatory=$true)]$xml,
        [Parameter(Mandatory=$true)]$xmlSave
    )
    # Выбираем ветку и удаляем ее из xml файла
    $n = $xml.SelectSingleNode("//settings/servers/s[@id='$id']")
    $n.ParentNode.RemoveChild($n)
    
    # Сохраняем изменения
    $xml.Save($xmlSave)

}

# Функция для запуска скрипта и инициализации необходимых файлов
function Start-Prog { 
    Param ([string]$Commandline)
    
    # Иннициализируем переменные и присваиваем начальные значения. Указываем располжение файлов данных
    $Global:LogPath = "$pwd\applog.log"
    $Global:XML = "$pwd\repository.xml"

    # Проверяем есть ли файл *.xml если нет то создаем новый
    if ((Test-Path -Path $Global:XML) -eq $false) {
        
        Set-MsgLog -pathLog $Global:LogPath -Msg 'File "repository.xml" not found, a new one will be created'
        Write-Host 'File "repository.xml" not found, a new one will be created'
        New-Item $Global:XML -Value '<settings Info="settings">
        <paramSchedTask Info="ScheduleTaskList">
        </paramSchedTask>
        <servers Info="ServerList">
        </servers>
        </settings>'

    }
    
    # Проверяем есть ли файл *.log если нет то создаем новый
    if ((Test-Path -Path $Global:LogPath) -eq $false) { 
        
        Set-MsgLog -pathLog $Global:LogPath -Msg 'File "applog.log" not found, a new one will be created'
        Write-Host 'File "applog.log" not found, a new one will be created'
        New-Item $Global:LogPath
    
    }

    # Инициализируем файл xml
    [xml]$Global:xmlConfig = Get-Content $Global:XML
    Write-Host "initialization xml files" # Отображаем в консоли 

    # Запускаем приложение
    Show-RpsForm
    Set-MsgLog -pathLog $Global:LogPath -Msg 'initialization xml files and Start Application'
}

#----------------------------------------------
# Создаем графический интерфейс и обрабатываем события в форме
#----------------------------------------------
function Show-RpsForm {

    # Добавляем сборку для работы с графическим интерфейсом
    [void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
		
    # Cоздаем форму и элементы управления
	[System.Windows.Forms.Application]::EnableVisualStyles()
    # Forms
    $Form = New-Object 'System.Windows.Forms.Form'
    $tabControl = New-Object 'System.Windows.Forms.TabControl'
    $t_pageMenu = New-Object 'System.Windows.Forms.TabPage'
    $t_pageConnectTo = New-Object 'System.Windows.Forms.TabPage'
    $InputForm = New-Object 'System.Windows.Forms.Form'
    # button
    $b_selectAll_Menu = New-Object 'System.Windows.Forms.Button'
    $b_selectAll_ConTo = New-Object 'System.Windows.Forms.Button'
    $b_deleteTask = New-Object 'System.Windows.Forms.Button'
    $b_createTask = New-Object 'System.Windows.Forms.Button'
    $b_loadTaskValue = New-Object 'System.Windows.Forms.Button'
    $b_upLoadFiles = New-Object 'System.Windows.Forms.Button'
    $b_removeFiles = New-Object 'System.Windows.Forms.Button'
    $b_pushTask = New-Object 'System.Windows.Forms.Button'
    $b_addServer = New-Object 'System.Windows.Forms.Button'
    $b_removeServer = New-Object 'System.Windows.Forms.Button'
    $b_removeServerMenu = New-Object 'System.Windows.Forms.Button'
    $b_connect = New-Object 'System.Windows.Forms.Button'
    $b_disconnect = New-Object 'System.Windows.Forms.Button'
    $b_inForAddServer = New-Object 'System.Windows.Forms.Button'
    $b_inForClose = New-Object 'System.Windows.Forms.Button'
    $b_SchedSave = New-Object 'System.Windows.Forms.Button' 
    # check list \ treeView
    $checkListBox_Menu = New-Object 'System.Windows.Forms.CheckedListBox'
    $checkListBox_ConTo = New-Object 'System.Windows.Forms.CheckedListBox'
    $checkListBox_MenuTask = New-Object 'System.Windows.Forms.CheckedListBox'
    $treeView_Menu = New-Object 'System.Windows.Forms.TreeView'
    # text box
    $textbox_Descript = New-Object 'System.Windows.Forms.TextBox'
    $textbox_InputForm = New-Object 'System.Windows.Forms.TextBox'
    $textbox_nameTask = New-Object 'System.Windows.Forms.TextBox'
    $textbox_Repository = New-Object 'System.Windows.Forms.TextBox'
    $textbox_local = New-Object 'System.Windows.Forms.TextBox'
    $textbox_log = New-Object 'System.Windows.Forms.TextBox'
    # lable
    $textLable_TreeView = New-Object 'System.Windows.Forms.Label'
    $textLable_Task = New-Object 'System.Windows.Forms.Label'
    $textboxLable_Descript = New-Object 'System.Windows.Forms.Label'
    $textboxLable_nameTask = New-Object 'System.Windows.Forms.Label'
    $textboxLable_Repository = New-Object 'System.Windows.Forms.Label'
    $textboxLable_local = New-Object 'System.Windows.Forms.Label'
    $textboxLable_log = New-Object 'System.Windows.Forms.Label'
    # Сombo box
    $comboBox_SchedTask = New-Object 'System.Windows.Forms.ComboBox'

	#----------------------------------------------
	# Обработчики событий
	#----------------------------------------------

    $Form_Load = {
        # При загрузке формы надо очистить $checkListBox_ConTo, и заполнить $checkListBox_Menu клиентами у кого созадны задания
        $checkListBox_ConTo.Items.Clear()
        
        # Получаем списко серверов вносим его в лист бокс на вкладке Меню
        $DataMenu = $Global:xmlConfig.SelectNodes('//settings/servers/s').Name
        $DataTask = $Global:xmlConfig.SelectNodes('//settings/paramSchedTask/Task').Name

        # В листбокс меню добавляем всех клиентов у кого установлены задания
        $checkListBox_Menu.Items.AddRange($DataMenu)
        $checkListBox_MenuTask.Items.AddRange($DataTask)
        
        # Заполняем TreeView для отображения установленных заданий на серверах
        Get-TreeViewTask # Загружаем данные
        Set-MsgLog -pathLog $Global:LogPath -Msg 'Configuration and data retrieval completed'
    } 

    ###> Логика первой страницы $t_pageConnectTo   
    $b_addServer_Click = {
        
        # Диалоговое окно для ввода данных в форму
        $r = $InputForm.ShowDialog()

        # Проверяем выбор пользователя         
        if ($r -eq [System.Windows.Forms.DialogResult]::OK){
            
            $checkListBox_ConTo.Items.AddRange($textbox_InputForm.Text) # Добавляем в список север
            $textbox_InputForm.Clear() # Очистка текстового поля

        } elseif ($r -eq [System.Windows.Forms.DialogResult]::Cancel){

            Set-MsgLog -pathLog $Global:LogPath -Msg 'User take buttons Close'

        }
    
    }
    
    $b_removeServer_Click = {
        
        # Выделяем сервера, проходимся по списку и удаляем
        while ($checkListBox_ConTo.CheckedItems.Count -gt 0) {

            $checkListBox_ConTo.Items.Remove($checkListBox_ConTo.CheckedItems[0])

        }

        # Обновляем информацию по нашему удаленном серверу в TreeView
        $treeView_Menu.Nodes.Clear() # Очищаем поле
        Get-TreeViewTask # Загружаем данные

    }
        
    $b_connect_Click = {

       # Получаем настройки из repository.xml
       $data = Get-SwitchComboBox -id $comboBox_SchedTask.text -xml $Global:xmlConfig
       $TaskName = $data.xmlName
       $Descript = $data.xmlDescript
       $PathReposit = $data.xmlPathRepository
       $PathLocal = $data.xmlPathLocal
       $LogTaskExec = $data.xmlPathLogs

       # Получаем массив серверов из вкладки "Connect to" и получаем массив со вкладки Menu
       $array_menu = @($checkListBox_Menu.Items)
       $array_conTo = @($checkListBox_ConTo.CheckedItems)
       
       # проходим по массиву выбранных значений на вкладке Connect to, сравниваем со списком на вкладке Menu
       foreach ($a in $array_conTo) {
            Write-Host "Name Server in Connect ====>" $a "<===="  # Выводим в консоль о сервере который сверям  
            
            # Проверяем есть ли указанный сервер в списке репозитория      
            $foundServ = Set-FoundServer -array1 $array_menu -array2 $a
            Write-Host "foundServ: " $foundServ

            # Если совпадения не найдены, добавляем серверв в список и создаем SchedulerTask на сервере
            if ($foundServ -ne 1) { 
                Write-Host "Server not found in list Menu page: "$a
                
                # Создаем задание для выполнения его на удаленной машине
                Invoke-Command -ScriptBlock {Register-ScheduledTask -TaskName $Using:TaskName -Description $Using:Descript -Action (New-ScheduledTaskAction -Execute 'C:\Windows\System32\Robocopy.exe' -Argument $Using:PathReposit" "$Using:PathLocal" /J /PURGE /TEE /LOG+:"$Using:LogTaskExec" /NP") -User "avi\SHP13_admin"} -ComputerName $a
                
                # Добавляем выбранный сервер в список доступных для загрузки из репозитория 
                $checkListBox_Menu.Items.AddRange($a)
                
                # Добавление сервера в файл repository.xml и обновлеяем TreeView
                New-createXmlServer -id $a -xml $Global:xmlConfig -name $a -task $TaskName -xmlSave $Global:XML
                
                # Обновляем информацию в TreeView
                $treeView_Menu.Nodes.Clear() # Очищаем поле
                Get-TreeViewTask # Загружаем данные

                # Удаляем сервер из списка на вкладке Connect to
                $checkListBox_ConTo.Items.Remove($a)

            } elseif ($foundServ -eq 1) {
                # Сервер находиться в списке на вкладке Меню, получаем имя сервера
                $show = $a
                [System.Windows.Forms.MessageBox]::Show(" $show : server is already in the list of repository",'Connected','OK','Info')

                $q = [System.Windows.Forms.MessageBox]::Show(" $show : Would you like to add another task?",'Connected','YesNo','Info')
                if ($q -eq 'Yes') {
                    Write-Host "Server not found in list Menu page: " $a
                
                    # Создаем задание для выполнения его на удаленной машине
                    Invoke-Command -ScriptBlock {Register-ScheduledTask -TaskName $Using:TaskName -Description $Using:Descript -Action (New-ScheduledTaskAction -Execute 'C:\Windows\System32\Robocopy.exe' -Argument $Using:PathReposit" "$Using:PathLocal" /J /PURGE /TEE /LOG+:"$Using:LogTaskExec" /NP") -User "avi\SHP13_admin"} -ComputerName $a
                    
                    # Добавляем еще одно задание в список сервера
                    $r = $Global:xmlConfig.SelectSingleNode("//s[@id='$a']")
                    
                    # Создаем раздел в файле xml
                    $add = $Global:xmlConfig.CreateElement("task")
                    $add.InnerText = "$TaskName"
                    
                    # Добавляем раздел по указанному пути
                    $r.AppendChild($add)
                    $Global:xmlConfig.Save($Global:XML)
                    
                    # Обновляем информацию в TreeView
                    $treeView_Menu.Nodes.Clear() # Очищаем поле
                    Get-TreeViewTask # Загружаем данные

                    # Удаляем сервер из списка на вкладке Connect to
                    $checkListBox_ConTo.Items.Remove($a)

                }

            }

        }    

    }
    
    $b_disconnect_Click = {
            
        # Получаем значение Name с textBox_nameTask
        $name = $textbox_nameTask.Text

        # Удаляем задания со всех серверов
        $array_conTo = @($checkListBox_ConTo.CheckedItems)  

        foreach ($a in $array_conTo) {
            
            Write-Host "Server Disconnect repository ====> " $a "<====" 

            Invoke-Command -ScriptBlock {Unregister-ScheduledTask -TaskName $Using:name -Confirm:$false} -ComputerName $a
            Set-MsgLog -pathLog $Global:LogPath -Msg "$a Server shutdown completed successfully"

        }

        [System.Windows.Forms.MessageBox]::Show("Info: Servers Disconnected to Repository",'Disconnect','OK','Info')
        

    }

    $b_SchedSave_Click = {

        # Сохраняем введеные задания в repository.xml
        Update-SwitchComboBox -id $comboBox_SchedTask.Text -xml $Global:xmlConfig -NameReplace $textbox_nameTask.text -pathRepositoryReplace $textbox_Repository.text -pathLocalReplace $textbox_local.text -pathLogsReplace $textbox_log.text -DescriptReplace $textbox_Descript.text -xmlSave $Global:XML
    
    }

    $b_loadTaskValue_Click = {

        Write-Host "Load Task Value" $comboBox_SchedTask.Text
        
        # Получчаем данные из xml и записываем их в текстовые поля
        $data = Get-SwitchComboBox -id $comboBox_SchedTask.text -xml $Global:xmlConfig
        $textbox_nameTask.text = $data.xmlName
        $textbox_Descript.text = $data.xmlDescript
        $textbox_Repository.text = $data.xmlPathRepository
        $textbox_local.text = $data.xmlPathLocal
        $textbox_log.text = $data.xmlPathLogs

        # Выводим в консоль то что получили из файла xml
        Write-Host $data.xmlName
        Write-Host $data.xmlPathRepository
        Write-Host $data.xmlPathLocal
        Write-Host $data.xmlPathLogs
        Write-Host $data.xmlDescript
        
    }

    $b_createTask_Click = {
        Write-Host "Create new task in repository.xml"

        # Создаем новое задание в xml файле
        New-createXmlTask -id $textbox_nameTask.text -xml $Global:xmlConfig -Name $textbox_nameTask.text -pathRepository $textbox_Repository.text -pathLocal $textbox_local.text -pathLogs $textbox_log.text -Descript $textbox_Descript.text -xmlSave $Global:XML
        
        # Обновлеяем список выпадающего меню
        $comboBox_SchedTask.DataSource = @($Global:xmlConfig.SelectNodes('//settings/paramSchedTask/Task').Name)
        $DataTask = $Global:xmlConfig.SelectNodes('//settings/paramSchedTask/Task').Name
        $checkListBox_MenuTask.Items.Clear()
        $checkListBox_MenuTask.Items.AddRange($DataTask)
        
        $n = $textbox_nameTask.text
        Set-MsgLog -pathLog $Global:LogPath -Msg "$n new task created"
    }

    $b_deleteTask_Click = {
        Write-Host "Delete task in repository.xml"

        # Удаляем необходимую задачу из файла
        Remove-SwitchXmlTask -id $textbox_nameTask.text -xml $Global:xmlConfig -xmlSave $Global:XML

        # Обновлеяем список выпадающего меню и список на вкладке Меню
        $comboBox_SchedTask.DataSource = @($Global:xmlConfig.SelectNodes('//settings/paramSchedTask/Task').Name)
        $DataTask = $Global:xmlConfig.SelectNodes('//settings/paramSchedTask/Task').Name
        $checkListBox_MenuTask.Items.Clear()
        $checkListBox_MenuTask.Items.AddRange($DataTask)
        
        $n = $textbox_nameTask.text
        Set-MsgLog -pathLog $Global:LogPath -Msg "$n task deleted"
    }

    $b_selectAll_ConTo_Click = {
        
        # Выделяем все значения в списке Connect to, если есть выделенное значение снимаем галки со всех, если нет то выделяем все
        $a = @($checkListBox_ConTo.CheckedItems)       
        if ($a) {   
            (0..($checkListBox_ConTo.Items.Count-1)) | ForEach-Object {$checkListBox_ConTo.SetItemChecked($_,$false)}
        } else {
            (0..($checkListBox_ConTo.Items.Count-1)) | ForEach-Object {$checkListBox_ConTo.SetItemChecked($_,$true)}
        }
        
    }

    ###> Логика второй страницы $t_pageMenu
    $b_upLoadFiles_Click = {
        [System.Windows.Forms.MessageBox]::Show('Download started','Upload files','ok','Info')

        # Получаем значение Name
        $name = $checkListBox_MenuTask.CheckedItems

        $array_menu = @($checkListBox_Menu.CheckedItems)
        Write-Host $array_menu

        foreach ($a in $array_menu) {
    
            Write-Host "Server start upload Files ====>" $a "<===="
            Write-Host "Upload Files: " $a
            
            Invoke-Command -ScriptBlock {Start-ScheduledTask -TaskName $Using:name} -ComputerName $a

        } 
             
    }
    
    $b_removeFiles_Click = {
        [System.Windows.Forms.MessageBox]::Show('Remove started','Remove files','ok','Error')
        
        # Получаем путь из файла xml
        $data = Get-SwitchComboBox -id $checkListBox_MenuTask.CheckedItems -xml $Global:xmlConfig
        $PathLocal = $data.xmlPathLocal 
        
        # Делаем замену ", иначе не правильно пишется путь к удалению
        $replace = $PathLocal.Replace('"',"")
        Write-Host "Path Local: "$replace
        
        # Получаем данные из списка
        $array_menu = @($checkListBox_Menu.CheckedItems)

        # Проходимся по каждому объекту из массива и удаляем необходимые каталоги в локатльном расположении.
        foreach ($a in $array_menu) {
            
            Write-Host "Server Remove Files ====> " $a "<===="    

            Invoke-Command -ScriptBlock {Remove-Item -Path $Using:replace -Force -Recurse} -ComputerName $a

        } 
        
    }

    $b_removeServerMenu_Click = {
        #[System.Windows.Forms.MessageBox]::Show('You remove server','Remove Server','YesNo','Error')
        
        # Проходим по списку удаляем все выдленные значения и переносим в список на вкладке ConnectTo
        while ($checkListBox_Menu.CheckedItems.Count -gt 0) {
            $r = $checkListBox_Menu.CheckedItems[0]

            [xml]$Global:xmlConfig = Get-Content $Global:XML

            # Удаление сервера из файла repository.xml
            Remove-ServerXml -id $r -xml $Global:xmlConfig -xmlSave $Global:XML
            
            # Удаляем из списка 
            $checkListBox_ConTo.Items.AddRange($checkListBox_Menu.CheckedItems[0])
            $checkListBox_Menu.Items.Remove($r)
            
            # Сохраняем файл
            $Global:xmlConfig.Save($Global:XML)

        }

    }

    $b_selectAll_Menu_Click = {
        # Выделяем все значения в списке Connect to, если есть выделенное значение снимаем галки со всех, если нет то выделяем все
        $a = @($checkListBox_Menu.CheckedItems)       
        if ($a) {   
            (0..($checkListBox_Menu.Items.Count-1)) | ForEach-Object {$checkListBox_Menu.SetItemChecked($_,$false)}
        } else {
            (0..($checkListBox_Menu.Items.Count-1)) | ForEach-Object {$checkListBox_Menu.SetItemChecked($_,$true)}
        }

    }

#----------------------------------------------
#Описание объектов формы
#----------------------------------------------

#
#----- Forms -----#
#
##. Form 
$Form.MinimumSize = '800,600'
$Form.MaximizeBox = $false
$Form.Text = 'Repository'
$Form.ResumeLayout($false)
$Form.StartPosition = 'CenterScreen'
$Form.Controls.Add($tabControl)
$Form.add_Load($Form_Load)

##. InputForm
$InputForm.Text = 'Enter server name'
$InputForm.Size = '300,170'
$InputForm.StartPosition = 'CenterScreen'
$InputForm.Controls.AddRange(@($b_inForAddServer,$b_inForClose,$textbox_InputForm))

##. tabControl
$tabControl.Size = '780,560'
$tabControl.Location ='2,0'
$t_pageMenu.Text = [System.String]'Menu'
$t_pageConnectTo.Text = [System.String]'Connect to'
$tabControl.Controls.Add($t_pageConnectTo)
$tabControl.Controls.Add($t_pageMenu)
$t_pageMenu.Controls.AddRange(@($textLable_TreeView,$treeView_Menu,$b_selectAll_Menu,$b_upLoadFiles,$b_removeFiles,$b_removeServerMenu,$checkListBox_Menu,$checkListBox_MenuTask,$textLable_Task))
$t_pageConnectTo.Controls.AddRange(@($b_selectAll_ConTo,$b_deleteTask,$b_createTask,$b_loadTaskValue,$comboBox_SchedTask,$b_SchedSave,$b_connect,$b_disconnect,$b_addServer,$b_removeServer,$checkListBox_ConTo,$textbox_Descript,$textboxLable_Descript,$textbox_Repository,$textboxLable_Repository,$textbox_local,$textboxLable_local,$textbox_log,$textboxLable_log,$textbox_nameTask,$textboxLable_nameTask))

#
#----- Buttons -----#
#
##. Select all Menu
$b_selectAll_Menu.Size = '30,20'
$b_selectAll_Menu.Text = 'all'
$b_selectAll_Menu.Location = '140,30'
$b_selectAll_Menu.add_Click($b_selectAll_Menu_Click)

##. Select all Connect to
$b_selectAll_ConTo.Size = '30,20'
$b_selectAll_ConTo.Text = 'all'
$b_selectAll_ConTo.Location = '140,30'
$b_selectAll_ConTo.add_Click($b_selectAll_ConTo_Click)

##. Delete Task
$b_deleteTask.Size = '100,20'
$b_deleteTask.Text = 'Delete Task'
$b_deleteTask.Location = '620,60'
$b_deleteTask.add_Click($b_deleteTask_Click)

##. Create new Task
$b_createTask.Size = '100,20'
$b_createTask.Text = 'Create Task'
$b_createTask.Location = '500,60'
$b_createTask.add_Click($b_createTask_Click)

##. Load Task to text box
$b_loadTaskValue.Size = '100,20'
$b_loadTaskValue.Text = 'Load Task'
$b_loadTaskValue.Location = '500,240'
$b_loadTaskValue.add_Click($b_loadTaskValue_Click)

##. Schedule Save
$b_SchedSave.Size = '100,20'
$b_SchedSave.Text = 'Save Schedule'
$b_SchedSave.Location = "620,240"
$b_SchedSave.add_Click($b_SchedSave_Click)

##. InputForm Add Server
$b_inForAddServer.Size = '85,30'
$b_inForAddServer.Text = "Add Server"
$b_inForAddServer.Location = '30,90'
$b_inForAddServer.DialogResult = [System.Windows.Forms.DialogResult]::OK
$b_inForAddServer.Add_Click($b_inForAddServer_Click)


##. InputForm close
$b_inForClose.Size = '85,30'
$b_inForClose.Text = "Cancel"
$b_inForClose.Location = '170,90'
$b_inForClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$b_inForClose.Add_Click($b_inForClose_Click)


##. Upload Files
$b_upLoadFiles.Size = '100,40'
$b_upLoadFiles.Text = "Upload Files"
$b_upLoadFiles.Location = '20,90'
$b_upLoadFiles.Add_Click($b_upLoadFiles_Click)

##. Remove files
$b_removeFiles.Size = '100,40'
$b_removeFiles.Text = "Remove Files"
$b_removeFiles.Location = '20,150'
$b_removeFiles.Add_Click($b_removeFiles_Click)

##. Push task
$b_pushTask.Size = '100,40'
$b_pushTask.Text = "Push Task"
$b_pushTask.Location = '20,150'
$b_pushTask.Add_Click($b_pushTask_Click)

##. Add Server
$b_addServer.Size = '78,20'
$b_addServer.Text = "Add Server"
$b_addServer.Location = '172,30'
$b_addServer.Add_Click($b_addServer_Click)

##. Remove Server Connect to
$b_removeServer.Size = '92,20'
$b_removeServer.Text = "Remove Server"
$b_removeServer.Location = '250,30'
$b_removeServer.Add_Click($b_removeServer_Click)

##. Remove Server Menu
$b_removeServerMenu.Size = '100,20'
$b_removeServerMenu.Text = "Remove Server"
$b_removeServerMenu.Location = '190,30'
$b_removeServerMenu.Add_Click($b_removeServerMenu_Click)

##. Connected
$b_connect.Size = '100,40'
$b_connect.Text = "Connected"
$b_connect.Location = '20,90'
$b_connect.Add_Click($b_connect_Click)

##. Disconnected
$b_disconnect.Size = '100,40'
$b_disconnect.Text = "Disconnected"
$b_disconnect.Location = '20,150'
$b_disconnect.Add_Click($b_disconnect_Click)

#
#----- Check List Box \ TreeView box -----#
#
##. Menu | TreeView box
$treeView_Menu.Location = '100,300'
$treeView_Menu.Size = '540,200'
##. Lable TreeView
$textLable_TreeView.Text = "Server information"
$textLable_TreeView.Location = '320,275'

##. Menu | check list PC
$checkListBox_MenuTask.Location = '400,62'
$checkListBox_MenuTask.Size = '200,200'
$checkListBox_MenuTask.CheckOnClick = $true
##. Lable Task
$textLable_Task.Text = "Acticve Task"
$textLable_Task.Location = '460,35'

##. Control menu | check list PC
$checkListBox_Menu.Location = '140,62'
$checkListBox_Menu.Size = '200,200'
$checkListBox_Menu.CheckOnClick = $true

##. Connect to | check list PC
$checkListBox_ConTo.Location = '140,62'
$checkListBox_ConTo.Size = '200,200'
$checkListBox_ConTo.CheckOnClick = $true

#
#----- Text Box \ Lable -----#
#
##. Description
$textbox_Descript.Location = '500,210'
$textbox_Descript.Size = '250,40'
$textbox_Descript.Text = $Descript
##. Description
$textboxLable_Descript.Text = "Description"
$textboxLable_Descript.Location = '434,213'

##. Add Server
$textbox_InputForm.Location = '30,30'
$textbox_InputForm.Size = '220,20'

##. Name Task 
$textbox_nameTask.Location = '500,90'
$textbox_nameTask.Size = '250,40'
$textbox_nameTask.Text = $TaskName
##. Lable Name Task
$textboxLable_nameTask.Text = "Name Task"
$textboxLable_nameTask.Location = '435,93'

##. Path Repository
$textbox_Repository.Location = '500,120'
$textbox_Repository.Size = '250,40'
$textbox_Repository.Text = $PathReposit
##. Lable Path Repository
$textboxLable_Repository.Text = "Path Repository"
$textboxLable_Repository.Location = '410,123'

##. Path Logs
$textbox_local.Location = '500,150'
$textbox_local.Size = '250,40'
$textbox_local.Text = $PathLocal
##. Lable Path Logs
$textboxLable_local.Text = "Path Local"
$textboxLable_local.Location = '439,153'

##. Path Logs
$textbox_log.Location = '500,180'
$textbox_log.Size = '250,40'
$textbox_log.Text = $LogTaskExec
##. Lable Path Logs
$textboxLable_log.Text = "Path Logs"
$textboxLable_log.Location = '439,183'
#
#----- Combo Box \ Lable -----#
#
##. Schedule task
$comboBox_SchedTask.Location = '500,30'
$comboBox_SchedTask.Size = '250,40'
$comboBox_SchedTask.DataSource = @($Global:xmlConfig.SelectNodes('//settings/paramSchedTask/Task').Name)

return $Form.ShowDialog()
}

Start-Prog($CommandLine)