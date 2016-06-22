# Date:        20/06/2016
# Author:      Podmis
# Description: PS script to restart the Agent Service 1C and the SQL Server
# Version:     1.0

# /////////////////////////////////////
# //////// ПРОЦЕДУРЫ И ФУНКЦИИ

# функция добавляет сообщения в log файл
function GetTime
{

    $dt = (get-date -format "dd.MM.yyyy|HH:mm:ss")+","+(get-date -UFormat "%a")

    Return $dt
}

# функция добавляет в log файл текст
function Write_log ($text)
{
    $CurText = (GetTime)+" "+$text

    Add-Content $LogFile -Value $CurText
}

# Управление службами
 Function Restart_Service($Name, $S, $Text, $Status="s")
{
    $Service = Get-Service $Name

    Write_log $Text
    Start-Sleep -s $S
   
    if ($Status -eq "s" -and $Service.Status -eq "Running") {$Service.Stop()}
    elseif ($Status -eq "r" -and $Service.Status -eq "Stopped") {$Service.Start()}
}

# Нумерация
 Function Get_Num
{
    $а+=1
    Return $а
}

# /////////////////////////////////////
# //////// ОСНОВНОЙ

# 1. Инициализация параметров
$Agnt1C        = '1C:Enterprise 8.3 Server Agent'     # Служба агент сервера 1С:Предприятия
$SqlAgnt       = 'SQLAgent$MSSQLSERVER2008'           # Служба агент SQL SERVER
$SqlServer     = 'MSSQL$MSSQLSERVER2008'              # Cлужба SQL SERVER    
$AddrAgnt1C    = "tcp://localhost:1540"               # Соедние с рабочим процессом 1540 в кластере 1541
                                                      # центрального сервера localhost
$LogFile       = "D:\1Cv82\Admin\RestartService.log"  # Путь к лог файлу
$admin_user    = ""
$admin_pass    = ""

# 2. Создать журнал для ведения log
If ((Test-Path $LogFile) –eq $false) {

    New-Item -ItemType file -Path $LogFile

}

Add-Content $LogFile ""
Add-Content $LogFile ((GetTime) + " ------ Начало выполнения скрипта ------")
Add-Content $LogFile ""

# 3. Удаление сеансов и соединений на сервере 1С:предприятия
Write_log "1. Начало удаления сеансов и соединений сервера 1С"

# 3.1 Создать COMСоединитель 1С:Предприятия
$Connector  = New-Object -COMObject "V83.COMConnector"
$Ragent     = $Connector.ConnectAgent($AddrAgnt1C)

# 3.2 Получим массив кластеров сервера
$clusters = $Ragent.GetClusters()

$DelSession = 0
foreach ($cluster in $clusters)
{
    # 3.3 Выполняет аутентификацию администратора кластера серверов
    $Ragent.Authenticate($cluster, $admin_user, $admin_pass)    
    
    # 3.3 Получает список сеансов, работающих с данным кластером
    $Sessions = $Ragent.GetSessions($cluster)

    foreach ($Session in $Sessions)
    {
       if ($Session.AppID -eq "SrvrConsole") { Continue }

       $Ragent.TerminateSession($cluster, $Session)
       $DelSession+=1
    }
  }

if ($DelSession -eq 0) 
    {Write_log "1.1 Открытых сеансов не обнаружено"} 
else 
    {Write_log ("1.1 Удалено " + $DelSession + " сеанса")}

# 4. Остановка служб Windows
Restart_Service ($Agnt1C)     5 "2. Остановка службы сервера 1С:Предприятия" s
Restart_Service ($SqlAgnt)    5 "3. Остановка агента SQL Server"             s
Restart_Service ($SqlServer) 10 "4. Остановка службы SQL Server"             s

# 5. Запуск служб Windows
Restart_Service ($SqlServer) 10 "5. Запуск службы SQL Server"             r
Restart_Service ($SqlAgnt)    5 "6. Запуск агента SQL Server"             r
Restart_Service ($Agnt1C)     5 "7. Запуск службы сервера 1С:Предприятия" r

# 6. Конец скрипта
Add-Content $LogFile ""
Add-Content $LogFile ((GetTime) + " ------ Конец выполнения скрипта ------")
