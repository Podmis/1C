# Date:        20/06/2016
# Author:      Podmis
# Description: PS script to restart the Agent Service 1C and the SQL Server
# Version:     1.0

# /////////////////////////////////////
# //////// ��������� � �������

# ������� ��������� ��������� � log ����
function GetTime
{

    $dt = (get-date -format "dd.MM.yyyy|HH:mm:ss")+","+(get-date -UFormat "%a")

    Return $dt
}

# ������� ��������� � log ���� �����
function Write_log ($text)
{
    $CurText = (GetTime)+" "+$text

    Add-Content $LogFile -Value $CurText
}

# ���������� ��������
 Function Restart_Service($Name, $S, $Text, $Status="s")
{
    $Service = Get-Service $Name

    Write_log $Text
    Start-Sleep -s $S
   
    if ($Status -eq "s" -and $Service.Status -eq "Running") {$Service.Stop()}
    elseif ($Status -eq "r" -and $Service.Status -eq "Stopped") {$Service.Start()}
}

# ���������
 Function Get_Num
{
    $�+=1
    Return $�
}

# /////////////////////////////////////
# //////// ��������

# 1. ������������� ����������
$Agnt1C        = '1C:Enterprise 8.3 Server Agent'     # ������ ����� ������� 1�:�����������
$SqlAgnt       = 'SQLAgent$MSSQLSERVER2008'           # ������ ����� SQL SERVER
$SqlServer     = 'MSSQL$MSSQLSERVER2008'              # C����� SQL SERVER    
$AddrAgnt1C    = "tcp://localhost:1540"               # ������� � ������� ��������� 1540 � �������� 1541
                                                      # ������������ ������� localhost
$LogFile       = "D:\1Cv82\Admin\RestartService.log"  # ���� � ��� �����
$admin_user    = ""
$admin_pass    = ""

# 2. ������� ������ ��� ������� log
If ((Test-Path $LogFile) �eq $false) {

    New-Item -ItemType file -Path $LogFile

}

Add-Content $LogFile ""
Add-Content $LogFile ((GetTime) + " ------ ������ ���������� ������� ------")
Add-Content $LogFile ""

# 3. �������� ������� � ���������� �� ������� 1�:�����������
Write_log "1. ������ �������� ������� � ���������� ������� 1�"

# 3.1 ������� COM����������� 1�:�����������
$Connector  = New-Object -COMObject "V83.COMConnector"
$Ragent     = $Connector.ConnectAgent($AddrAgnt1C)

# 3.2 ������� ������ ��������� �������
$clusters = $Ragent.GetClusters()

$DelSession = 0
foreach ($cluster in $clusters)
{
    # 3.3 ��������� �������������� �������������� �������� ��������
    $Ragent.Authenticate($cluster, $admin_user, $admin_pass)    
    
    # 3.3 �������� ������ �������, ���������� � ������ ���������
    $Sessions = $Ragent.GetSessions($cluster)

    foreach ($Session in $Sessions)
    {
       if ($Session.AppID -eq "SrvrConsole") { Continue }

       $Ragent.TerminateSession($cluster, $Session)
       $DelSession+=1
    }
  }

if ($DelSession -eq 0) 
    {Write_log "1.1 �������� ������� �� ����������"} 
else 
    {Write_log ("1.1 ������� " + $DelSession + " ������")}

# 4. ��������� ����� Windows
Restart_Service ($Agnt1C)     5 "2. ��������� ������ ������� 1�:�����������" s
Restart_Service ($SqlAgnt)    5 "3. ��������� ������ SQL Server"             s
Restart_Service ($SqlServer) 10 "4. ��������� ������ SQL Server"             s

# 5. ������ ����� Windows
Restart_Service ($SqlServer) 10 "5. ������ ������ SQL Server"             r
Restart_Service ($SqlAgnt)    5 "6. ������ ������ SQL Server"             r
Restart_Service ($Agnt1C)     5 "7. ������ ������ ������� 1�:�����������" r

# 6. ����� �������
Add-Content $LogFile ""
Add-Content $LogFile ((GetTime) + " ------ ����� ���������� ������� ------")
