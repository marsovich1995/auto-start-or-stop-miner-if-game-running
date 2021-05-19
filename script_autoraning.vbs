Option Explicit
dim target_app, target_app_params,games_all, hidden_window, scriptdir, objItems, games, miner,command
'################################################
target_app = "t-rex.exe"
target_app_params = "-a kawpow -o stratum+tcp://rvn.2miners.com:6060 -u RNQMq5Gyzr83PwYNFig3DajVujpjS7vKbN.rig -p x -i 10"
games_all = "witcher3.exe csgo.exe Fallout4.exe calc.exe" 'one or more applications separated by a space (case sensitive)
hidden_window=False
'################################################

scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) 'current folder
Set objItems = running_apps() 'list running all application

games = status_apps(games_all, objItems) '1 if game is running
miner = status_app(target_app,objItems)  '1 if target app is running

if  games and miner Then ' 1 and 1
 
Shell_hidden("taskkill /f /im " & target_app) 'close target app
ElseIf not (games or miner) Then  ' 0 and 0
' a = "cmd.exe /c " & scriptdir & "\" & "calce.exe"
command = "cmd.exe /c " & scriptdir & "\" & target_app & " " & target_app_params ' command to running application
    If hidden_window Then
        Shell_hidden(command) 'hidden running
    ElseIf not hidden_window Then
        Shell(command) 'running in a minimized window
    End If

End If


Function status_app(app, run_apps)
dim run_app,flag
flag=False
For Each run_app In run_apps
if app = run_app.Name Then
    flag = True
    Exit For
End If
Next
status_app = flag
End Function


Function status_apps(apps, run_apps)
    dim flag, run_app,app
    flag=False
    apps = Split(apps)
    ' WScript.Echo TypeName(apps)
    For Each run_app In run_apps
        For Each app in apps
        '   WScript.Echo app
            if app = run_app.Name Then
                flag=True
                Exit For
            End If
        Next
        if flag Then Exit For
    Next
    status_apps = flag
End Function


Function Shell(cmd)
    Shell = WScript.CreateObject("WScript.Shell").Run(cmd,2,False)
End Function


Function Shell_hidden(cmd)
    Shell_hidden = WScript.CreateObject("WScript.Shell").Run(cmd,0,False)
End Function


Function running_apps()
    dim sComputerName,objWMIService,sQuery
    sComputerName = "."
    Set objWMIService = GetObject("winmgmts:\\" & sComputerName & "\root\cimv2")
    sQuery = "SELECT * FROM Win32_Process"
    Set running_apps = objWMIService.ExecQuery(sQuery)

End Function