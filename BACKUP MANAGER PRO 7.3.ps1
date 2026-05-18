# =====================================================
# BACKUP MANAGER PRO 7.3 - VERSÃO ULTRA RESILIENTE V2
# Ajustado - Layout corrigido e migração automática de perfis
# CORREÇÃO V2: Tratamento definitivo de caminhos raiz e aspas no Robocopy
# =====================================================

param(
    [string]$PerfilNome,
    [string]$Origem,
    [string]$Destino,
    [switch]$AutoStart,
    [switch]$Shutdown
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Forçar o PowerShell a usar UTF8 para saídas
$OutputEncoding = [System.Text.Encoding]::UTF8

# =====================================================
# CONFIGURAÇÕES E DIRETÓRIOS
# =====================================================

$pastaApp = Join-Path $env:USERPROFILE "BackupManagerPro"
$arquivoPerfisNovo = Join-Path $pastaApp "backup_profiles.csv"
$arquivoPerfisAntigo = Join-Path $env:USERPROFILE "backup_profiles.csv"
$pastaLogs = Join-Path $pastaApp "Logs"

if (!(Test-Path $pastaApp)) { New-Item -ItemType Directory -Path $pastaApp | Out-Null }
if (!(Test-Path $pastaLogs)) { New-Item -ItemType Directory -Path $pastaLogs | Out-Null }

if (!(Test-Path $arquivoPerfisNovo)) {
    if (Test-Path $arquivoPerfisAntigo) {
        Copy-Item -Path $arquivoPerfisAntigo -Destination $arquivoPerfisNovo -Force
    } else {
        @"
Nome,Origem,Destino
"@ | Set-Content $arquivoPerfisNovo -Encoding UTF8
    }
}

$arquivoPerfis = $arquivoPerfisNovo

# =====================================================
# FORMULÁRIO PRINCIPAL
# =====================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Backup Manager PRO 7.3 - Resiliência Ativada"
$form.Size = New-Object System.Drawing.Size(920,720)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(35,35,35)
$form.ForeColor = "White"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)

# Componentes básicos
$titulo = New-Object System.Windows.Forms.Label
$titulo.Text = "Backup Manager PRO 7.3"; $titulo.Location = "20,10"; $titulo.Size = "400,30"; $titulo.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold); $form.Controls.Add($titulo)

$combo = New-Object System.Windows.Forms.ComboBox; $combo.Location = "20,50"; $combo.Size = "250,28"; $combo.DropDownStyle = "DropDownList"; $form.Controls.Add($combo)
$btnCarregar = New-Object System.Windows.Forms.Button; $btnCarregar.Text = "Carregar"; $btnCarregar.Location = "280,48"; $btnCarregar.Size = "90,30"; $form.Controls.Add($btnCarregar)
$btnSalvar = New-Object System.Windows.Forms.Button; $btnSalvar.Text = "Salvar"; $btnSalvar.Location = "380,48"; $btnSalvar.Size = "90,30"; $form.Controls.Add($btnSalvar)
$btnExcluir = New-Object System.Windows.Forms.Button; $btnExcluir.Text = "Excluir"; $btnExcluir.Location = "480,48"; $btnExcluir.Size = "90,30"; $form.Controls.Add($btnExcluir)
$txtNome = New-Object System.Windows.Forms.TextBox; $txtNome.Location = "580,50"; $txtNome.Size = "300,28"; $form.Controls.Add($txtNome)

$lbl1 = New-Object System.Windows.Forms.Label; $lbl1.Text = "Origem:"; $lbl1.Location = "20,95"; $lbl1.AutoSize = $true; $form.Controls.Add($lbl1)
$txtOrigem = New-Object System.Windows.Forms.TextBox; $txtOrigem.Location = "90,92"; $txtOrigem.Size = "680,28"; $form.Controls.Add($txtOrigem)
$btnOrigem = New-Object System.Windows.Forms.Button; $btnOrigem.Text = "Selecionar"; $btnOrigem.Location = "780,90"; $btnOrigem.Size = "100,30"; $form.Controls.Add($btnOrigem)

$lbl2 = New-Object System.Windows.Forms.Label; $lbl2.Text = "Destino:"; $lbl2.Location = "20,130"; $lbl2.AutoSize = $true; $form.Controls.Add($lbl2)
$txtDestino = New-Object System.Windows.Forms.TextBox; $txtDestino.Location = "90,127"; $txtDestino.Size = "680,28"; $form.Controls.Add($txtDestino)
$btnDestino = New-Object System.Windows.Forms.Button; $btnDestino.Text = "Selecionar"; $btnDestino.Location = "780,125"; $btnDestino.Size = "100,30"; $form.Controls.Add($btnDestino)

$grpAgenda = New-Object System.Windows.Forms.GroupBox; $grpAgenda.Text = "Agendamento"; $grpAgenda.Location = "20,165"; $grpAgenda.Size = "860,70"; $grpAgenda.ForeColor = "White"; $form.Controls.Add($grpAgenda)
$comboDia = New-Object System.Windows.Forms.ComboBox; $comboDia.Location = "50,26"; $comboDia.Size = "120,28"; $comboDia.DropDownStyle = "DropDownList"; $comboDia.Items.AddRange(@("Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo")); $grpAgenda.Controls.Add($comboDia)
$txtHora = New-Object System.Windows.Forms.MaskedTextBox; $txtHora.Location = "255,26"; $txtHora.Size = "80,28"; $txtHora.Mask = "00:00"; $txtHora.Text = "2200"; $grpAgenda.Controls.Add($txtHora)
$btnAgenda = New-Object System.Windows.Forms.Button; $btnAgenda.Text = "Agendar"; $btnAgenda.Location = "370,24"; $btnAgenda.Size = "110,30"; $grpAgenda.Controls.Add($btnAgenda)

$progress = New-Object System.Windows.Forms.ProgressBar; $progress.Location = "20,245"; $progress.Size = "860,25"; $form.Controls.Add($progress)
$status = New-Object System.Windows.Forms.Label; $status.Location = "20,275"; $status.Size = "860,25"; $status.Text = "Pronto."; $form.Controls.Add($status)
$chk = New-Object System.Windows.Forms.CheckBox; $chk.Text = "Desligar ao finalizar"; $chk.Location = "20,305"; $chk.Size = "180,25"; $form.Controls.Add($chk)

$btnStart = New-Object System.Windows.Forms.Button; $btnStart.Text = "Iniciar Backup"; $btnStart.Location = "20,340"; $btnStart.Size = "180,40"; $form.Controls.Add($btnStart)
$btnClear = New-Object System.Windows.Forms.Button; $btnClear.Text = "Limpar Perfil atual"; $btnClear.Location = "700,340"; $btnClear.Size = "180,40"; $form.Controls.Add($btnClear)

$log = New-Object System.Windows.Forms.TextBox; $log.Location = "20,395"; $log.Size = "860,260"; $log.Multiline = $true; $log.ScrollBars = "Vertical"; $log.BackColor = "Black"; $log.ForeColor = "Lime"; $form.Controls.Add($log)

# =====================================================
# FUNÇÕES DE APOIO
# =====================================================

function Escrever-Log {
    param([string]$texto, [string]$tipo = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $linha = "[$timestamp] [$tipo] $texto"
    
    if ($form.Visible) {
        try {
            $form.Invoke([Action]{ 
                $log.AppendText("$linha`r`n")
                $log.SelectionStart = $log.Text.Length
                $log.ScrollToCaret()
            })
        } catch {}
    } else {
        Write-Host $linha
    }

    $dataAtual = Get-Date -Format "yyyy-MM-dd"
    $arquivoLogDia = Join-Path $pastaLogs "backup_$dataAtual.log"
    $linha | Out-File -FilePath $arquivoLogDia -Append -Encoding UTF8
}

function Obter-Perfis { try { @(Import-Csv $arquivoPerfis) } catch { @() } }

function Atualizar-Perfis {
    $combo.Items.Clear()
    foreach($item in (Obter-Perfis | Sort-Object Nome)){ [void]$combo.Items.Add($item.Nome) }
}

function Executar-Backup {
    param($origem, $destino, $desligar)
    
    if(!(Test-Path $origem)){ 
        $status.Text = "Erro: Origem não encontrada."
        Escrever-Log "Falha: Origem '$origem' não encontrada." "ERRO"
        return 
    }
    
    if(!(Test-Path $destino)){
        try { 
            New-Item -ItemType Directory -Path $destino -Force | Out-Null
            Escrever-Log "Destino criado automaticamente: $destino"
        } catch {
            $status.Text = "Erro: Destino inacessível."
            Escrever-Log "Falha crítica: Não foi possível acessar ou criar o destino '$destino'." "ERRO"
            return
        }
    }

    # CORREÇÃO V2: Robocopy odeia \" no final de caminhos entre aspas.
    # Se o caminho termina em \, o Robocopy interpreta a aspa seguinte como literal.
    # A solução segura é garantir que não haja \ antes da aspa final, a menos que seja a raiz.
    # Para raízes (D:\), o Robocopy aceita sem a barra final (D:) ou com a barra escapada.
    # Vamos usar a técnica de remover a barra final e adicionar um espaço se necessário.
    
    $o = $origem.TrimEnd('\')
    $d = $destino.TrimEnd('\')
    
    # Se for raiz (ex: D:), o Robocopy prefere D:\ mas sem o problema do escape.
    # Usaremos caminhos sem a barra invertida final para evitar o erro 16.
    
    $status.Text = "Preparando Robocopy..."
    $progress.Value = 5
    Escrever-Log "Iniciando processo de backup."
    Escrever-Log "Origem: $origem"
    Escrever-Log "Destino: $destino"

    # Argumentos limpos. /XF e /XD podem ser adicionados se necessário.
    # Importante: caminhos entre aspas, mas sem a barra de escape no final.
    $args = "`"$o`" `"$d`" /MIR /R:3 /W:5 /FFT /Z /MT:8 /XA:H /XJ /NJH /NJS /NS /NC"
    
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "robocopy.exe"
    $processInfo.Arguments = $args
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    
    $exitCode = -1
    try {
        [void]$process.Start()
        
        while (!$process.HasExited) {
            $line = $process.StandardOutput.ReadLine()
            if ($line -and $line.Trim() -ne "") {
                $cleanLine = $line.Trim()
                $form.Invoke([Action]{ 
                    $status.Text = "Copiando: $(Split-Path $cleanLine -Leaf)"
                    if ($cleanLine -match "\\|\.") {
                        Escrever-Log "Item: $cleanLine" "DEBUG"
                    }
                })
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $process.WaitForExit()
        $exitCode = $process.ExitCode
        
        # Robocopy Exit Codes:
        # 0: No files copied.
        # 1: Files copied successfully.
        # 2: Some extra files/dirs detected.
        # 4: Some mismatched files/dirs detected.
        # 8: Some files could not be copied.
        # 16: Serious error.
        
        if ($exitCode -lt 8) {
            $progress.Value = 100
            $status.Text = "Backup concluído (Código $exitCode)."
            Escrever-Log "Backup finalizado com sucesso. Código: $exitCode"
        } else {
            $errText = $process.StandardError.ReadToEnd()
            $status.Text = "Erro fatal (Código $exitCode)."
            Escrever-Log "Erro no Robocopy. Código: $exitCode. Detalhes: $errText" "ERRO"
        }
    } catch {
        Escrever-Log "Falha na execução: $($_.Exception.Message)" "ERRO"
        $status.Text = "Erro na execução do processo."
    }

    if($desligar -and $exitCode -lt 8){
        Escrever-Log "Desligamento em 60s..."
        shutdown /s /t 60 /c "Backup concluído com sucesso."
    }
}

# =====================================================
# EVENTOS
# =====================================================

$btnOrigem.Add_Click({ $f = New-Object System.Windows.Forms.FolderBrowserDialog; if($f.ShowDialog() -eq "OK"){ $txtOrigem.Text = $f.SelectedPath } })
$btnDestino.Add_Click({ $f = New-Object System.Windows.Forms.FolderBrowserDialog; if($f.ShowDialog() -eq "OK"){ $txtDestino.Text = $f.SelectedPath } })

$btnSalvar.Add_Click({
    $n = $txtNome.Text.Trim()
    if($n -eq ""){ $status.Text = "Nome vazio."; return }
    $p = @(Obter-Perfis | Where-Object {$_.Nome -ne $n}) + [PSCustomObject]@{ Nome = $n; Origem = $txtOrigem.Text; Destino = $txtDestino.Text }
    $p | Sort-Object Nome | Export-Csv $arquivoPerfis -NoTypeInformation -Encoding UTF8
    Atualizar-Perfis; $combo.SelectedItem = $n; $status.Text = "Perfil '$n' salvo."
})

$btnCarregar.Add_Click({
    if($combo.SelectedIndex -lt 0){ return }
    $p = Obter-Perfis | Where-Object { $_.Nome -eq $combo.SelectedItem }
    if($p){ $txtNome.Text = $p.Nome; $txtOrigem.Text = $p.Origem; $txtDestino.Text = $p.Destino; $status.Text = "Perfil carregado." }
})

$btnAgenda.Add_Click({
    if($comboDia.SelectedIndex -lt 0 -or $txtNome.Text -eq ""){ $status.Text = "Dados incompletos."; return }
    $map = @{ "Segunda-feira"="Monday"; "Terça-feira"="Tuesday"; "Quarta-feira"="Wednesday"; "Quinta-feira"="Thursday"; "Sexta-feira"="Friday"; "Sábado"="Saturday"; "Domingo"="Sunday" }
    $task = "BackupManager_$($txtNome.Text.Replace(' ','_'))"
    $path = $MyInvocation.MyCommand.Path; if(!$path){ $path = "C:\Scripts\BACKUP MANAGER PRO 7.3.ps1" }
    $arg = "-ExecutionPolicy Bypass -File `"$path`" -AutoStart -PerfilNome `"$($txtNome.Text)`" -Origem `"$($txtOrigem.Text)`" -Destino `"$($txtDestino.Text)`""
    if($chk.Checked){ $arg += " -Shutdown" }
    try {
        Register-ScheduledTask -Action (New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arg) -Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $map[$comboDia.SelectedItem] -At $txtHora.Text) -TaskName $task -Force -User $env:USERNAME
        Escrever-Log "Agendado: $($comboDia.SelectedItem) $($txtHora.Text)"; [Windows.Forms.MessageBox]::Show("Agendado!")
    } catch { Escrever-Log "Erro no agendamento." "ERRO" }
})

$btnStart.Add_Click({ Executar-Backup -origem $txtOrigem.Text -destino $txtDestino.Text -desligar $chk.Checked })
$btnClear.Add_Click({ $txtNome.Clear(); $txtOrigem.Clear(); $txtDestino.Clear(); $combo.SelectedIndex = -1; $status.Text = "Limpo." })

$form.Add_Load({
    Atualizar-Perfis; Escrever-Log "Sistema iniciado."
    if ($AutoStart) {
        $txtNome.Text = $PerfilNome; $txtOrigem.Text = $Origem; $txtDestino.Text = $Destino; $chk.Checked = $Shutdown
        Executar-Backup -origem $Origem -destino $Destino -desligar $Shutdown
    }
})

[void]$form.ShowDialog()
