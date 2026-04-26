# =====================================================
# BACKUP MANAGER PRO 7.3
# Revisado e corrigido
# Melhorias:
# - Salva múltiplos perfis sem erro
# - Atualiza perfil existente
# - Botão Limpar
# - Melhor tratamento do CSV
# - Validação básica de campos
# - Interface mantida
# =====================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =====================================================
# ARQUIVO DE PERFIS
# =====================================================

$arquivoPerfis = Join-Path $env:USERPROFILE "backup_profiles.csv"

if (!(Test-Path $arquivoPerfis)) {
    @"
Nome,Origem,Destino
"@ | Set-Content $arquivoPerfis -Encoding UTF8
}

# =====================================================
# FORM
# =====================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Backup Manager PRO 7.3"
$form.Size = New-Object System.Drawing.Size(860,680)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(35,35,35)
$form.ForeColor = "White"

# =====================================================
# PERFIL
# =====================================================

$combo = New-Object System.Windows.Forms.ComboBox
$combo.Location = "20,20"
$combo.Size = "280,23"
$combo.DropDownStyle = "DropDownList"
$form.Controls.Add($combo)

$btnCarregar = New-Object System.Windows.Forms.Button
$btnCarregar.Text = "Carregar Perfil"
$btnCarregar.Location = "320,18"
$btnCarregar.Size = "110,28"
$form.Controls.Add($btnCarregar)

$btnSalvar = New-Object System.Windows.Forms.Button
$btnSalvar.Text = "Salvar Perfil"
$btnSalvar.Location = "450,18"
$btnSalvar.Size = "110,28"
$form.Controls.Add($btnSalvar)

$txtNome = New-Object System.Windows.Forms.TextBox
$txtNome.Location = "580,20"
$txtNome.Size = "220,23"
$form.Controls.Add($txtNome)

# =====================================================
# ORIGEM
# =====================================================

$lbl1 = New-Object System.Windows.Forms.Label
$lbl1.Text = "Origem:"
$lbl1.Location = "20,60"
$lbl1.AutoSize = $true
$form.Controls.Add($lbl1)

$txtOrigem = New-Object System.Windows.Forms.TextBox
$txtOrigem.Location = "90,58"
$txtOrigem.Size = "620,23"
$form.Controls.Add($txtOrigem)

$btnOrigem = New-Object System.Windows.Forms.Button
$btnOrigem.Text = "Selecionar"
$btnOrigem.Location = "730,56"
$btnOrigem.Size = "90,28"
$form.Controls.Add($btnOrigem)

# =====================================================
# DESTINO
# =====================================================

$lbl2 = New-Object System.Windows.Forms.Label
$lbl2.Text = "Destino:"
$lbl2.Location = "20,95"
$lbl2.AutoSize = $true
$form.Controls.Add($lbl2)

$txtDestino = New-Object System.Windows.Forms.TextBox
$txtDestino.Location = "90,93"
$txtDestino.Size = "620,23"
$form.Controls.Add($txtDestino)

$btnDestino = New-Object System.Windows.Forms.Button
$btnDestino.Text = "Selecionar"
$btnDestino.Location = "730,91"
$btnDestino.Size = "90,28"
$form.Controls.Add($btnDestino)

# =====================================================
# AGENDAMENTO
# =====================================================

$lblDia = New-Object System.Windows.Forms.Label
$lblDia.Text = "Dia:"
$lblDia.Location = "20,132"
$lblDia.AutoSize = $true
$form.Controls.Add($lblDia)

$comboDia = New-Object System.Windows.Forms.ComboBox
$comboDia.Location = "55,130"
$comboDia.Size = "100,23"
$comboDia.Items.AddRange(@("MON","TUE","WED","THU","FRI","SAT","SUN"))
$form.Controls.Add($comboDia)

$lblHora = New-Object System.Windows.Forms.Label
$lblHora.Text = "Hora:"
$lblHora.Location = "180,132"
$lblHora.AutoSize = $true
$form.Controls.Add($lblHora)

$txtHora = New-Object System.Windows.Forms.TextBox
$txtHora.Location = "225,130"
$txtHora.Size = "80,23"
$txtHora.Text = "22:00"
$form.Controls.Add($txtHora)

$btnAgenda = New-Object System.Windows.Forms.Button
$btnAgenda.Text = "Agendar Perfil"
$btnAgenda.Location = "330,128"
$btnAgenda.Size = "140,28"
$form.Controls.Add($btnAgenda)

$btnLimpar = New-Object System.Windows.Forms.Button
$btnLimpar.Text = "Limpar"
$btnLimpar.Location = "480,128"
$btnLimpar.Size = "100,28"
$form.Controls.Add($btnLimpar)

# =====================================================
# STATUS
# =====================================================

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = "20,170"
$progress.Size = "800,28"
$form.Controls.Add($progress)

$status = New-Object System.Windows.Forms.Label
$status.Location = "20,205"
$status.Size = "800,25"
$status.Text = "Pronto."
$form.Controls.Add($status)

# =====================================================
# CHECKBOX
# =====================================================

$chk = New-Object System.Windows.Forms.CheckBox
$chk.Text = "Desligar ao finalizar"
$chk.Location = "20,235"
$chk.Size = "180,25"
$chk.ForeColor = "White"
$form.Controls.Add($chk)

# =====================================================
# BACKUP
# =====================================================

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Iniciar Backup"
$btnStart.Location = "20,270"
$btnStart.Size = "180,40"
$form.Controls.Add($btnStart)

# =====================================================
# LOG
# =====================================================

$log = New-Object System.Windows.Forms.TextBox
$log.Location = "20,330"
$log.Size = "800,290"
$log.Multiline = $true
$log.ScrollBars = "Vertical"
$log.BackColor = "Black"
$log.ForeColor = "Lime"
$form.Controls.Add($log)

# =====================================================
# OBJETOS
# =====================================================

$folder = New-Object System.Windows.Forms.FolderBrowserDialog

# =====================================================
# FUNÇÕES
# =====================================================

function Obter-Perfis {

    if (!(Test-Path $arquivoPerfis)) { return @() }

    try {
        return @(Import-Csv $arquivoPerfis)
    }
    catch {
        return @()
    }
}

function Atualizar-Perfis {

    $combo.Items.Clear()

    $dados = Obter-Perfis | Sort-Object Nome

    foreach($item in $dados){
        [void]$combo.Items.Add($item.Nome)
    }
}

function Limpar-Campos {

    $combo.SelectedIndex = -1
    $txtNome.Clear()
    $txtOrigem.Clear()
    $txtDestino.Clear()
    $comboDia.Text = ""
    $txtHora.Text = "22:00"
    $progress.Value = 0
    $chk.Checked = $false
    $log.Clear()
    $status.Text = "Campos limpos."
}

Atualizar-Perfis

# =====================================================
# EVENTOS
# =====================================================

$btnOrigem.Add_Click({
    if($folder.ShowDialog() -eq "OK"){
        $txtOrigem.Text = $folder.SelectedPath
    }
})

$btnDestino.Add_Click({
    if($folder.ShowDialog() -eq "OK"){
        $txtDestino.Text = $folder.SelectedPath
    }
})

$btnSalvar.Add_Click({

    $nome = $txtNome.Text.Trim()
    $origem = $txtOrigem.Text.Trim()
    $destino = $txtDestino.Text.Trim()

    if($nome -eq ""){
        $status.Text = "Informe o nome do perfil."
        return
    }

    $dados = @(Obter-Perfis | Where-Object { $_.Nome -ne $nome })

    $novo = [PSCustomObject]@{
        Nome = $nome
        Origem = $origem
        Destino = $destino
    }

    @($dados + $novo) |
        Sort-Object Nome |
        Export-Csv $arquivoPerfis -NoTypeInformation -Encoding UTF8

    Atualizar-Perfis
    $combo.SelectedItem = $nome
    $status.Text = "Perfil salvo com sucesso."
})

$btnCarregar.Add_Click({

    if($combo.SelectedIndex -lt 0){ return }

    $nome = $combo.SelectedItem.ToString()

    $perfil = Obter-Perfis | Where-Object { $_.Nome -eq $nome } | Select-Object -First 1

    if($perfil){
        $txtNome.Text = $perfil.Nome
        $txtOrigem.Text = $perfil.Origem
        $txtDestino.Text = $perfil.Destino
        $status.Text = "Perfil carregado."
    }
})

$btnLimpar.Add_Click({
    Limpar-Campos
})

# =====================================================
# AGENDAMENTO
# =====================================================

$btnAgenda.Add_Click({

    $nome = $txtNome.Text.Trim()
    $hora = $txtHora.Text.Trim()

    if($nome -eq ""){
        $status.Text = "Informe o nome do perfil."
        return
    }

    if($hora -notmatch '^\d{2}:\d{2}$'){
        $status.Text = "Hora inválida. Use HH:MM"
        return
    }

    switch ($comboDia.Text) {
        "MON" { $dia = "SEG" }
        "TUE" { $dia = "TER" }
        "WED" { $dia = "QUA" }
        "THU" { $dia = "QUI" }
        "FRI" { $dia = "SEX" }
        "SAT" { $dia = "SAB" }
        "SUN" { $dia = "DOM" }
        default {
            $status.Text = "Selecione um dia."
            return
        }
    }

    $comando = 'powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\backup_profissional.ps1"'

    schtasks /create `
        /sc weekly `
        /d $dia `
        /tn "Backup_$nome" `
        /tr $comando `
        /st $hora `
        /f | Out-Null

    $status.Text = "Agendamento criado."
    [System.Windows.Forms.MessageBox]::Show("Agendamento criado com sucesso.")
})

# =====================================================
# BACKUP
# =====================================================

$btnStart.Add_Click({

    $origem = $txtOrigem.Text.Trim()
    $destino = $txtDestino.Text.Trim()

    if(!(Test-Path $origem)){
        $status.Text = "Origem inválida."
        return
    }

    if(!(Test-Path $destino)){
        $status.Text = "Destino inválido."
        return
    }

    $status.Text = "Executando backup..."
    $progress.Value = 5
    $log.AppendText("Iniciando backup...`r`n")

    robocopy $origem $destino `
        /MIR `
        /R:2 `
        /W:5 `
        /FFT `
        /Z `
        /MT:8 `
        /XA:H `
        /XJ `
        /XD "log" "$RECYCLE.BIN" "System Volume Information" | Out-Null

    $progress.Value = 100
    $status.Text = "Concluído."
    $log.AppendText("Backup concluído com sucesso.`r`n")

    if($chk.Checked){
        shutdown /s /t 0
    }
})

# =====================================================
# EXECUTAR
# =====================================================

[void]$form.ShowDialog()