Import-Module ActiveDirectory

# === CONFIGURAZIONE ===
$SMTPServer = "smtp.gmail.com"
$EmailMittente = "noreply@tuodominio.it"
$EmailAmministratore = "admin@tuodominio.it"
$GiorniAvviso = 10

# Configurazione credenziali SMTP
$Username = "noreply@tuodominio.it"
$Password = "LatuaPasswordQui" | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)

# === CALCOLO SCADENZA PASSWORD ===
$maxPwdAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
$oggi = Get-Date

# Prova con un filtro più semplice
Write-Host "Ricerca utenti in Active Directory..."

# Prima prova: tutti gli utenti abilitati
$tuttiUtenti = Get-ADUser -Filter "Enabled -eq 'True'" -Properties DisplayName, EmailAddress, PasswordLastSet, PasswordNeverExpires
Write-Host "Utenti abilitati totali: $($tuttiUtenti.Count)"

# Seconda prova: filtra quelli con password che scade
$utentiConScadenza = $tuttiUtenti | Where-Object { $_.PasswordNeverExpires -eq $false }
Write-Host "Utenti con password che scade: $($utentiConScadenza.Count)"

# MODIFICA: Usa TUTTI gli utenti abilitati (anche senza PasswordLastSet)
$utenti = $tuttiUtenti
Write-Host "Utenti da includere nel report: $($utenti.Count)"

# Debug statistiche
$conPasswordNeverExpires = ($utenti | Where-Object { $_.PasswordNeverExpires -eq $true }).Count
$conPasswordLastSetNull = ($utenti | Where-Object { $_.PasswordLastSet -eq $null }).Count
$conPasswordLastSetValido = ($utenti | Where-Object { $_.PasswordLastSet -ne $null -and $_.PasswordNeverExpires -eq $false }).Count

Write-Host "Statistiche utenti:"
Write-Host "- Con PasswordNeverExpires = True: $conPasswordNeverExpires"
Write-Host "- Con PasswordLastSet = null: $conPasswordLastSetNull"
Write-Host "- Con password che scade normalmente: $conPasswordLastSetValido"

$utentiScadenza = @()

foreach ($utente in $utenti) {
    if ($utente.PasswordNeverExpires) {
        # Password non scade mai
        $dataScadenza = $null
        $giorniAllaScadenza = 9999
        $statoPassword = "NON SCADE"
    } elseif ($utente.PasswordLastSet -eq $null -or $utente.PasswordLastSet -eq 0) {
        # Utente non ha mai cambiato password (controllo anche per 0)
        $dataScadenza = $null
        $giorniAllaScadenza = -1
        $statoPassword = "MAI CAMBIATA"
    } else {
        # Password ha scadenza
        $dataScadenza = $utente.PasswordLastSet + $maxPwdAge
        $giorniAllaScadenza = ($dataScadenza - $oggi).Days
        $statoPassword = if ($giorniAllaScadenza -lt 0) { "SCADUTA" }
                        elseif ($giorniAllaScadenza -le $GiorniAvviso) { "IN SCADENZA" }
                        else { "OK" }
    }

    # Aggiungi SEMPRE al report (tutti gli utenti)
    $utentiScadenza += [PSCustomObject]@{
        Nome             = $utente.DisplayName
        Email            = $utente.EmailAddress
        ScadenzaPassword = $dataScadenza
        GiorniAllaScadenza = $giorniAllaScadenza
        StatoPassword    = $statoPassword
        PasswordNeverExpires = $utente.PasswordNeverExpires
        PasswordLastSet  = $utente.PasswordLastSet
    }

    # Invia email solo agli utenti con password in scadenza entro i giorni configurati
    if ($giorniAllaScadenza -le $GiorniAvviso -and $giorniAllaScadenza -ge 0 -and !$utente.PasswordNeverExpires -and $utente.PasswordLastSet -ne $null) {
        if ($utente.EmailAddress) {
            $oggettoUtente = "La tua password scadrà tra $giorniAllaScadenza giorni"
            $corpoUtente = @"
Ciao $($utente.DisplayName),

La tua password scadrà il $($dataScadenza.ToString("dd/MM/yyyy")), cioè tra $giorniAllaScadenza giorni.

Ti consigliamo di cambiarla il prima possibile per evitare interruzioni di accesso.

Grazie,
Il team IT
"@
            Send-MailMessage -From $EmailMittente -To $utente.EmailAddress -Subject $oggettoUtente -Body $corpoUtente -SmtpServer $SMTPServer -Port 587 -UseSsl -Credential $Credential
        }
    }
}

# === ORDINA E CREA REPORT ===
# Ordina: prima quelli in scadenza (giorni crescenti), poi gli altri
$reportOrdinato = $utentiScadenza | Sort-Object @{
    Expression = {
        if ($_.StatoPassword -eq "SCADUTA") { 1 }
        elseif ($_.StatoPassword -eq "IN SCADENZA") { 2 }
        elseif ($_.StatoPassword -eq "OK") { 3 }
        elseif ($_.StatoPassword -eq "MAI CAMBIATA") { 4 }
        else { 5 } # NON SCADE
    }
}, GiorniAllaScadenza

# Crea report HTML con tabella
$corpoReport = @"
<html>
<head>
    <title>Report Scadenze Password</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #2c3e50; text-align: center; }
        .summary { background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background-color: #3498db; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .scaduta { background-color: #e74c3c; color: white; font-weight: bold; }
        .in-scadenza { background-color: #f39c12; color: white; font-weight: bold; }
        .ok { background-color: #27ae60; color: white; }
        .mai-cambiata { background-color: #e67e22; color: white; }
        .non-scade { background-color: #95a5a6; color: white; }
    </style>
</head>
<body>
    <h1>Report Giornaliero - Scadenze Password Utenti</h1>
    <div class="summary">
        <strong>Data generazione:</strong> $($oggi.ToString("dd/MM/yyyy HH:mm"))<br>
        <strong>Totale utenti trovati:</strong> $($reportOrdinato.Count)<br>
        <strong>Utenti con password in scadenza (entro $GiorniAvviso giorni):</strong> $($reportOrdinato | Where-Object { $_.StatoPassword -eq "IN SCADENZA" } | Measure-Object | Select-Object -ExpandProperty Count)<br>
        <strong>Utenti con password scaduta:</strong> $($reportOrdinato | Where-Object { $_.StatoPassword -eq "SCADUTA" } | Measure-Object | Select-Object -ExpandProperty Count)
    </div>
    
    <table>
        <thead>
            <tr>
                <th>Utente</th>
                <th>Email</th>
                <th>Stato Password</th>
                <th>Data Scadenza</th>
                <th>Giorni Rimanenti</th>
            </tr>
        </thead>
        <tbody>
"@

foreach ($riga in $reportOrdinato) {
    $classeCSS = switch ($riga.StatoPassword) {
        "SCADUTA" { "scaduta" }
        "IN SCADENZA" { "in-scadenza" }
        "OK" { "ok" }
        "MAI CAMBIATA" { "mai-cambiata" }
        "NON SCADE" { "non-scade" }
        default { "" }
    }
    
    $dataScadenzaStr = if ($riga.ScadenzaPassword) { 
        $riga.ScadenzaPassword.ToString("dd/MM/yyyy") 
    } else { 
        "-" 
    }
    
    $giorniStr = if ($riga.GiorniAllaScadenza -eq 9999) { 
        "∞" 
    } elseif ($riga.GiorniAllaScadenza -eq -1) { 
        "-" 
    } else { 
        $riga.GiorniAllaScadenza.ToString() 
    }
    
    $corpoReport += @"
            <tr>
                <td>$($riga.Nome)</td>
                <td>$($riga.Email)</td>
                <td class="$classeCSS">$($riga.StatoPassword)</td>
                <td>$dataScadenzaStr</td>
                <td>$giorniStr</td>
            </tr>
"@
}

$corpoReport += @"
        </tbody>
    </table>
</body>
</html>
"@

# === INVIA REPORT ALL'AMMINISTRATORE ===
Write-Host "Invio report all'amministratore con $($reportOrdinato.Count) utenti..."
try {
    Send-MailMessage -From $EmailMittente -To $EmailAmministratore -Subject "Report giornaliero scadenze password" -Body $corpoReport -SmtpServer $SMTPServer -Port 587 -UseSsl -Credential $Credential -BodyAsHtml
    Write-Host "Report inviato con successo!"
} catch {
    Write-Host "Errore nell'invio del report: $($_.Exception.Message)" -ForegroundColor Red
}
