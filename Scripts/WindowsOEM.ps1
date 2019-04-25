# Ativar Windows 10 OEM
# Ivo Dias

# Verificar licenciamento
# Cria uma funcao para verificar o licenciamento atual do Windows
function Verificar-Ativacao {
    [CmdletBinding()]
     param(
     [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
     [string]$DNSHostName = $Env:COMPUTERNAME # Utiliza como parametro padrao o hostname do equipamento atual
     )
     process {
        try {
            $wpa = Get-WmiObject SoftwareLicensingProduct -ComputerName $DNSHostName ` # Recebe o licenciamento atual
            -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" `
            -Property LicenseStatus -ErrorAction Stop
        } 
        catch {
            $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
            $wpa = $null 
        }
        $out = New-Object psobject -Property @{
        ComputerName = $DNSHostName;
        Status = [string]::Empty;
        }
        if ($wpa) {
            :outer foreach($item in $wpa) {
            switch ($item.LicenseStatus) {
            0 {$out.Status = "Nao Licenciado"}
            1 {$out.Status = "Licenciado"; break outer}
            2 {$out.Status = "Fora do periodo de carencia"; break outer}
            3 {$out.Status = "Fora do periodo de tolerancia"; break outer}
            4 {$out.Status = "Nao genuino"; break outer}
            5 {$out.Status = "Notificado"; break outer}
            6 {$out.Status = "Extendido"; break outer}
            default {$out.Status = "Unknown value"}
            }
            }
        }  
        else { $out.Status = $status.Message }
        $out
     }
}
# Ativa o Windows
# Localiza a chave enviada pelo fabricante e faz a ativacao
function Ativar-Windows10OEM {
    <#
        .SYNOPSIS 
            Faz a ativacao do Windows 10 OEM
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    Clear-Host   
    # Utiliza os comandos do SLMGR para fazer a ativacao do Windows 10 com a chave de KMS
    $host.ui.RawUI.WindowTitle = "Ativar o Windows OEM" # Coloca um titulo no Terminal
    try {
         # Pega a chave OEM
        Write-Host "Recuperando chave OEM"
        $DPK = powershell "(Get-WmiObject -query ‘select * from SoftwareLicensingService’).OA3xOriginalProductKey" # Recebe a chave OEM
        Write-Host "Carregando os arquivos de licenciamento do Windows"
        cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc # Carrega os arquivos de licencimento
        sleep 10
        Write-Host "Limpando os arquivos antigos de licenciamento"
        cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk # Remove a chave atual
        sleep 10
        Write-Host "Fazendo a ativacao com a chave $DPK"
        cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk $DPK # Carrega a chave OEM
        cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato # Faz a ativacao
        sleep 10
        Clear-Host
        $validacao = Verificar-Ativacao # Verifica se esta atualmente ativo
        if ($validacao.Status -eq "Licenciado") {
            Write-Host "O Windows esta ativo" # Informa que o Windows esta ativo
        }
        else {
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /rearm # Se nao estiver, recarrega o sistema de ativacao
            Write-Host "Reinicie o computador e verifique a ativacao"
        }
    }
    # Caso a ativacao nao seja possivel, retorna a mensagem de erro
    catch {
        $ErrorMessage = $_.Exception.Message # Recebe a mensagem de erro
        Write-Host "Um erro ocorreu ao tentar ativar o Windows"
        Write-Host "Erro: $ErrorMessage" # Exibe ela
    }
}
# Inicia a funcao
Ativar-Windows10OEM