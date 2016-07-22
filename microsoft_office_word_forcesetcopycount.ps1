If (Test-Path  -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0'){
        Write-Host "Office 2013 encontrato, criando chave no registro..."
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Options' -Name ForceSetCopyCount -Type DWord -Value 1
}
ElseIf (Test-Path -Path 'HKCU:\SOFTWARE\Microsoft\Office\15.0'){
        Write-Host "Office 2010 encontrato, criando chave no registro..."
        Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\15.0\Word\Options' -Name ForceSetCopyCount -Type DWord -Value 1
}
ElseIf (Test-Path -Path 'HKCU:\SOFTWARE\Microsoft\Office\14.0'){
        Write-Host "Office 2010 encontrato, criando chave no registro..."
        Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\14.0\Word\Options' -Name ForceSetCopyCount -Type DWord -Value 1
}
ElseIf (Test-Path -Path 'HKCU:\SOFTWARE\Microsoft\Office\13.0'){
        Write-Host "Office 2007 encontrato, criando chave no registro..."
        Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\13.0\Word\Options' -Name ForceSetCopyCount -Type DWord -Value 1   
}
ElseIf (Test-Path -Path 'HKCU:\SOFTWARE\Microsoft\Office\12.0'){
        Write-Host "Office 2003 encontrato, criando chave no registro..."
        Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\13.0\Word\Options' -Name ForceSetCopyCount -Type DWord -Value 1   
}
ElseIf (Test-Path -Path 'HKCU:\SOFTWARE\Microsoft\Office\11.0'){
        Write-Host "Office 2003 encontrato, criando chave no registro..."
        Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\13.0\Word\Options' -Name ForceSetCopyCount -Type DWord -Value 1   
}
ElseIf (Test-Path -Path 'HKCU:\SOFTWARE\Microsoft\Office\10.0'){
        Write-Host "Office 2003 encontrato, criando chave no registro..."
        Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\13.0\Word\Options' -Name ForceSetCopyCount -Type DWord -Value 1   
}
Else{
    Write-Host "Nenhum Office encontrado!"
}
