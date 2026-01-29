# Script para copiar DLL do Engine LerDoExcel para o Tosca
# IMPORTANTE: Execute este script como Administrador

$sourceDll = "c:\Users\rcorrear\Documents\repositorio local\EngineTosca2\bin\Release\EngineTosca2.dll"
$destinationFolder = "C:\Program Files (x86)\TRICENTIS\Tosca Testsuite\TBox\"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Copiando Engine LerDoExcel para Tosca" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "NOTA: DocumentFormat.OpenXml.dll NÃO precisa ser copiado." -ForegroundColor Yellow
Write-Host "      O Tosca já possui a versão 3.1.1 instalada." -ForegroundColor Yellow
Write-Host ""

Write-Host "Copiando EngineTosca2.dll..." -ForegroundColor White

try {
    Copy-Item $sourceDll -Destination $destinationFolder -Force
    Write-Host "✓ DLL copiada com sucesso!" -ForegroundColor Green
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "  PRÓXIMOS PASSOS OBRIGATÓRIOS:" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "1. Reinicie o Tosca Commander" -ForegroundColor White
    Write-Host "2. Crie um XModule (ver GUIA_CONFIGURACAO_TOSCA.md)" -ForegroundColor White
    Write-Host "   - Engine = LerDoExcel" -ForegroundColor White
    Write-Host "   - SpecialExecutionTask = LerDoExcel" -ForegroundColor White
    Write-Host ""
    Write-Host "Arquivo copiado:" -ForegroundColor Cyan
    Get-ChildItem "$destinationFolder\EngineTosca2.dll" | Select-Object Name, Length, LastWriteTime | Format-Table
}
catch {
    Write-Host "✗ Erro ao copiar DLL: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "SOLUÇÃO:" -ForegroundColor Yellow
    Write-Host "Execute o PowerShell como Administrador e rode este script novamente." -ForegroundColor White
}
