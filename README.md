# Tosca-Custom-Excel-Engine: SpecialExecutionTask para Automação TBox

Este projeto é um **template educacional** para desenvolvedores de QA que desejam criar extensões customizadas (SpecialExecutionTasks) para o **Tricentis Tosca** utilizando o TBox SDK e C#.

O objetivo principal desta engine é permitir a leitura dinâmica de dados de planilhas Excel e sua transformação automática em **Buffers** do Tosca, facilitando automações complexas baseadas em dados (Data-Driven Testing).

## 🚀 Funcionalidades

- **Contagem Dinâmica**: Conta quantas ocorrências de um Caso de Teste existem na planilha.
- **Leitura Baseada em Ocorrência**: Permite selecionar qual linha de dados ler através do parâmetro `Occurrence` (útil para loops).
- **Mapeamento Automático**: Transforma o cabeçalho das colunas (Linha 1) em nomes de Buffers automaticamente.
- **Feedback Visual**: Mensagens detalhadas no ScratchBook indicando quais dados foram lidos.

## 📋 Pré-requisitos

- **Tricentis Tosca** (Testado na versão 2024.2).
- **.NET SDK 6.0+** (O projeto compila para .NET Framework 4.8, compatível com o Tosca).
- Referência ao pacote Nuget `DocumentFormat.OpenXml` (v3.1.1).

## 🛠️ Instalação e Build

1. **Clone o repositório**:
   ```bash
   git clone https://github.com/seu-usuario/Tosca-Excel-Engine.git
   ```

2. **Ajuste as Referências**:
   No arquivo `EngineTosca2.csproj`, verifique se os caminhos das DLLs do Tosca apontam para sua pasta de instalação (padrão: `C:\Program Files (x86)\TRICENTIS\Tosca Testsuite\TBox\`).

3. **Compilação**:
   Abra o terminal na pasta do projeto e execute:
   ```bash
   dotnet build -c Release
   ```

4. **Instalação no Tosca**:
   Copie a DLL gerada em `bin/Release/EngineTosca2.dll` para a pasta TBox do seu Tosca:
   `C:\Program Files (x86)\TRICENTIS\Tosca Testsuite\TBox\`
   *(Nota: Você pode usar o script `CopiarParaTosca.ps1` incluso no projeto rodando como Administrador).*

## ⚙️ Configuração no Tosca Commander

### 1. Criar o XModule
- Crie um novo XModule chamado `LerDoExcel`.
- Nas **Properties**, adicione:
  - `Configuration -> Engine`: `LerDoExcel`
  - `Configuration -> SpecialExecutionTask`: `LerDoExcel`

### 2. Adicionar Parâmetros (Module Attributes)
Crie os seguintes atributos e marque-os como `Configuration -> Parameter = true`:
- `Path`: Caminho do arquivo .xlsx.
- `WorkSheet`: Nome da aba da planilha.
- `TC Name`: Identificador (Coluna A) que o engine deve buscar.
- `Occurrence`: (Opcional) Número da ocorrência (Ex: `{REPETITION:1}`).
- `Action`: (Opcional) Use `Count` para apenas contar linhas ou deixe vazio para ler.

## 📖 Como Usar (Exemplo Prático)

Suponha uma planilha com:
| TC Name | URL | Usuario |
| :--- | :--- | :--- |
| Login_Teste | google.com | admin |

No Tosca, ao rodar o TestStep:
1. Ele criará um buffer `{B[URL]}` com valor `google.com`.
2. Ele criará um buffer `{B[Usuario]}` com valor `admin`.

---

## 🤝 Contribuição
Sinta-se à vontade para abrir Issues ou enviar Pull Requests com melhorias. Este projeto visa ajudar a comunidade de QA a dominar as extensões do Tosca.

---
**Autor**: [rbtepss](https://github.com/rbtepss)
**Licença**: MIT
