# Guia de Configuração do XModule no Tosca Commander

## ⚠️ IMPORTANTE: Configuração Obrigatória

Para usar o engine customizado `LerDoExcel`, você **DEVE** criar um XModule no Tosca Commander. Simplesmente copiar a DLL não é suficiente!

---

## 📋 Passo a Passo

### 1. Copiar a DLL para o Tosca

Execute o PowerShell como **Administrador** e rode:
```powershell
cd "c:\Users\rcorrear\Documents\repositorio local\EngineTosca2"
.\CopiarParaTosca.ps1
```

### 2. Reiniciar o Tosca Commander

Feche completamente e reabra o Tosca Commander.

### 3. Criar um XModule

1. No **Tosca Commander**, vá para a seção **Modules**
2. Clique com botão direito → **Create** → **XModule**
3. Nomeie o módulo (ex: `LerDoExcel`)

### 4. Configurar o XModule

#### 4.1 Configurar o Engine
1. Selecione o XModule criado
2. Na aba **Details** (lado direito), adicione uma **Configuration**:
   - **Name**: `Engine`
   - **Value**: `LerDoExcel`

#### 4.2 Configurar o SpecialExecutionTask
1. Ainda nas **Configurations**, adicione outra:
   - **Name**: `SpecialExecutionTask`
   - **Value**: `LerDoExcel`

### 5. Adicionar Parâmetros (Module Attributes)

Clique com botão direito no XModule → **Create** → **Module Attribute** para cada parâmetro:

#### Parâmetros Obrigatórios:

**1. Path**
- Name: `Path`
- Configuration → **Parameter**: `true`
- Configuration → **ActionMode**: `Input`

**2. WorkSheet**
- Name: `WorkSheet`
- Configuration → **Parameter**: `true`
- Configuration → **ActionMode**: `Input`

**3. TC Name**
- Name: `TC Name`
- Configuration → **Parameter**: `true`
- Configuration → **ActionMode**: `Input`

#### Parâmetros Opcionais:

**4. Occurrence** (opcional)
- Name: `Occurrence`
- Configuration → **Parameter**: `true`
- Configuration → **ActionMode**: `Input`
- DefaultValue: `1`

**5. Action** (opcional)
- Name: `Action`
- Configuration → **Parameter**: `true`
- Configuration → **ActionMode**: `Input`
- DefaultValue: `Read`

---

## 🧪 Como Usar no Test Case

1. No seu **TestCase**, adicione um **TestStep**
2. Arraste o XModule `LerDoExcel` para o TestStep
3. Preencha os parâmetros:
   - **Path**: Caminho completo do arquivo Excel (ex: `C:\Dados\MeuArquivo.xlsx`)
   - **WorkSheet**: Nome da planilha (ex: `Sheet1`)
   - **TC Name**: Nome do caso de teste a buscar (ex: `TC001`)
   - **Occurrence** (opcional): Qual ocorrência usar se houver múltiplas linhas (padrão: `1`)
   - **Action** (opcional): `Read` ou `Count` (padrão: `Read`)

---

## 📊 Funcionalidades

### Modo Read (Padrão)
- Lê os dados da linha correspondente ao TC Name
- Popula buffers do Tosca com os valores das colunas
- Os nomes dos buffers são os cabeçalhos da planilha

### Modo Count
- Conta quantas linhas correspondem ao TC Name
- Define o buffer `RowCount` com o total

---

## ❌ Solução de Problemas

Se ainda aparecer o erro "Engine 'LerDoExcel' is not valid":

1. ✅ Verifique se a DLL está em: `C:\Program Files (x86)\TRICENTIS\Tosca Testsuite\TBox\EngineTosca2.dll`
2. ✅ Confirme que o Tosca foi reiniciado após copiar a DLL
3. ✅ Verifique se o XModule tem as configurações `Engine` e `SpecialExecutionTask` com valor `LerDoExcel`
4. ✅ Certifique-se de que todos os Module Attributes têm `Parameter=true`

---

## 💡 Casos de Uso e Exemplos Práticos

Agora que seu engine está funcionando, você pode usar os parâmetros opcionais `Occurrence` e `Action` para fluxos mais complexos.

### 1. Parâmetro `Occurrence`
**Cenário**: Você tem o mesmo `TC Name` várias vezes na planilha (ex: para testar diferentes conjuntos de dados para o mesmo teste "Login").

| TC Name | Usuario | Senha |
|---------|---------|-------|
| Login   | user_A  | 123   |
| Login   | user_B  | 456   |
| Login   | user_C  | 789   |

- **Se preencher `Occurrence` com `1`**: O engine lerá `user_A`.
- **Se preencher `Occurrence` com `2`**: O engine lerá `user_B`.

**Configuração no Tosca**:
- Path: `{CP[Path]}`
- WorkSheet: `Dados`
- TC Name: `Login`
- **Occurrence**: `2`

---

### 2. Parâmetro `Action`
**Cenário**: Você quer saber quantas linhas de dados existem para um teste antes de começar a rodar.

- **Se `Action` estiver vazio**: Ele assume `Read` (lê os dados e cria buffers).
- **Se `Action` for `Count`**: Ele ignora a leitura de dados e apenas conta quantas vezes o `TC Name` aparece. O resultado é salvo no buffer automático `{B[RowCount]}`.

**Configuração no Tosca**:
- Path: `{CP[Path]}`
- WorkSheet: `Dados`
- TC Name: `Login`
- **Action**: `Count`

**Resultado**: O Tosca criará um buffer chamado `RowCount` com o valor `3` (baseado na tabela acima).

---

### 4. Uso com `Repetition` (Loop)

Esta é a forma mais poderosa de usar o engine. Se você tem 5 linhas para o mesmo teste e quer rodar o Tosca 5 vezes (uma para cada linha), siga este padrão:

1. **Passo 1: Contar as linhas**
   - No Tosca, crie um TestStep da engine.
   - Configure: `Action` = `Count`.
   - Isso criará o buffer `{B[RowCount]}`.

2. **Passo 2: Configurar a Repetição**
   - No Folder ou no TestStep que você quer repetir, vá em **Properties**.
   - No campo **Repetition**, coloque: `{B[RowCount]}`.

3. **Passo 3: Ler os dados dinamicamente**
   - Dentro do loop, coloque outro TestStep da engine (agora para ler).
   - Configure: `Action` = `Read` (ou deixe vazio).
   - Configure: `Occurrence` = `{REPETITION:1}`.

> [!TIP]
> **DICA DE OURO**: Use `{REPETITION:1}` com o sinal de dois pontos e o número 1. Isso diz ao Tosca: "Use o número da repetição atual, mas se não houver um loop ativo, use 1". Isso evita o erro **"No parent Repetition found"** se você tentar rodar o passo sozinho para testar.

### ❓ Erro: "No parent Repetition found"
Se este erro aparecer, significa que o folder **Fluxo** não está com a repetição ativa. Verifique:
1. Clique no folder **Fluxo**.
2. Vá na aba **Properties** (fica ao lado de Details/TestConfiguration).
3. Verifique se o valor `{B[RowCount]}` está realmente no campo **Repetition** (e não enganado em outro lugar).
4. Certifique-se de que o folder **Fluxo** está com o ícone de repetição (uma setinha em círculo).

---

### 6. Uso Avançado: While Loop com Fluxo Completo (Pre/Flow/Pos)

Se você precisa de um controle mais granular ou quer executar **Pré-Condições** e **Pós-Condições** para *cada* linha de dados (ex: Abrir o sistema, testar, e fechar o sistema para cada usuário), a estrutura de `Repetition` simples pode não ser suficiente. Nesses casos, use um **While Loop**.

**Estrutura Recomendada no Tosca:**

1.  **TestCase** (ou Folder Principal)
    *   **[Step] Inicializar**: `TBox Set Buffer` -> `Iterator` = `1`
    *   **[Step] Contar Linhas**: Engine `LerDoExcel` -> `Action` = `Count`
    *   **[Loop] While Statement**:
        *   **Condition**: `{B[Iterator]} <= {B[RowCount]}`
        *   **[Step] Ler Dados da Linha Atual**:
            *   Engine `LerDoExcel` -> `Action` = `Read`
            *   `Occurrence` = `{B[Iterator]}`
        *   **[Folder] Pré-Condição**: (Ex: Login, Navegar para Home, Resetar Estado)
        *   **[Folder] Fluxo Principal**: (Seus passos de teste usando os buffers `{B[Usuario]}`, `{B[Senha]}`, etc.)
        *   **[Folder] Pós-Condição**: (Ex: Logout, Fechar Janela, Voltar para Busca)
        *   **[Step] Incrementar**: `TBox Set Buffer` -> `Iterator` = `{MATH[{B[Iterator]} + 1]}`

**Por que usar assim?**
*   **Controle Total**: Você garante que o pré/pós teste roda para cada linha de dados.
*   **Recuperação**: Se um teste falhar no meio, a Pós-Condição ainda pode tentar limpar o ambiente para a próxima iteração (se configurado corretamente o Recovery).
*   **Depuração**: Você pode alterar o valor do buffer `Iterator` manualmente para pular linhas ou re-testar uma linha específica.

---

### 5. Exemplo de Preenchimento (Sintaxe)

Você pode preencher os campos usando valores fixos ou referências do Tosca:

| Parâmetro | Exemplo de Valor | Explicação |
|-----------|------------------|------------|
| **Path** | `{CP[CaminhoExcel]}` | Pega de um Configuration Parameter |
| **WorkSheet** | `CalculoFrete` | Nome fixo da aba |
| **TC Name** | `{TestCase.Name}` | Pega o nome do Test Case atual do Tosca |
| **Occurrence**| `1` ou `1+1` | Pega a ocorrência (aceita soma simples) |
| **Action** | `Read` | Lê os dados (Comportamento padrão) |

> [!TIP]
> Você pode deixar `Occurrence` e `Action` vazios no TestStep se quiser apenas a primeira linha e a leitura padrão. O engine cuidará disso automaticamente.
