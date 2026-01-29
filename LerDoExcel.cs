using System;
using System.Collections.Generic;
using System.Linq;
using Tricentis.Automation.AutomationInstructions.Configuration;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ToscaCustom.ExcelEngine
{
    /// <summary>
    /// Engine customizada para o Tricentis Tosca (TBox SDK).
    /// Esta classe permite ler dados de uma planilha Excel e transformá-los em Buffers dinamicamente.
    /// Para que o Tosca reconheça esta classe, é necessário:
    /// 1. Herdar de SpecialExecutionTask.
    /// 2. Utilizar o atributo [SpecialExecutionTaskName] com o nome que será usado no XModule.
    /// </summary>
    [SpecialExecutionTaskName("LerDoExcel")]
    public class LerDoExcel : SpecialExecutionTask
    {
        // O construtor deve receber o objeto Validator para processar parâmetros do Tosca.
        public LerDoExcel(Validator validator) : base(validator)
        {
        }

        // Define se o motor de execução deve esperar que as ações anteriores terminem.
        public override bool EnableSynchronization => true;

        /// <summary>
        /// Método principal de execução chamado pelo Tosca Commander.
        /// </summary>
        /// <param name="testAction">Contém os parâmetros de entrada e métodos para definir o resultado do teste.</param>
        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            // Obtendo valores dos parâmetros definidos no XModule.
            // O segundo parâmetro (false/true) indica se o parâmetro é obrigatório.
            var filePath = testAction.GetParameterAsInputValue("Path", false)?.Value;
            var sheetName = testAction.GetParameterAsInputValue("WorkSheet", false)?.Value;
            var tcName = testAction.GetParameterAsInputValue("TC Name", false)?.Value;

            var occurrenceVal = testAction.GetParameterAsInputValue("Occurrence", true)?.Value;
            var actionVal = testAction.GetParameterAsInputValue("Action", true)?.Value;

            if (string.IsNullOrWhiteSpace(filePath) ||
                string.IsNullOrWhiteSpace(sheetName) ||
                string.IsNullOrWhiteSpace(tcName))
            {
                return new UnknownFailedActionResult(
                    "Missing mandatory arguments: Path, WorkSheet, or TC Name.");
            }

            // ===== AJUSTE: Occurrence aceita número OU soma simples (ex.: "1+1", "2 + 1") =====
            int occurrence = 1;
            if (!string.IsNullOrWhiteSpace(occurrenceVal))
            {
                if (!TryParseOccurrence(occurrenceVal, out occurrence))
                {
                    string hint = occurrenceVal.Contains("[")
                        ? " Tente usar {REPETITION} com chaves em vez de colchetes."
                        : " Aceita também soma simples tipo '1+1'.";

                    return new UnknownFailedActionResult(
                        $"Valor inválido para Occurrence: '{occurrenceVal}'. Deve ser um número.{hint}");
                }
            }
            if (occurrence < 1) occurrence = 1;

            string action = string.IsNullOrWhiteSpace(actionVal) ? "Read" : actionVal;

            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    if (workbookPart?.Workbook == null)
                    {
                        return new UnknownFailedActionResult("Invalid workbook.");
                    }

                    var sheet = workbookPart.Workbook.Descendants<Sheet>()
                        .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

                    if (sheet == null)
                    {
                        return new UnknownFailedActionResult(
                            $"Sheet '{sheetName}' not found in '{filePath}'.");
                    }

                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    if (sheetData == null)
                    {
                        return new PassedActionResult("Empty SheetData.");
                    }

                    var rows = sheetData.Elements<Row>().ToList();
                    if (rows.Count == 0)
                    {
                        return new PassedActionResult("Empty Sheet.");
                    }

                    // ---------- LOGICA: COUNT ----------
                    // Se a ação for "Count", apenas contamos as ocorrências do TC Name
                    // e armazenamos em um buffer específico chamado "RowCount".
                    if (action.Equals("Count", StringComparison.OrdinalIgnoreCase))
                    {
                        int totalMatches = 0;

                        for (int i = 1; i < rows.Count; i++) // pula header (Linha 1)
                        {
                            var firstCell = GetCellAt(rows[i], 0);
                            var val = GetCellValue(doc, firstCell);

                            if (!string.IsNullOrWhiteSpace(val) &&
                                val.Trim().Equals(tcName.Trim(), StringComparison.OrdinalIgnoreCase))
                            {
                                totalMatches++;
                            }
                        }

                        // Define o buffer no repositório do Tosca.
                        Buffers.Instance.SetBuffer("RowCount", totalMatches.ToString(), true);
                        return new PassedActionResult(
                            $"Found {totalMatches} rows for '{tcName}'. Buffer 'RowCount' set.");
                    }

                    // ---------- LOGICA: READ (Padrão) ----------
                    // Aqui mapeamos os cabeçalhos das colunas (Linha 1) para criar buffers com seus nomes.
                    var headerRow = rows[0];
                    var headerMap = new Dictionary<int, string>();

                    foreach (Cell cell in headerRow.Elements<Cell>())
                    {
                        int idx = GetColumnIndex(cell.CellReference?.Value);
                        string header = GetCellValue(doc, cell);

                        if (!string.IsNullOrWhiteSpace(header))
                            headerMap[idx] = header.Trim();
                    }

                    Row targetRow = null;
                    int matchCount = 0;

                    for (int i = 1; i < rows.Count; i++)
                    {
                        var firstCell = GetCellAt(rows[i], 0);
                        var val = GetCellValue(doc, firstCell);

                        if (!string.IsNullOrWhiteSpace(val) &&
                            val.Trim().Equals(tcName.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            matchCount++;
                            if (matchCount == occurrence)
                            {
                                targetRow = rows[i];
                                break;
                            }
                        }
                    }

                    if (targetRow == null)
                    {
                        return new UnknownFailedActionResult(
                            $"Test Case '{tcName}' with Occurrence '{occurrence}' not found. Total matches found: {matchCount}.");
                    }

                    // Itera sobre as células da linha encontrada e cria um buffer para cada coluna.
                    var buffersSet = new List<string>();
                    foreach (Cell cell in targetRow.Elements<Cell>())
                    {
                        int colIndex = GetColumnIndex(cell.CellReference?.Value);
                        if (!headerMap.TryGetValue(colIndex, out var bufferName))
                            continue;

                        var bufferValue = GetCellValue(doc, cell);
                        if (!string.IsNullOrWhiteSpace(bufferName))
                        {
                            string valFinal = bufferValue ?? string.Empty;
                            // Buffers.Instance.SetBuffer cria ou atualiza o valor no Buffer Viewer do Tosca.
                            Buffers.Instance.SetBuffer(bufferName, valFinal, true);
                            buffersSet.Add($"{bufferName}='{valFinal}'");
                        }
                    }

                    string resultMsg = $"Dados lidos com sucesso para '{tcName}' (Occurrence {occurrence}).";
                    if (buffersSet.Count > 0)
                        resultMsg += " Buffers definidos: " + string.Join(", ", buffersSet);
                    else
                        resultMsg += " Nenhum buffer foi definido (verifique se os nomes na Linha 1 coincidem com sua planilha).";

                    // Retorna o resultado de sucesso (cor verde no Tosca).
                    return new PassedActionResult(resultMsg);
                }
            }
            catch (Exception ex)
            {
                var msg = ex.Message ?? ex.ToString();
                if (msg.IndexOf("process cannot access the file", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return new UnknownFailedActionResult(
                        "Erro: O arquivo está aberto no Excel. Por favor, feche-o e tente novamente.");
                }

                return new UnknownFailedActionResult($"Erro ao ler o Excel: {msg}");
            }
        }

        // ===== AJUSTE: parse do Occurrence aceita número OU soma simples =====
        private static bool TryParseOccurrence(string input, out int result)
        {
            result = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;

            // 1) número puro
            if (int.TryParse(input.Trim(), out result))
                return true;

            // 2) aceita soma simples: "1+1" / "1 + 1 + 2"
            string s = input.Replace(" ", "");
            if (!s.Contains("+")) return false;

            var parts = s.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) return false;

            int sum = 0;
            foreach (var p in parts)
            {
                if (!int.TryParse(p, out int n))
                    return false;

                sum += n;
            }

            result = sum;
            return true;
        }

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell == null) return string.Empty;

            string value = cell.CellValue?.InnerText ?? cell.InnerText ?? string.Empty;

            if (cell.DataType != null)
            {
                if (cell.DataType.Value == CellValues.SharedString)
                {
                    var sst = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                    if (sst != null && int.TryParse(value, out int idx))
                    {
                        return sst.ChildElements[idx].InnerText ?? string.Empty;
                    }
                }
                else if (cell.DataType.Value == CellValues.Boolean)
                {
                    return value == "0" ? "FALSE" : "TRUE";
                }
            }

            // Se for InlineString
            if (string.IsNullOrEmpty(value) && cell.InlineString != null)
            {
                return cell.InlineString.InnerText ?? string.Empty;
            }

            return value;
        }

        private static Cell GetCellAt(Row row, int columnIndex)
        {
            return row.Elements<Cell>()
                .FirstOrDefault(c => GetColumnIndex(c.CellReference?.Value) == columnIndex);
        }

        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrWhiteSpace(cellReference)) return -1;

            string columnName = string.Concat(cellReference.Where(c => !char.IsDigit(c)));
            int value = 0;

            foreach (char ch in columnName.ToUpperInvariant())
            {
                value *= 26;
                value += (ch - 'A' + 1);
            }

            return value - 1; // 0-indexed
        }
    }
}