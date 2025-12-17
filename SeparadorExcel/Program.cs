using System;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    static void Main(string[] args)
    {
        string arquivoOriginal = @"C:\Fiotec\dados.xlsx";
        string arquivoNovo = @"C:\Fiotec\dados_separados.xlsx";

        using var workbook = new XLWorkbook(arquivoOriginal);
        var ws = workbook.Worksheet("Dados");
        var range = ws.RangeUsed();

        // Descobre a posição da coluna "Unidade organizacional"
        int unidadeCol = range.FirstRow().Cells()
            .Select((c, i) => new { c, i })
            .First(x => x.c.Value.ToString().Trim() == "Unidade organizacional").i + 1;

        // Pega todas as unidades distintas
        var unidades = range.Rows()
            .Skip(1) // pula cabeçalho
            .Select(r => r.Cell(unidadeCol).GetString().Trim())
            .Where(u => !string.IsNullOrWhiteSpace(u))
            .Distinct();

        foreach (var unidade in unidades)
        {
            // Gera nome único para aba
            string nomeAba = GetUniqueSheetName(workbook, unidade);

            var novaWs = workbook.Worksheets.Add(nomeAba);

            // Copia cabeçalho
            int col = 1;
            foreach (var cell in range.FirstRow().Cells())
            {
                novaWs.Cell(1, col).Value = cell.Value;
                col++;
            }

            int linhaDestino = 2;

            // Copia todas as linhas dessa unidade
            var linhas = range.Rows()
                .Skip(1)
                .Where(r => r.Cell(unidadeCol).GetString().Trim() == unidade);

            foreach (var linha in linhas)
            {
                int colIndex = 1;
                foreach (var cell in linha.Cells())
                {
                    novaWs.Cell(linhaDestino, colIndex).Value = cell.Value;
                    colIndex++;
                }
                linhaDestino++;
            }

            novaWs.Columns().AdjustToContents();
        }

        workbook.SaveAs(arquivoNovo);
        Console.WriteLine("✅ Arquivo gerado com 1 aba por unidade organizacional!");
    }

    // Garante nomes válidos de aba (máx. 31 chars, sem caracteres inválidos)
    static string SanitizeSheetName(string name)
    {
        var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        foreach (var ch in invalid) name = name.Replace(ch.ToString(), "-");
        name = name.Length > 31 ? name.Substring(0, 31) : name;
        return string.IsNullOrWhiteSpace(name) ? "Unidade" : name;
    }

    // Gera nome único com sufixo numérico se já existir
    static string GetUniqueSheetName(XLWorkbook workbook, string unidade)
    {
        string baseName = SanitizeSheetName(unidade);
        string nome = baseName;
        int contador = 1;

        while (workbook.Worksheets.Any(ws => ws.Name == nome))
        {
            string sufixo = "_" + contador;
            if (baseName.Length + sufixo.Length > 31)
                nome = baseName.Substring(0, 31 - sufixo.Length);
            else
                nome = baseName;

            nome += sufixo;
            contador++;
        }

        return nome;
    }
}
