using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;

class Program
{
    public static void Main(string[] args)
    {
        try
        {
            var planilha1 = @"Caminho do arquivo";
            var planilha2 = @"Caminho do arquivo";
            var planilhaFinal = @"Caminho do arquivo";

            // Dicionário com ticket e tempo (em horas inteiras)
            var horas = new Dictionary<string, string>();

            using (var wbHoras = new XLWorkbook(planilha2))
            {
                var wsHoras = wbHoras.Worksheet(1);

                foreach (var linha in wsHoras.RowsUsed().Skip(1))
                {
                    var ticket = linha.Cell(1).GetValue<string>().Trim();
                    var tempoStr = linha.Cell(8).GetValue<string>().Trim();

                    if (string.IsNullOrWhiteSpace(ticket) || string.IsNullOrWhiteSpace(tempoStr))
                        continue;

                    if (TimeSpan.TryParse(tempoStr, out var tempoConvertido))
                    {
                        int horasInteiras = (int)Math.Round(tempoConvertido.TotalHours); // arredonda para cima ou baixo
                        horas[ticket] = horasInteiras;
                    }
                }
            }

            using (var wbTickets = new XLWorkbook(planilha1))
            {
                var wsTickets = wbTickets.Worksheet("Aba da planilha");
                var ultimaLinha = wsTickets.LastRowUsed().RowNumber();
                var destacar = XLColor.LightYellow;

                for (int linha = 2; linha <= ultimaLinha; linha++)
                {
                    var ticket = wsTickets.Cell(linha, 1).GetString().Trim();
                    if (horas.TryGetValue(ticket, out var tempo))
                    {
                        var cellHoras = wsTickets.Cell(linha, 8);
                        cellHoras.Value = tempo;
                        cellHoras.Style.Fill.SetBackgroundColor(destacar);
                    }
                }

                wbTickets.SaveAs(planilhaFinal);
            }

            Console.WriteLine("Planilha alterada com sucesso");
            Console.WriteLine($"Caminho: {planilhaFinal}");
            Console.ReadKey();
        }
        catch (Exception e)
        {
            Console.WriteLine($"Erro: {e.Message}");
        }
    }
}
