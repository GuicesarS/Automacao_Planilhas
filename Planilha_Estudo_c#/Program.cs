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
        
        // Carregar as horas em um dicionario
        var horas = new Dictionary<string, string>();
        using (var wbHoras = new XLWorkbook(planilha2))
        {
            var wsHoras = wbHoras.Worksheet(1);
                        
            // Percorre as linhas da planilha de horas, começando da segunda linha (ignorando o cabeçalho)
            foreach (var linhas in wsHoras.RowsUsed().Skip(1))
            {
                
                var ticket = linhas.Cell(1).GetValue<string>().Trim();
                    // linhas.Cell(1) → Coluna A (número do ticket)
                    // Trim() → tira espaços em branco antes/depois
                    var tempo = linhas.Cell(3).GetValue<string>().Trim();
                    
                      Console.WriteLine($"[DEBUG] Ticket carregado: {ticket}, Tempo: {tempo}");

                    
                // Verifica se o ticket não está vazio e ainda não existe no dicionário
                if (!string.IsNullOrEmpty(ticket) && !horas.ContainsKey(ticket))
                {
                    // Adiciona o ticket e seu respectivo tempo no dicionário
                    horas[ticket] = tempo;
                    // As minhas horas, na chave ticket recebe o tempo que gastei na demanda.
                    //  "123" => "01:00:00",
                }
            }
        }
        
        // abrir a planilha dos tickets e preencher os campos
        using (var wbTickets = new XLWorkbook(planilha1))
        {
            var wsTickets = wbTickets.Worksheet(1);
            
            // Descobre quantas linhas há
            var ultimaLinha = wsTickets.LastRowUsed().RowNumber();
            
            // A Coluna H é onde vamos adicionar as horas coletadas
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
                 Console.WriteLine($"[DEBUG] Ticket procurando: {ticket}");
            }
            
            // salvar no novo arquivo
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