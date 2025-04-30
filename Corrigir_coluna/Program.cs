using ClosedXML.Excel;
using System.Linq;
using System;


class Program
{
    public static void Main(string[] args)
    {
        var caminhoArquivo = @"Caminho do arquivo";
        try
        {
            using var workbook = new XLWorkbook(caminhoArquivo); // Criar um novo arquivo no Excel
            var planilha = workbook.Worksheet(1); // acessar a primeira aba do Excel

            int linha = 2;
            while (!planilha.Cell(linha,1).IsEmpty())
            {
                var coluna = planilha.Cell(linha, 8).GetString();
                var removeEspaco = coluna.Trim();

                string extrairValor;
                if (removeEspaco.Contains(":"))
                {
                    // se for formato de hora hh:mm:ss, pega só as horas (antes dos ':')
                    extrairValor = removeEspaco.Split(':')[0];
                }
                else
                {
                    extrairValor = string.Concat(removeEspaco.Where(char.IsDigit));
                }
                planilha.Cell(linha,8).Value = int.Parse(extrairValor);
                linha++;
            }

            workbook.SaveAs(caminhoArquivo);
            Console.WriteLine("Limpeza concluída e arquivo salvo com sucesso!");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
   
}