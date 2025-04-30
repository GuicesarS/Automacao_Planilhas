using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;

record Planilha(
    int Id,
    string Produto,
    decimal Valor,
    DateTime DataVenda,
    string Vendedor
);

class RelatorioPlanilha
{
    public static List<Planilha> GerarDados() => new()
    {
        new(1, "notebook", 3500.00m, DateTime.Now.AddDays(-10), "João da Silva"),
        new(1, "Smartphone", 2000.50m, DateTime.Now.AddDays(-5), "Mariana"),
        new(1, "Tv", 1200.75m, DateTime.Now.AddDays(-2), "Pedro Soares"), 
    };

public static void CriarPlanilha(List<Planilha> planilha, string pastaDestino)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Relatorio Planilha");

        // Criar cabeçalho
        worksheet.Cell(1, 1).Value = "ID";             
        worksheet.Cell(1, 2).Value = "PRODUTO";         
        worksheet.Cell(1, 3).Value = "VALOR";           
        worksheet.Cell(1, 4).Value = "DATA_VENDA";      
        worksheet.Cell(1, 5).Value = "VENDEDOR";       

        // Formatação de cabeçalho
        var cabecalhoRange = worksheet.Range(1, 1, 1, 5);
        cabecalhoRange.Style.Font.Bold = true;     
        cabecalhoRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
        cabecalhoRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // Preencher dados na planilha
        int linha = 2;
        foreach (var item in planilha)
        {
            worksheet.Cell(linha, 1).Value = item.Id;
            worksheet.Cell(linha, 2).Value = item.Produto;
            worksheet.Cell(linha, 3).Value = item.Valor;
            worksheet.Cell(linha, 4).Value = item.DataVenda;
            worksheet.Cell(linha, 5).Value = item.Vendedor;
            linha++;
        }
        // Formatação de Moeda para coluna de Valor
        var valorRange = worksheet.Range(2, 3, planilha.Count + 1, 3);
        valorRange.Style.NumberFormat.Format = "R$ #, ##0.00";

        //Formatação de data para coluna de data
        var dataRange = worksheet.Range(2, 4, planilha.Count + 1, 4);
        dataRange.Style.NumberFormat.Format = "dd/mm/yyyy";

        //Combinar o caminho da pasta com o nome do arquivo
        string caminhoCompleto = Path.Combine(pastaDestino, "Planilha2 VS CODE.xlsx");

        //Salvar arquivo

        workbook.SaveAs(caminhoCompleto);

    }
}
class Program
{
    public static void Main()
    {
        var dadosDaPlanilha = RelatorioPlanilha.GerarDados();

        // Criar o diretório se n existir
        string caminhoPastaDestino = @"Caminho do arquivo";
        Directory.CreateDirectory(caminhoPastaDestino);

        RelatorioPlanilha.CriarPlanilha(dadosDaPlanilha, caminhoPastaDestino);

        Console.WriteLine("Planilha Gerada!");
        Console.WriteLine($"Arquivo: {Path.Combine(caminhoPastaDestino, "Planilha2 VS CODE.xlsx")}");
    }
}
