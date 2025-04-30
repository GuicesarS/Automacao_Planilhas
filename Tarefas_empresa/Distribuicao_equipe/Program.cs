using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

public class Dados
{
    public int Ticket { get; set; }
    public string Responsavel { get; set; }
    public Dados(int ticket, string responsavel)
    {
        Ticket = ticket;
        Responsavel = responsavel ?? "User_NOTFOUND";
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string path = @"Caminho do arquivo";
            List<Dados> dados_list = new List<Dados>();
            using (StreamReader sr = File.OpenText(path))
            {
                sr.ReadLine();

                while (!sr.EndOfStream)
                {

                    string line = sr.ReadLine();

                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] parts = line.Split(',');
                    if (parts.Length < 2)
                        continue;

                    if (!int.TryParse(parts[0], out int ticket))
                        continue;

                    string responsavel = string.IsNullOrWhiteSpace(parts[1]) ? "SEM_RESPONSAVEL" : parts[1];
                    dados_list.Add(new Dados(ticket, responsavel));
                }

                string novo_arquivo_resp = @"Caminho do arquivo";

                if (!Directory.Exists(novo_arquivo_resp))
                {
                    Directory.CreateDirectory(novo_arquivo_resp);
                }

                var colaboradoresUnicos = dados_list.Select(d => d.Responsavel).Distinct();

                foreach (var colaborador in colaboradoresUnicos)
                {
                    var dadosColaborador = dados_list.Where(d => d.Responsavel == colaborador).ToList();
                    string nomeArquivo = $"{colaborador.Replace(" ", "_")}.csv";
                    string caminhoCompleto = Path.Combine(novo_arquivo_resp, nomeArquivo);

                    using (StreamWriter sw = new StreamWriter(caminhoCompleto))
                    {
                        sw.WriteLine("ID,RESPONSÁVEL");
                        foreach (var d in dadosColaborador)
                        {
                            sw.WriteLine($"{d.Ticket},{d.Responsavel}");
                        }
                    }

                    Console.WriteLine($"Arquivo gerado para: {colaborador}");
                }
                
                System.Diagnostics.Process.Start("explorer.exe", novo_arquivo_resp);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ocorreu um erro: {ex.Message}");
        }
    }
}