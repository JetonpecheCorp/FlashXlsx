using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using ClosedXML.Excel;
using FlashXlsx;
using FlashXlsx.Abstractions;
using MiniExcelLibs;

BenchmarkRunner.Run<ExcelBenchmarks>();

[XlsxSerializable]
public class BenchmarkRow
{
     public int Id { get; set; }
     public string Code { get; set; } = string.Empty;
     public string Description { get; set; } = string.Empty;
     public decimal Montant { get; set; }
     public double Taux { get; set; }
     public DateTime DateCreation { get; set; }
     public bool EstActif { get; set; }
}

[MemoryDiagnoser]
[RankColumn]
public class ExcelBenchmarks
{
     private const int RowCount = 100_000;
     private List<BenchmarkRow> _rows = null!;
     private byte[] _xlsxBytes = null!;

     [GlobalSetup]
     public async Task Setup()
     {
          var baseDate = new DateTime(2026, 1, 1, 10, 0, 0);
          _rows = new List<BenchmarkRow>(RowCount);

          for (int i = 1; i <= RowCount; i++)
          {
               _rows.Add(new BenchmarkRow
               {
                    Id = i,
                    Code = $"ART-{i:D6}",
                    Description = $"Ligne de facturation numéro {i}",
                    Montant = i * 12.50m,
                    Taux = 0.20,
                    DateCreation = baseDate.AddMinutes(i),
                    EstActif = (i % 2 == 0)
               });
          }

          // Pré-génération d'un fichier XLSX de 100 000 lignes en mémoire pour les benchmarks de lecture
          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, _rows);
          _xlsxBytes = ms.ToArray();
     }

     // =========================================================================
     // BENCHMARKS D'EXPORT (100 000 LIGNES)
     // =========================================================================

     [Benchmark(Baseline = true, Description = "Write 100k - FlashXlsx (AOT)")]
     public async Task Write_FlashXlsx()
     {
          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, _rows);
     }

     [Benchmark(Description = "Write 100k - MiniExcel")]
     public async Task Write_MiniExcel()
     {
          using var ms = new MemoryStream();
          await ms.SaveAsAsync(_rows);
     }

     [Benchmark(Description = "Write 100k - ClosedXML")]
     public void Write_ClosedXML()
     {
          using var wb = new XLWorkbook();
          var ws = wb.Worksheets.Add("Sheet1");
          ws.Cell(1, 1).InsertTable(_rows);
          using var ms = new MemoryStream();
          wb.SaveAs(ms);
     }

     // =========================================================================
     // BENCHMARKS D'IMPORT (100 000 LIGNES)
     // =========================================================================

     [Benchmark(Description = "Read 100k - FlashXlsx (AOT)")]
     public async Task<int> Read_FlashXlsx()
     {
          using var ms = new MemoryStream(_xlsxBytes);
          int count = 0;
          await foreach (var _ in FlashXlsxSerializer.ReadAsync<BenchmarkRow>(ms))
          {
               count++;
          }
          return count;
     }

     [Benchmark(Description = "Read 100k - MiniExcel")]
     public async Task<int> Read_MiniExcel()
     {
          using var ms = new MemoryStream(_xlsxBytes);
          int count = 0;
          var rows = await ms.QueryAsync<BenchmarkRow>();
          foreach (var _ in rows)
          {
               count++;
          }
          return count;
     }

     [Benchmark(Description = "Read 100k - ClosedXML")]
     public int Read_ClosedXML()
     {
          using var ms = new MemoryStream(_xlsxBytes);
          using var wb = new XLWorkbook(ms);
          var ws = wb.Worksheet(1);
          int count = 0;
          foreach (var row in ws.RowsUsed().Skip(1))
          {
               _ = new BenchmarkRow
               {
                    Id = row.Cell(1).GetValue<int>(),
                    Code = row.Cell(2).GetString(),
                    Description = row.Cell(3).GetString(),
                    Montant = row.Cell(4).GetValue<decimal>(),
                    Taux = row.Cell(5).GetValue<double>(),
                    DateCreation = row.Cell(6).GetDateTime(),
                    EstActif = row.Cell(7).GetBoolean()
               };
               count++;
          }
          return count;
     }
}