using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AnsReportController : ControllerBase
    {
        private readonly string _connectionString;
        private const string BrandColor = "#56723e";

        public AnsReportController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        // Norma legala citita din DB - fara hardcoding in C#
        private Dictionary<int, decimal> LoadNorme(SqlConnection conn, int idAnUniv)
        {
            var d = new Dictionary<int, decimal>();
            using var cmd = new SqlCommand(
                @"SELECT ID_TipGradDidactic, NrOreConventionaleTitular
                  FROM [AGSIS].[pi].[NormaOreConventionale]
                  WHERE ID_AnUniv = @id AND NrOreConventionaleTitular > 0", conn);
            cmd.Parameters.AddWithValue("@id", idAnUniv);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                d[Convert.ToInt32(r[0])] = Convert.ToDecimal(r[1]);
            return d;
        }

        // Titulari cu ore efective: TitularAnUniv=1 + TitularSauSuplinitor=1 + NrOreConventionale>0
        // Filtru dual-mode: accepta ID numeric SAU string denumire pentru facultate/departament
        private const string SqlTitulari = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                MIN(p.NumeIntreg)                                    AS NumeIntreg,
                MIN(p.DenumireFacultate)                             AS Facultate,
                MIN(p.DenumireCatedra)                               AS Departament,
                MIN(p.DenumireGradDidactic)                          AS GradDidactic,
                MIN(p.ID_TipGradDidacticAnUniv)                      AS ID_TipGrad
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                ON ppm.ID_Profesor = p.ID_Profesor
                AND ppm.ID_AnUniv = @ID_AnUniv
            WHERE p.ID_AnUnivCatedra = @ID_AnUniv
                AND p.TitularAnUniv = 1
                AND ppm.TitularSauSuplinitor = 1
                AND ppm.NrOreConventionale > 0
                AND (@fac = N'Toti'
                     OR p.ID_Facultate = TRY_CAST(@fac AS INT)
                     OR p.DenumireFacultate COLLATE DATABASE_DEFAULT = @fac COLLATE DATABASE_DEFAULT)
                AND (@dept = N'Toti'
                     OR p.ID_Catedra = TRY_CAST(@dept AS INT)
                     OR p.DenumireCatedra COLLATE DATABASE_DEFAULT = @dept COLLATE DATABASE_DEFAULT)
            GROUP BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))
            HAVING MIN(p.CNP) IS NOT NULL OR COUNT(DISTINCT p.ID_Profesor) = 1
            ORDER BY MIN(p.NumeIntreg)";

        // Ore per domeniu ANS cu JOIN pe ID_PlanMaterie_Prestator pentru deduplicare corecta
        // si JOIN pe DenumireSpecializare pentru domeniu ANS (singurul JOIN valid din DB)
        private const string SqlOreAns = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                fds.ID_N_Domeniu_Studiu_ANS                          AS ID_ANS,
                dc.ID_Cuplaj,
                ppm.NrOreConventionale                               AS OreConventionale
            FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
            INNER JOIN [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                ON ppm.ID_Profesor = p.ID_Profesor
                AND p.ID_AnUnivCatedra = @ID_AnUniv
            LEFT JOIN [AGSIS].[pi].[View_DetaliereCuplaje] dc
                ON ppm.ID_PlanMaterie_Prestator = dc.ID_PlanMaterie_Prestator
                AND dc.ID_AnUniv = @ID_AnUniv
            INNER JOIN [AGSIS].[dbo].[View_FDS] fds
                ON ppm.DenumireSpecializare COLLATE DATABASE_DEFAULT
                 = fds.DenumireSpecializare COLLATE DATABASE_DEFAULT
                AND fds.ID_AnUniv = @ID_AnUniv
            WHERE ppm.ID_AnUniv = @ID_AnUniv
                AND p.TitularAnUniv = 1
                AND ppm.TitularSauSuplinitor = 1
                AND ppm.NrOreConventionale > 0
                AND fds.ID_N_Domeniu_Studiu_ANS IS NOT NULL
                AND (@fac = N'Toti'
                     OR p.ID_Facultate = TRY_CAST(@fac AS INT)
                     OR p.DenumireFacultate COLLATE DATABASE_DEFAULT = @fac COLLATE DATABASE_DEFAULT)
                AND (@dept = N'Toti'
                     OR p.ID_Catedra = TRY_CAST(@dept AS INT)
                     OR p.DenumireCatedra COLLATE DATABASE_DEFAULT = @dept COLLATE DATABASE_DEFAULT)";

        private (Dictionary<string, ProfAns> titulari,
                 Dictionary<string, Dictionary<int, decimal>> fractiuni,
                 Dictionary<int, decimal> norme)
            LoadData(int idAnUniv, string fac, string dept)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            var norme = LoadNorme(conn, idAnUniv);

            var titulari = new Dictionary<string, ProfAns>();
            using (var cmd = new SqlCommand(SqlTitulari, conn))
            {
                cmd.CommandTimeout = 60;
                cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                cmd.Parameters.AddWithValue("@fac", fac);
                cmd.Parameters.AddWithValue("@dept", dept);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    var cnp = r["IdentificatorUnic"].ToString()!;
                    if (string.IsNullOrWhiteSpace(cnp)) continue;
                    titulari[cnp] = new ProfAns
                    {
                        NumeIntreg = r["NumeIntreg"].ToString()!,
                        Facultate = r["Facultate"].ToString()!,
                        Departament = r["Departament"].ToString()!,
                        GradDidactic = r["GradDidactic"].ToString()!,
                        ID_TipGrad = r["ID_TipGrad"] == DBNull.Value
                            ? null : (int?)Convert.ToInt32(r["ID_TipGrad"])
                    };
                }
            }

            var oreRaw = new Dictionary<string, Dictionary<int, decimal>>();
            var cuplajeVazute = new Dictionary<string, HashSet<long>>();
            using (var cmd2 = new SqlCommand(SqlOreAns, conn))
            {
                cmd2.CommandTimeout = 120;
                cmd2.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                cmd2.Parameters.AddWithValue("@fac", fac);
                cmd2.Parameters.AddWithValue("@dept", dept);
                using var r2 = cmd2.ExecuteReader();
                while (r2.Read())
                {
                    var cnp = r2["IdentificatorUnic"].ToString()!;
                    if (string.IsNullOrWhiteSpace(cnp) || !titulari.ContainsKey(cnp)) continue;
                    int idAns = Convert.ToInt32(r2["ID_ANS"]);
                    decimal ore = Convert.ToDecimal(r2["OreConventionale"]);
                    long? cup = r2["ID_Cuplaj"] == DBNull.Value
                        ? null : (long?)Convert.ToInt64(r2["ID_Cuplaj"]);

                    var dk = cnp + "|" + idAns;
                    if (!cuplajeVazute.ContainsKey(dk)) cuplajeVazute[dk] = new HashSet<long>();
                    if (cup.HasValue && !cuplajeVazute[dk].Add(cup.Value)) continue;

                    if (!oreRaw.ContainsKey(cnp)) oreRaw[cnp] = new Dictionary<int, decimal>();
                    oreRaw[cnp].TryAdd(idAns, 0m);
                    oreRaw[cnp][idAns] += ore;
                }
            }

            var fractiuni = new Dictionary<string, Dictionary<int, decimal>>();
            foreach (var (cnp, orePerAns) in oreRaw)
            {
                if (!titulari.TryGetValue(cnp, out var prof)) continue;
                decimal norma = prof.ID_TipGrad.HasValue &&
                                norme.TryGetValue(prof.ID_TipGrad.Value, out var n) ? n : 15m;
                decimal totalOre = orePerAns.Values.Sum();
                decimal baza = Math.Max(totalOre, norma);
                if (totalOre <= 0 || baza <= 0) continue;
                fractiuni[cnp] = new();
                foreach (var kv in orePerAns)
                {
                    decimal frac = Math.Round(kv.Value / baza, 4);
                    if (frac > 0) fractiuni[cnp][kv.Key] = frac;
                }
            }
            return (titulari, fractiuni, norme);
        }

        [HttpGet("date")]
        public IActionResult GetDateAns([FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null, [FromQuery] string? departament = null)
        {
            var fac = string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim();
            var dept = string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim();
            var (titulari, fractiuni, norme) = LoadData(idAnUniv, fac, dept);
            int nrCrt = 1;
            var result = titulari.OrderBy(kv => kv.Value.NumeIntreg).Select(kv =>
            {
                var (cnp, prof) = kv;
                decimal norma = prof.ID_TipGrad.HasValue &&
                                norme.TryGetValue(prof.ID_TipGrad.Value, out var n) ? n : 15m;
                fractiuni.TryGetValue(cnp, out var frac);
                decimal totalOre = frac?.Values.Sum() ?? 0m;
                return new
                {
                    NrCrt = nrCrt++,
                    prof.NumeIntreg,
                    prof.Facultate,
                    prof.Departament,
                    Grad = prof.GradDidactic,
                    NormaLegalaOre = norma,
                    TotalOreConv = Math.Round(totalOre, 2),
                    FractiuniPerDomeniu = frac ?? new Dictionary<int, decimal>()
                };
            }).ToList();
            return Ok(result);
        }

        [HttpGet("export")]
        public IActionResult ExportAns([FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null, [FromQuery] string? departament = null)
        {
            var fac = string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim();
            var dept = string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim();
            var (titulari, fractiuni, norme) = LoadData(idAnUniv, fac, dept);
            var profesori = titulari.OrderBy(kv => kv.Value.NumeIntreg).ToList();

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("CD DRU");
            var green = XLColor.FromHtml(BrandColor);

            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, 50).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov";
            ws.Range(3, 1, 3, 6).Merge();
            ws.Cell(5, 1).Value = "Nr.\nCrt.";
            ws.Cell(5, 2).Value = "Nume si prenume cadru didactic";
            ws.Cell(5, 3).Value = "CNP";
            ws.Cell(5, 4).Value = "Functie cadru didactic sau cercetare";
            ws.Cell(5, 5).Value = "Forma de angajare";
            ws.Cell(5, 6).Value = "Calitate conducator doctorat";
            ws.Cell(5, 7).Value = "Varsta";
            ws.Cell(5, 8).Value = "Facultate";
            ws.Cell(5, 9).Value = "Departament";
            ws.Cell(5, 50).Value = "Total";
            for (int c = 1; c <= 9; c++) ws.Range(5, c, 7, c).Merge();
            ws.Range(5, 50, 7, 50).Merge();
            ws.Range(5, 1, 5, 50).Style.Fill.BackgroundColor = green;
            ws.Range(5, 1, 5, 50).Style.Font.FontColor = XLColor.White;
            ws.Range(5, 1, 8, 50).Style.Font.Bold = true;
            ws.Range(5, 1, 8, 50).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(5, 1, 8, 50).Style.Alignment.WrapText = true;
            ws.Column(1).Width = 5; ws.Column(2).Width = 30;
            ws.Column(3).Width = 14; ws.Column(4).Width = 22;
            for (int c = 5; c <= 50; c++) ws.Column(c).Width = 12;

            int dataRow = 9;
            if (profesori.Count > 0) ws.Row(dataRow).InsertRowsBelow(profesori.Count - 1);

            for (int i = 0; i < profesori.Count; i++)
            {
                var (cnp, prof) = profesori[i];
                int r = dataRow + i;
                decimal norma = prof.ID_TipGrad.HasValue &&
                                norme.TryGetValue(prof.ID_TipGrad.Value, out var n) ? n : 15m;
                ws.Cell(r, 1).Value = i + 1;
                ws.Cell(r, 2).Value = prof.NumeIntreg;
                ws.Cell(r, 3).Value = "";
                ws.Cell(r, 4).Value = prof.GradDidactic;
                ws.Cell(r, 5).Value = 1; ws.Cell(r, 6).Value = 0;
                ws.Cell(r, 7).Value = ""; ws.Cell(r, 8).Value = prof.Facultate;
                ws.Cell(r, 9).Value = prof.Departament;
                if (fractiuni.TryGetValue(cnp, out var frac))
                    foreach (var kv in frac)
                    {
                        int col = 9 + kv.Key;
                        if (col >= 10 && col <= 49)
                            ws.Cell(r, col).Value = Math.Round(kv.Value, 2);
                    }
                ws.Cell(r, 50).FormulaA1 = $"=SUM(J{r}:AW{r})";
                if (i % 2 != 0)
                    for (int c = 1; c <= 50; c++)
                        ws.Cell(r, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5f5f5");
                for (int c = 1; c <= 50; c++)
                    ws.Cell(r, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int totalRow = dataRow + profesori.Count;
            ws.Cell(totalRow, 2).Value = "TOTAL GENERAL"; ws.Cell(totalRow, 2).Style.Font.Bold = true;
            for (int c = 10; c <= 49; c++)
            {
                string cl = ColLetter(c);
                ws.Cell(totalRow, c).FormulaA1 = $"=SUM({cl}{dataRow}:{cl}{totalRow - 1})";
                ws.Cell(totalRow, c).Style.Font.Bold = true;
            }
            ws.Cell(totalRow, 50).FormulaA1 = $"=SUM(J{totalRow}:AW{totalRow})";
            ws.Cell(totalRow, 50).Style.Font.Bold = true;

            using var stream = new MemoryStream();
            wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Raport_ANS_{idAnUniv}.xlsx");
        }

        private static string ColLetter(int col)
        {
            string r = "";
            while (col > 0) { col--; r = (char)('A' + col % 26) + r; col /= 26; }
            return r;
        }

        private class ProfAns
        {
            public string NumeIntreg { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string Departament { get; set; } = "";
            public string GradDidactic { get; set; } = "";
            public int? ID_TipGrad { get; set; }
        }
    }
}