using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ForeignLangController : ControllerBase
    {
        private readonly string _connectionString;
        private const string BrandColor = "#56723e";

        public ForeignLangController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        private const string SqlLimbi = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                MIN(p.NumeIntreg)                                    AS NumeIntreg,
                MIN(ppm.DenumireFacultate)                           AS Facultate,
                MIN(ppm.DenumireCatedra)                             AS Departament,
                ppm.LimbaDePredare                                   AS Limba,
                ppm.NrSemestruDinAn                                  AS Semestru,
                ppm.ApartineDeCuplaj,
                MAX(ppm.NrOreConventionale)                          AS OreConventionale
            FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
            INNER JOIN [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                ON ppm.ID_Profesor = p.ID_Profesor
                AND p.ID_AnUnivCatedra = @ID_AnUniv
            WHERE ppm.ID_AnUniv = @ID_AnUniv
                AND ppm.LimbaDePredare COLLATE DATABASE_DEFAULT
                    IN (N'Engleza', N'Franceza', N'Germana')
            GROUP BY
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20))),
                ppm.LimbaDePredare,
                ppm.NrSemestruDinAn,
                ppm.ApartineDeCuplaj
            ORDER BY MIN(p.NumeIntreg), ppm.LimbaDePredare, ppm.NrSemestruDinAn";

        private Dictionary<string, ProfLimba> RunQuery(int idAnUniv, string? limba)
        {
            var acc = new Dictionary<string, ProfLimba>();
            var cuplaje = new Dictionary<string, HashSet<long>>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(SqlLimbi, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var cnp = r["IdentificatorUnic"].ToString()!;
                var limb = r["Limba"].ToString()!.Trim();
                int sem = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]);
                decimal o = r["OreConventionale"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConventionale"]);
                long? cup = r["ApartineDeCuplaj"] == DBNull.Value ? null : (long?)Convert.ToInt64(r["ApartineDeCuplaj"]);

                if (!string.IsNullOrEmpty(limba) && !string.Equals(limb, limba, StringComparison.OrdinalIgnoreCase)) continue;

                var key = cnp + "|" + limb;
                if (!cuplaje.ContainsKey(key)) cuplaje[key] = new HashSet<long>();
                if (cup.HasValue && !cuplaje[key].Add(cup.Value)) continue;

                if (!acc.ContainsKey(key))
                    acc[key] = new ProfLimba
                    {
                        NumeIntreg = r["NumeIntreg"].ToString()!,
                        Facultate = r["Facultate"].ToString()!,
                        Departament = r["Departament"].ToString()!,
                        Limba = limb
                    };

                if (sem == 1) acc[key].OreSem1 += o;
                else if (sem == 2) acc[key].OreSem2 += o;
                else acc[key].OreSem1 += o;
            }
            return acc;
        }

        [HttpGet]
        public IActionResult GetLimbiStraine([FromQuery] int idAnUniv = 45, [FromQuery] string? limba = null)
        {
            var data = RunQuery(idAnUniv, limba);
            int nr = 1;
            var result = data.Values.OrderBy(x => x.Limba).ThenBy(x => x.NumeIntreg)
                .Select(x => new {
                    NrCrt = nr++,
                    x.NumeIntreg,
                    x.Facultate,
                    x.Departament,
                    x.Limba,
                    OreSem1 = Math.Round(x.OreSem1, 2),
                    OreSem2 = Math.Round(x.OreSem2, 2),
                    Total = Math.Round(x.OreSem1 + x.OreSem2, 2)
                }).ToList();
            return Ok(result);
        }

        [HttpGet("export")]
        public IActionResult ExportLimbi([FromQuery] int idAnUniv = 45, [FromQuery] string? limba = null)
        {
            var data = RunQuery(idAnUniv, limba);
            var wb = new XLWorkbook();
            foreach (var grup in data.Values.GroupBy(x => x.Limba).OrderBy(g => g.Key))
            {
                var ws = wb.Worksheets.Add(grup.Key.Length > 31 ? grup.Key[..31] : grup.Key);
                ws.Cell(1, 1).Value = "Ore predate in limba: " + grup.Key + " | An: " + idAnUniv;
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
                ws.Range(1, 1, 1, 7).Merge();
                string[] h = { "Nr.", "Nume si prenume", "Facultate", "Departament", "Ore Sem.1", "Ore Sem.2", "Total" };
                for (int c = 0; c < h.Length; c++)
                {
                    ws.Cell(3, c + 1).Value = h[c];
                    ws.Cell(3, c + 1).Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColor);
                    ws.Cell(3, c + 1).Style.Font.FontColor = XLColor.White;
                    ws.Cell(3, c + 1).Style.Font.Bold = true;
                }
                var lista = grup.OrderBy(x => x.NumeIntreg).ToList();
                for (int i = 0; i < lista.Count; i++)
                {
                    int row = 4 + i;
                    ws.Cell(row, 1).Value = i + 1; ws.Cell(row, 2).Value = lista[i].NumeIntreg;
                    ws.Cell(row, 3).Value = lista[i].Facultate; ws.Cell(row, 4).Value = lista[i].Departament;
                    ws.Cell(row, 5).Value = Math.Round(lista[i].OreSem1, 2);
                    ws.Cell(row, 6).Value = Math.Round(lista[i].OreSem2, 2);
                    ws.Cell(row, 7).Value = Math.Round(lista[i].OreSem1 + lista[i].OreSem2, 2);
                    if (i % 2 != 0) ws.Range(row, 1, row, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5f5f5");
                }
                int tr = 4 + lista.Count;
                ws.Cell(tr, 2).Value = "TOTAL"; ws.Cell(tr, 2).Style.Font.Bold = true;
                ws.Cell(tr, 5).FormulaA1 = "=SUM(E4:E" + (tr - 1) + ")";
                ws.Cell(tr, 6).FormulaA1 = "=SUM(F4:F" + (tr - 1) + ")";
                ws.Cell(tr, 7).FormulaA1 = "=SUM(G4:G" + (tr - 1) + ")";
                ws.Range(tr, 1, tr, 7).Style.Font.Bold = true;
                ws.Columns().AdjustToContents();
            }
            using var stream = new MemoryStream();
            wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                string.IsNullOrEmpty(limba) ? "Raport_Limbi_Straine.xlsx" : "Raport_Limba_" + limba + ".xlsx");
        }

        private class ProfLimba
        {
            public string NumeIntreg { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string Departament { get; set; } = "";
            public string Limba { get; set; } = "";
            public decimal OreSem1 { get; set; }
            public decimal OreSem2 { get; set; }
        }
    }
}