using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Data;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TitulariController : ControllerBase
    {
        private readonly string _connectionString;
        private const string BrandColor = "#56723e";

        public TitulariController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        // Filtrare exclusiv prin TitularAnUniv = 1 din DB
        // Unificare prin ISNULL(CNP, CAST(ID_Profesor AS VARCHAR))
        private const string SqlTitulari = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                MIN(p.NumeIntreg)                                    AS NumeIntreg,
                MIN(p.DenumireCatedra)                               AS Departament,
                MIN(p.DenumireFacultate)                             AS Facultate,
                MIN(p.DenumireGradDidactic)                          AS GradDidactic,
                MIN(p.ID_TipGradDidacticAnUniv)                      AS ID_TipGrad
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            WHERE p.ID_AnUnivCatedra = @ID_AnUniv
                AND p.TitularAnUniv = 1
                AND (@facultate = 'Toti'
                     OR p.DenumireFacultate COLLATE DATABASE_DEFAULT = @facultate COLLATE DATABASE_DEFAULT)
                AND (@departament = 'Toti'
                     OR p.DenumireCatedra COLLATE DATABASE_DEFAULT = @departament COLLATE DATABASE_DEFAULT)
            GROUP BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))
            ORDER BY MIN(p.NumeIntreg)";

        [HttpGet]
        public async Task<IActionResult> GetTitulari(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null,
            [FromQuery] string? departament = null)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulari, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
            cmd.Parameters.AddWithValue("@facultate", string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim());
            cmd.Parameters.AddWithValue("@departament", string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim());

            using var r = await cmd.ExecuteReaderAsync();
            int nrCrt = 1;
            while (await r.ReadAsync())
            {
                result.Add(new
                {
                    NrCrt = nrCrt++,
                    Profesor = r["NumeIntreg"].ToString(),
                    Departament = r["Departament"].ToString(),
                    Facultate = r["Facultate"].ToString(),
                    Grad = r["GradDidactic"].ToString()
                });
            }
            return Ok(result);
        }

        [HttpGet("export")]
        public async Task<IActionResult> ExportTitulari(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null,
            [FromQuery] string? departament = null)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nr. Crt.", typeof(int)),
                new DataColumn("Nume si prenume"),
                new DataColumn("Departament"),
                new DataColumn("Facultate"),
                new DataColumn("Grad didactic")
            });

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulari, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
            cmd.Parameters.AddWithValue("@facultate", string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim());
            cmd.Parameters.AddWithValue("@departament", string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim());

            using var r = await cmd.ExecuteReaderAsync();
            int nrCrt = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nrCrt++, r["NumeIntreg"].ToString(), r["Departament"].ToString(),
                    r["Facultate"].ToString(), r["GradDidactic"].ToString());

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Titulari");
            ws.Cell(1, 1).Value = "Cadre didactice titulare | An: " + idAnUniv;
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();

            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Nr. Crt.").TotalsRowFunction = XLTotalsRowFunction.Count;
            tbl.Field("Nume si prenume").TotalsRowLabel = "TOTAL";

            ws.Range(3, 1, 3, dt.Columns.Count).Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColor);
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Font.FontColor = XLColor.White;
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Titulari_" + idAnUniv + ".xlsx");
        }
    }
}