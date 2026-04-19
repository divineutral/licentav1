using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Data;
using System.IO.Compression;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DisciplineController : ControllerBase
    {
        private readonly string _connectionString;
        private const string BrandColor = "#56723e";

        public DisciplineController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        // SQL Server 2016: STUFF + FOR XML PATH in loc de STRING_AGG DISTINCT
        // .value('.','NVARCHAR(MAX)') previne erorile de encoding XML
        private const string SqlDiscipline = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))       AS IdentificatorUnic,
                MIN(p.NumeIntreg)                                        AS NumeIntreg,
                MIN(p.DenumireCatedra)                                   AS Departament,
                MIN(p.DenumireFacultate)                                 AS Facultate,
                ppm.DenumireFormaInv                                     AS FormaInv,
                STUFF((
                    SELECT DISTINCT N' | ' + ISNULL(d2.Denumire, N'')
                    FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] d2
                    WHERE d2.ID_Profesor  = ppm.ID_Profesor
                      AND d2.ID_AnUniv    = ppm.ID_AnUniv
                      AND d2.DenumireFormaInv COLLATE DATABASE_DEFAULT
                        = ppm.DenumireFormaInv COLLATE DATABASE_DEFAULT
                      AND d2.Denumire IS NOT NULL
                      AND LTRIM(RTRIM(d2.Denumire)) != N''
                    FOR XML PATH(N''), TYPE
                ).value(N'.', N'NVARCHAR(MAX)'), 1, 3, N'')              AS Discipline
            FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
            INNER JOIN [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                ON ppm.ID_Profesor      = p.ID_Profesor
                AND p.ID_AnUnivCatedra  = @ID_AnUniv
            WHERE ppm.ID_AnUniv = @ID_AnUniv
                AND ppm.Denumire IS NOT NULL
                AND LTRIM(RTRIM(ppm.Denumire)) != N''
                AND (@fac = N'Toti'
                     OR p.ID_Facultate = TRY_CAST(@fac AS INT)
                     OR ppm.DenumireFacultate COLLATE DATABASE_DEFAULT
                      = @fac COLLATE DATABASE_DEFAULT)
                AND (@dept = N'Toti'
                     OR p.ID_Catedra = TRY_CAST(@dept AS INT)
                     OR ppm.DenumireCatedra COLLATE DATABASE_DEFAULT
                      = @dept COLLATE DATABASE_DEFAULT)
            GROUP BY
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20))),
                ppm.ID_Profesor,
                ppm.ID_AnUniv,
                ppm.DenumireFormaInv
            ORDER BY MIN(p.NumeIntreg), ppm.DenumireFormaInv";

        private void AddParams(SqlCommand cmd, int id, string fac, string dept)
        {
            cmd.Parameters.AddWithValue("@ID_AnUniv", id);
            cmd.Parameters.AddWithValue("@fac", fac);
            cmd.Parameters.AddWithValue("@dept", dept);
        }

        [HttpGet]
        public async Task<IActionResult> GetDiscipline(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null,
            [FromQuery] string? departament = null)
        {
            var fac = string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim();
            var dept = string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim();
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDiscipline, conn);
            cmd.CommandTimeout = 180;
            AddParams(cmd, idAnUniv, fac, dept);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                result.Add(new
                {
                    Profesor = r["NumeIntreg"].ToString(),
                    Departament = r["Departament"].ToString(),
                    Facultate = r["Facultate"].ToString(),
                    FormaInvatamant = r["FormaInv"].ToString(),
                    Discipline = r["Discipline"].ToString()
                });
            return Ok(result);
        }

        [HttpGet("export")]
        public async Task<IActionResult> ExportDiscipline(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null,
            [FromQuery] string? departament = null)
        {
            var fac = string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim();
            var dept = string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]
            {
                new DataColumn("Profesor"), new DataColumn("Departament"),
                new DataColumn("Facultate"), new DataColumn("Forma Inv."),
                new DataColumn("Discipline predate")
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDiscipline, conn);
            cmd.CommandTimeout = 180;
            AddParams(cmd, idAnUniv, fac, dept);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                dt.Rows.Add(r["NumeIntreg"].ToString(), r["Departament"].ToString(),
                    r["Facultate"].ToString(), r["FormaInv"].ToString(), r["Discipline"].ToString());

            var forme = dt.AsEnumerable()
                .Select(row => row["Forma Inv."].ToString())
                .Distinct().Where(f => !string.IsNullOrWhiteSpace(f)).ToList();

            using var memZip = new MemoryStream();
            using (var archive = new ZipArchive(memZip, ZipArchiveMode.Create, true))
            {
                foreach (var forma in forme)
                {
                    var rows = dt.AsEnumerable()
                        .Where(row => row["Forma Inv."].ToString() == forma).ToList();
                    if (rows.Count == 0) continue;
                    var subset = rows.CopyToDataTable();
                    var safeName = string.Concat((forma ?? "N")
                        .Take(25).Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                    var entry = archive.CreateEntry($"Discipline_{safeName}.xlsx");
                    using var es = entry.Open();
                    using var wb = new XLWorkbook();
                    var sheetName = $"Disc {forma}";
                    var ws = wb.Worksheets.Add(sheetName.Length > 31 ? sheetName[..31] : sheetName);
                    ws.Cell(1, 1).Value = $"Discipline - {forma} | An: {idAnUniv}";
                    ws.Cell(1, 1).Style.Font.Bold = true;
                    ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
                    ws.Range(1, 1, 1, 5).Merge();
                    var tbl = ws.Cell(3, 1).InsertTable(subset);
                    tbl.Theme = XLTableTheme.None;
                    ws.Range(3, 1, 3, subset.Columns.Count).Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColor);
                    ws.Range(3, 1, 3, subset.Columns.Count).Style.Font.FontColor = XLColor.White;
                    ws.Range(3, 1, 3, subset.Columns.Count).Style.Font.Bold = true;
                    ws.Columns(1, 4).AdjustToContents();
                    ws.Column(5).Width = 80; ws.Column(5).Style.Alignment.WrapText = true;
                    using var wbs = new MemoryStream(); wb.SaveAs(wbs); wbs.Position = 0; wbs.CopyTo(es);
                }
            }
            memZip.Position = 0;
            return File(memZip.ToArray(), "application/zip", "Discipline_Predate.zip");
        }
    }
}