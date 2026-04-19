using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Data;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NormaTotaluriController : ControllerBase
    {
        private readonly string _connectionString;
        private const string BrandColor = "#56723e";

        public NormaTotaluriController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        // Valori reale DenumireFormaInv din DB (verificate): 'Cu frecvență', 'Învățământ la distanță', 'Frecvență redusă'
        private const string SqlTotaluri = @"
            WITH DateBrute AS (
                SELECT
                    ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                    MIN(p.NumeIntreg) OVER
                        (PARTITION BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20))))
                                                                         AS NumeIntreg,
                    MIN(p.DenumireCatedra) OVER
                        (PARTITION BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20))))
                                                                         AS Departament,
                    MIN(p.DenumireFacultate) OVER
                        (PARTITION BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20))))
                                                                         AS Facultate,
                    CASE
                        WHEN ppm.DenumireFormaInv = N'Cu frecvență'           THEN 'IF'
                        WHEN ppm.DenumireFormaInv = N'Învățământ la distanță' THEN 'ID'
                        WHEN ppm.DenumireFormaInv = N'Frecvență redusă'       THEN 'IFR'
                        ELSE 'IF'
                    END                                                  AS FormaInv,
                    ppm.Denumire                                         AS Materie,
                    ppm.NrSemestruDinAn                                  AS Semestru,
                    ppm.NrOreConventionale
                FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                INNER JOIN [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                    ON ppm.ID_Profesor = p.ID_Profesor
                    AND p.ID_AnUnivCatedra = @ID_AnUniv
                WHERE ppm.ID_AnUniv = @ID_AnUniv
                    AND (@fac = N'Toti'
                         OR p.ID_Facultate = TRY_CAST(@fac AS INT)
                         OR p.DenumireFacultate COLLATE DATABASE_DEFAULT
                          = @fac COLLATE DATABASE_DEFAULT)
                    AND (@dept = N'Toti'
                         OR p.ID_Catedra = TRY_CAST(@dept AS INT)
                         OR p.DenumireCatedra COLLATE DATABASE_DEFAULT
                          = @dept COLLATE DATABASE_DEFAULT)
                    AND (@prof = N'Toti'
                         OR p.NumeIntreg COLLATE DATABASE_DEFAULT
                          = @prof COLLATE DATABASE_DEFAULT)
            ),
            Dedup AS (
                SELECT IdentificatorUnic, NumeIntreg, Departament, Facultate,
                       FormaInv, Materie, Semestru, MAX(NrOreConventionale) AS OreConv
                FROM DateBrute
                GROUP BY IdentificatorUnic, NumeIntreg, Departament, Facultate,
                         FormaInv, Materie, Semestru
            ),
            Agregat AS (
                SELECT IdentificatorUnic, NumeIntreg, Departament, Facultate,
                    CAST(SUM(CASE WHEN FormaInv='IF'  THEN OreConv ELSE 0 END) AS DECIMAL(10,2)) AS OreIF,
                    CAST(SUM(CASE WHEN FormaInv='ID'  THEN OreConv ELSE 0 END) AS DECIMAL(10,2)) AS OreID,
                    CAST(SUM(CASE WHEN FormaInv='IFR' THEN OreConv ELSE 0 END) AS DECIMAL(10,2)) AS OreIFR,
                    CAST(SUM(OreConv) AS DECIMAL(10,2))                                           AS TotalOreConv
                FROM Dedup GROUP BY IdentificatorUnic, NumeIntreg, Departament, Facultate
            )
            SELECT NumeIntreg, Departament, Facultate,
                   OreIF, OreID, OreIFR, TotalOreConv,
                   CAST(TotalOreConv * 14 AS DECIMAL(10,2)) AS TotalAnual
            FROM Agregat ORDER BY NumeIntreg";

        private void AddParams(SqlCommand cmd, int id, string fac, string dept, string prof)
        {
            cmd.Parameters.AddWithValue("@ID_AnUniv", id);
            cmd.Parameters.AddWithValue("@fac", fac);
            cmd.Parameters.AddWithValue("@dept", dept);
            cmd.Parameters.AddWithValue("@prof", prof);
        }

        [HttpGet]
        public async Task<IActionResult> GetTotaluri(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null,
            [FromQuery] string? departament = null,
            [FromQuery] string? profesor = null)
        {
            var fac = string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim();
            var dept = string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim();
            var prof = string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim();
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri, conn);
            cmd.CommandTimeout = 120;
            AddParams(cmd, idAnUniv, fac, dept, prof);
            using var r = await cmd.ExecuteReaderAsync();
            int nrCrt = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nrCrt++,
                    Profesor = r["NumeIntreg"].ToString(),
                    Departament = r["Departament"].ToString(),
                    Facultate = r["Facultate"].ToString(),
                    OreIF = Convert.ToDecimal(r["OreIF"]),
                    OreID = Convert.ToDecimal(r["OreID"]),
                    OreIFR = Convert.ToDecimal(r["OreIFR"]),
                    TotalOreConv = Convert.ToDecimal(r["TotalOreConv"]),
                    TotalAnual = Convert.ToDecimal(r["TotalAnual"])
                });
            return Ok(result);
        }

        [HttpGet("export")]
        public async Task<IActionResult> ExportTotaluri(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null,
            [FromQuery] string? departament = null,
            [FromQuery] string? profesor = null)
        {
            var fac = string.IsNullOrWhiteSpace(facultate) ? "Toti" : facultate.Trim();
            var dept = string.IsNullOrWhiteSpace(departament) ? "Toti" : departament.Trim();
            var prof = string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]
            {
                new DataColumn("Nr. Crt.", typeof(int)),
                new DataColumn("Profesor"),
                new DataColumn("Departament"),
                new DataColumn("Facultate"),
                new DataColumn("Ore IF",          typeof(decimal)),
                new DataColumn("Ore ID",          typeof(decimal)),
                new DataColumn("Ore IFR",         typeof(decimal)),
                new DataColumn("Total Ore Conv.", typeof(decimal)),
                new DataColumn("Total Anual",     typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri, conn);
            cmd.CommandTimeout = 120;
            AddParams(cmd, idAnUniv, fac, dept, prof);
            using var r = await cmd.ExecuteReaderAsync();
            int nrCrt = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nrCrt++, r["NumeIntreg"].ToString(), r["Departament"].ToString(), r["Facultate"].ToString(),
                    Convert.ToDecimal(r["OreIF"]), Convert.ToDecimal(r["OreID"]), Convert.ToDecimal(r["OreIFR"]),
                    Convert.ToDecimal(r["TotalOreConv"]), Convert.ToDecimal(r["TotalAnual"]));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Totaluri Norme");
            ws.Cell(1, 1).Value = $"Totaluri norme | An: {idAnUniv}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var col in new[] { "Ore IF", "Ore ID", "Ore IFR", "Total Ore Conv.", "Total Anual" })
                tbl.Field(col).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr. Crt.").TotalsRowLabel = "TOTAL";
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColor);
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Font.FontColor = XLColor.White;
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Totaluri_Norme_{idAnUniv}.xlsx");
        }
    }
}