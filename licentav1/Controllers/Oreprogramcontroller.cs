using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Data;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OreProgramController : ControllerBase
    {
        private readonly string _connectionString;
        private const string BrandColor = "#56723e";

        public OreProgramController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        // Sincronizat cu Excel: adaugam Nr Ore Curs si Nr Ore Aplicatii per program
        private const string SqlOreProgram = @"
    WITH OrePerProgram AS (
        SELECT
            ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
            p.NumeIntreg,
            -- Curatam denumirea specializarii (taiem dupa '...')
            LTRIM(RTRIM(
                CASE 
                    WHEN CHARINDEX('...', ppm.DenumireSpecializare) > 0 
                    THEN LEFT(ppm.DenumireSpecializare, CHARINDEX('...', ppm.DenumireSpecializare) - 1)
                    ELSE ppm.DenumireSpecializare 
                END
            )) AS ProgramStudiuCurat,
            SUM(ppm.Nr_Ore_Curs)                                 AS OreCursProgram,
            SUM(ppm.Nr_Ore_Seminar + ppm.Nr_Ore_Laborator + ppm.Nr_Ore_Proiect) AS OreAplicatiiProgram,
            SUM(ppm.NrOreConventionale)                          AS OreConvProgram
        FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
        INNER JOIN [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p 
            ON ppm.ID_Profesor = p.ID_Profesor
        WHERE ppm.ID_AnUniv = @ID_AnUniv
            -- Filtram dupa BIROUL profesorului (p), nu dupa unde tine cursul (ppm)
            -- Asta rezolva cazul Baicoianu (apare la Mate-Info chiar daca preda la Stiinte Ec.)
            AND (@fac = N'Toti' 
                 OR p.ID_Facultate = TRY_CAST(@fac AS INT) 
                 OR p.DenumireFacultate COLLATE DATABASE_DEFAULT = @fac COLLATE DATABASE_DEFAULT)
            AND (@dept = N'Toti' 
                 OR p.ID_Catedra = TRY_CAST(@dept AS INT) 
                 OR p.DenumireCatedra COLLATE DATABASE_DEFAULT = @dept COLLATE DATABASE_DEFAULT)
            AND (@prof = N'Toti' 
                 OR p.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
            
            -- IATA UNDE VINE FILTRUL DE SPECIALIZARI (Smart Match):
            AND (@specs = N'Toti' OR EXISTS (
                SELECT 1 FROM STRING_SPLIT(@specs, ',') s 
                WHERE ppm.DenumireSpecializare LIKE s.value + '%' 
                   OR ppm.DenumireSpecializare LIKE '%' + s.value + '%'
            ))
        GROUP BY 
            ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20))), 
            p.NumeIntreg,
            CASE 
                WHEN CHARINDEX('...', ppm.DenumireSpecializare) > 0 
                THEN LEFT(ppm.DenumireSpecializare, CHARINDEX('...', ppm.DenumireSpecializare) - 1)
                ELSE ppm.DenumireSpecializare 
            END
    ),
    TotalPerProfesor AS (
        SELECT IdentificatorUnic, SUM(OreConvProgram) AS TotalPost FROM OrePerProgram GROUP BY IdentificatorUnic
    )
    SELECT
        o.NumeIntreg                                       AS Profesor,
        o.ProgramStudiuCurat                               AS ProgramStudiu,
        CAST(o.OreCursProgram AS DECIMAL(10,2))           AS NrOreCurs,
        CAST(o.OreAplicatiiProgram AS DECIMAL(10,2))      AS NrOreAplicatii,
        CAST(o.OreConvProgram AS DECIMAL(10,2))           AS OreConvProgram,
        CAST(t.TotalPost AS DECIMAL(10,2))                AS TotalPost,
        CAST(CASE WHEN t.TotalPost > 0 THEN (o.OreConvProgram * 100.0) / t.TotalPost ELSE 0 END AS DECIMAL(10,2)) AS ProcentPost
    FROM OrePerProgram o
    INNER JOIN TotalPerProfesor t ON t.IdentificatorUnic = o.IdentificatorUnic
    ORDER BY o.NumeIntreg ASC";

        private void AddParams(SqlCommand cmd, int id, string fac, string dept, string prof)
        {
            cmd.Parameters.AddWithValue("@ID_AnUniv", id);
            cmd.Parameters.AddWithValue("@fac", fac);
            cmd.Parameters.AddWithValue("@dept", dept);
            cmd.Parameters.AddWithValue("@prof", prof);
        }

        [HttpGet]
        public async Task<IActionResult> GetOreProgram(
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
            using var cmd = new SqlCommand(SqlOreProgram, conn);
            cmd.CommandTimeout = 120;
            AddParams(cmd, idAnUniv, fac, dept, prof);
            using var r = await cmd.ExecuteReaderAsync();
            int nrCrt = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nrCrt++,
                    Profesor = r["Profesor"].ToString(),
                    ProgramStudiu = r["ProgramStudiu"].ToString(),
                    NrOreCurs = Convert.ToDecimal(r["NrOreCurs"]),
                    NrOreAplicatii = Convert.ToDecimal(r["NrOreAplicatii"]),
                    OreConvProgram = Convert.ToDecimal(r["OreConvProgram"]),
                    TotalPost = Convert.ToDecimal(r["TotalPost"]),
                    ProcentPost = Convert.ToDecimal(r["ProcentPost"])
                });
            return Ok(result);
        }

        [HttpGet("export")]
        public async Task<IActionResult> ExportOreProgram(
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
                new DataColumn("Nr. Crt.",         typeof(int)),
                new DataColumn("Profesor"),
                new DataColumn("Program Studiu"),
                new DataColumn("Nr Ore Curs",       typeof(decimal)),
                new DataColumn("Nr Ore Aplicatii",  typeof(decimal)),
                new DataColumn("Ore Conv. Program", typeof(decimal)),
                new DataColumn("Total Post",        typeof(decimal)),
                new DataColumn("Procent Post %",    typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlOreProgram, conn);
            cmd.CommandTimeout = 120;
            AddParams(cmd, idAnUniv, fac, dept, prof);
            using var r = await cmd.ExecuteReaderAsync();
            int nrCrt = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nrCrt++, r["Profesor"].ToString(), r["ProgramStudiu"].ToString(),
                    Convert.ToDecimal(r["NrOreCurs"]), Convert.ToDecimal(r["NrOreAplicatii"]),
                    Convert.ToDecimal(r["OreConvProgram"]), Convert.ToDecimal(r["TotalPost"]),
                    Convert.ToDecimal(r["ProcentPost"]));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Distributie Ore");
            ws.Cell(1, 1).Value = $"Distributie ore per program | An: {idAnUniv}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Conv. Program").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Aplicatii").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr. Crt.").TotalsRowLabel = "TOTAL";
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColor);
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Font.FontColor = XLColor.White;
            ws.Range(3, 1, 3, dt.Columns.Count).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"StatisticaOre_{idAnUniv}.xlsx");
        }
    }
}