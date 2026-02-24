using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RapoarteController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;
        private readonly IMemoryCache _cache;
        private const string BrandColorHex = "#56723e";

        public RapoarteController(IConfiguration configuration, IMemoryCache cache)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAni()
        {
            return Ok(_cache.GetOrCreate("ListaAniUniv", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<object>();

                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    string sql = @"
                        SELECT UPPER(LTRIM(RTRIM(Denumire))) COLLATE DATABASE_DEFAULT AS AnCurat 
                        FROM [AGSIS].[dbo].[AnUniversitar] 
                        WHERE Denumire IS NOT NULL 
                        ORDER BY Ordine DESC";

                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var an = reader["AnCurat"]?.ToString() ?? "";
                        lista.Add(new { id = an, nume = an });
                    }
                }
                return lista;
            }));
        }

        [HttpGet("liste/facultati")]
        public ActionResult GetFacultati()
        {
            return Ok(_cache.GetOrCreate("ListaFacultati", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<string> { "Toti" };

                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    string sql = @"
                        SELECT DISTINCT 
                            UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT as FacCurata
                        FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                        WHERE ppm.DenumireFacultate IS NOT NULL
                        ORDER BY FacCurata ASC";

                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        lista.Add(reader["FacCurata"]?.ToString() ?? "");
                    }
                }
                return lista;
            }));
        }

        [HttpGet("liste/specializari-per-facultate")]
        public ActionResult GetSpecializari(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql;
                if (!string.IsNullOrEmpty(numeFacultate) && numeFacultate != "Toti")
                {
                    sql = @"
                    DECLARE @TargetFacId INT;
                    SELECT TOP 1 @TargetFacId = ID_FacultateSpecializare
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie]
                    WHERE UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) = @fac
                    GROUP BY ID_FacultateSpecializare
                    ORDER BY COUNT(*) DESC;

                    SELECT DISTINCT 
                        UPPER(LTRIM(RTRIM(
                            REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                                CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                     THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                     ELSE sf.DenumireSpecializare END, 
                            ' - CORECT', ''), ' CORECT', ''), ' - COPIE', ''), 'Ș', 'S'), 'Ț', 'T')
                        ))) as SpecCurata
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    WHERE sf.ID_Facultate = @TargetFacId 
                      AND sf.DenumireSpecializare IS NOT NULL
                    ORDER BY SpecCurata";
                }
                else
                {
                    sql = @"
                    SELECT DISTINCT 
                        UPPER(LTRIM(RTRIM(
                            REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                                CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                     THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                     ELSE sf.DenumireSpecializare END, 
                            ' - CORECT', ''), ' CORECT', ''), ' - COPIE', ''), 'Ș', 'S'), 'Ț', 'T')
                        ))) as SpecCurata
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    WHERE sf.DenumireSpecializare IS NOT NULL
                    ORDER BY SpecCurata";
                }

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                if (!string.IsNullOrEmpty(numeFacultate) && numeFacultate != "Toti")
                {
                    cmd.Parameters.AddWithValue("@fac", numeFacultate);
                }

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string val = reader["SpecCurata"]?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(val) && !lista.Contains(val))
                        lista.Add(val);
                }
            }
            return Ok(lista);
        }

        [HttpGet("liste/profesori-per-specializari")]
        public ActionResult GetProfesori(string? anUniv, string? facultate, string? specializari)
        {
            var lista = new List<string> { "Toti" };
            bool toateSpecializarile = string.IsNullOrEmpty(specializari) || specializari == "Toti";

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                SELECT DISTINCT ppm.NumeIntreg
                FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                WHERE 
                    (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                    AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) = @fac)
                    AND (@allSpecs = 1 OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'Ș', 'S'), 'Ț', 'T')))) IN (SELECT value FROM STRING_SPLIT(@listaSpecs, ',')))
                ORDER BY ppm.NumeIntreg";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@allSpecs", toateSpecializarile ? 1 : 0);
                cmd.Parameters.AddWithValue("@listaSpecs", specializari ?? "");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    lista.Add(reader["NumeIntreg"]?.ToString() ?? "");
                }
            }
            return Ok(lista);
        }

        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti")
        {
            var result = new List<object>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg AS Profesor,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'Ș', 'S'), 'Ț', 'T')))) AS Specializare,
                        ISNULL(sf.DenumireMaterie, 'Nedefinit') AS Materie,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        (ISNULL(sf.Nr_Ore_Curs, 0) * ISNULL(sf.CoefOreConvCurs, 1)) + 
                        ((ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0)) * ISNULL(sf.CoefOreConvApp, 1)) AS OreConvLinie,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                    LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                ),
                Filtrat AS (
                    SELECT 
                        Profesor, Specializare, Materie, TipPost, Semestru, SUM(OreConvLinie) AS TotalOreConvItem
                    FROM BaseData
                    WHERE 
                        (@an = 'Toti' OR AnCurat = @an) AND
                        (@fac = 'Toti' OR FacultateCurata = @fac) AND
                        (@prof = 'Toti' OR Profesor = @prof) AND
                        (@specs = 'Toti' OR Specializare IN (SELECT value FROM STRING_SPLIT(@specs, ','))) AND
                        (@semestru = 0 OR Semestru = @semestru) AND
                        (@tipPost = 'Toti' OR TipPost = @tipPost)
                    GROUP BY Profesor, Specializare, Materie, TipPost, Semestru
                ),
                TotalProfesor AS (
                    SELECT Profesor, SUM(TotalOreConvItem) AS TotalOreConvProf
                    FROM Filtrat
                    GROUP BY Profesor
                )
                SELECT 
                    f.Profesor, f.Specializare, f.Materie, f.TipPost, f.Semestru, f.TotalOreConvItem AS OreConventionale, t.TotalOreConvProf AS TotalPost,
                    CAST(CASE WHEN t.TotalOreConvProf = 0 THEN 0 ELSE (f.TotalOreConvItem / t.TotalOreConvProf) * 100 END AS DECIMAL(10,2)) AS ProcentOre
                FROM Filtrat f
                INNER JOIN TotalProfesor t ON f.Profesor = t.Profesor
                ORDER BY f.Profesor, f.Materie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
                cmd.Parameters.AddWithValue("@semestru", semestru);
                cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        Profesor = reader["Profesor"],
                        Specializare = reader["Specializare"],
                        Materie = reader["Materie"],
                        TipPost = reader["TipPost"],
                        Semestru = reader["Semestru"],
                        OreConventionale = reader["OreConventionale"],
                        TotalPost = reader["TotalPost"],
                        ProcentOre = reader["ProcentOre"]
                    });
                }
            }
            return Ok(result);
        }

        [HttpGet("ore-profesor-program")]
        public async Task<IActionResult> GetOreProfProgram(string? anUniv = "Toti", string? facultate = "Toti", string? specializari = "Toti", string? profesor = "Toti", int semestru = 0, string tipPost = "Toti")
        {
            var listaResult = new List<object>();

            string sql = @"
            WITH BaseData AS (
                SELECT 
                    ppm.NumeIntreg AS Profesor,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'Ș', 'S'), 'Ț', 'T')))) AS ProgramStudiu,
                    (ISNULL(sf.Nr_Ore_Curs, 0) * ISNULL(sf.CoefOreConvCurs, 1)) + 
                    ((ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0)) * ISNULL(sf.CoefOreConvApp, 1)) AS OreConvLinie,
                    ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost,
                    ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) AS FacultateCurata,
                    UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat
                FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
            ),
            Filtrat AS (
                SELECT 
                    Profesor, ProgramStudiu, SUM(OreConvLinie) AS OreConvProgram
                FROM BaseData
                WHERE 
                    (@AnUniv = 'Toti' OR AnCurat = @AnUniv) AND
                    (@Facultate = 'Toti' OR FacultateCurata = @Facultate) AND
                    (@Profesor = 'Toti' OR Profesor = @Profesor) AND
                    (@Specializari = 'Toti' OR ProgramStudiu IN (SELECT value FROM STRING_SPLIT(@Specializari, ','))) AND
                    (@Semestru = 0 OR Semestru = @Semestru) AND
                    (@TipPost = 'Toti' OR TipPost = @TipPost)
                GROUP BY Profesor, ProgramStudiu
            ),
            TotalProfesor AS (
                SELECT Profesor, SUM(OreConvProgram) AS TotalPost
                FROM Filtrat
                GROUP BY Profesor
            )
            SELECT 
                f.Profesor, ISNULL(f.ProgramStudiu, 'Nespecificat') AS ProgramStudiu, f.OreConvProgram AS NrOreConv, t.TotalPost,
                CAST(CASE WHEN t.TotalPost = 0 THEN 0 ELSE (f.OreConvProgram / t.TotalPost) * 100 END AS DECIMAL(10,2)) AS ProcentPost
            FROM Filtrat f
            INNER JOIN TotalProfesor t ON f.Profesor = t.Profesor
            ORDER BY f.Profesor, f.OreConvProgram DESC";

            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                using (var command = new SqlCommand(sql, connection))
                {
                    command.Parameters.AddWithValue("@AnUniv", anUniv ?? "Toti");
                    command.Parameters.AddWithValue("@Facultate", facultate ?? "Toti");
                    command.Parameters.AddWithValue("@Specializari", specializari ?? "Toti");
                    command.Parameters.AddWithValue("@Profesor", profesor ?? "Toti");
                    command.Parameters.AddWithValue("@Semestru", semestru);
                    command.Parameters.AddWithValue("@TipPost", tipPost ?? "Toti");

                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            listaResult.Add(new
                            {
                                Profesor = reader["Profesor"].ToString(),
                                ProgramStudiu = reader["ProgramStudiu"].ToString(),
                                NrOreConv = Convert.ToDouble(reader["NrOreConv"]),
                                TotalPost = Convert.ToDouble(reader["TotalPost"]),
                                ProcentPost = Convert.ToDouble(reader["ProcentPost"])
                            });
                        }
                    }
                }
            }
            return Ok(listaResult);
        }

        [HttpGet("export/norma")]
        public IActionResult ExportNormaExcel(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti")
        {
            var result = new DataTable("NormaProfesori");
            result.Columns.AddRange(new[] {
                new DataColumn("Profesor"),
                new DataColumn("Specializare"),
                new DataColumn("Materie"),
                new DataColumn("Tip Post"),
                new DataColumn("Semestru"),
                new DataColumn("Ore Conventionale", typeof(double)),
                new DataColumn("Total Post", typeof(double))
            });

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'Ș', 'S'), 'Ț', 'T')))) AS SpecializareCurata,
                        ISNULL(sf.DenumireMaterie, 'Nedefinit') AS DenumireMaterie,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        (ISNULL(sf.Nr_Ore_Curs, 0) * ISNULL(sf.CoefOreConvCurs, 1)) + 
                        ((ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0)) * ISNULL(sf.CoefOreConvApp, 1)) AS OreConvLinie,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                    LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                ),
                Filtrat AS (
                    SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru, SUM(OreConvLinie) AS TotalOreConvItem
                    FROM BaseData
                    WHERE 
                        (@an = 'Toti' OR AnCurat = @an) AND (@fac = 'Toti' OR FacultateCurata = @fac) AND
                        (@prof = 'Toti' OR NumeIntreg = @prof) AND
                        (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ','))) AND
                        (@semestru = 0 OR Semestru = @semestru) AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                    GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru
                ),
                TotalProfesor AS (
                    SELECT NumeIntreg, SUM(TotalOreConvItem) AS TotalOreConvProf
                    FROM Filtrat GROUP BY NumeIntreg
                )
                SELECT f.NumeIntreg, f.SpecializareCurata, f.DenumireMaterie, f.TipPost, f.Semestru, f.TotalOreConvItem, t.TotalOreConvProf
                FROM Filtrat f INNER JOIN TotalProfesor t ON f.NumeIntreg = t.NumeIntreg
                ORDER BY f.NumeIntreg, f.DenumireMaterie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
                cmd.Parameters.AddWithValue("@semestru", semestru);
                cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Rows.Add(reader["NumeIntreg"], reader["SpecializareCurata"], reader["DenumireMaterie"], reader["TipPost"], reader["Semestru"], reader["TotalOreConvItem"], reader["TotalOreConvProf"]);
                }
            }

            string fileName = "NormaProfesori_General.xlsx";
            if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
            {
                fileName = $"NormaProfesori_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.xlsx";
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Norme");

                ws.Cell(1, 1).Value = "Filtre Aplicate";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);

                ws.Cell(2, 1).Value = $"An Universitar: {anUniv} | Facultate: {facultate} | Specializare: {specializari}";
                ws.Cell(3, 1).Value = $"Profesor: {profesor} | Semestru: {(semestru == 0 ? "Toate" : semestru.ToString())} | Tip Post: {tipPost}";

                var table = ws.Cell(5, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;

                ws.Columns().AdjustToContents();

                var headerRange = ws.Range(5, 1, 5, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;

                var dataRange = ws.Range(5, 1, 5 + result.Rows.Count, result.Columns.Count);
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }

        [HttpGet("export/ore-program")]
        public async Task<IActionResult> ExportOreProgramExcel(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti")
        {
            var result = new DataTable("OreProgram");
            result.Columns.AddRange(new[] {
                new DataColumn("Profesor"), new DataColumn("Program Studiu"), new DataColumn("Nr Ore Conv", typeof(double)),
                new DataColumn("Procent Post", typeof(double)), new DataColumn("Total Post", typeof(double))
            });

            string sql = @"
            WITH BaseData AS (
                SELECT ppm.NumeIntreg,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'Ș', 'S'), 'Ț', 'T')))) AS ProgramStudiu,
                    (ISNULL(sf.Nr_Ore_Curs, 0) * ISNULL(sf.CoefOreConvCurs, 1)) + ((ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0)) * ISNULL(sf.CoefOreConvApp, 1)) AS OreConvLinie,
                    ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost, ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) AS FacultateCurata,
                    UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat
                FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
            ),
            Filtrat AS (
                SELECT NumeIntreg, ProgramStudiu, SUM(OreConvLinie) AS OreConvProgram FROM BaseData
                WHERE (@AnUniv = 'Toti' OR AnCurat = @AnUniv) AND (@Facultate = 'Toti' OR FacultateCurata = @Facultate) AND
                      (@Profesor = 'Toti' OR NumeIntreg = @Profesor) AND (@Specializari = 'Toti' OR ProgramStudiu IN (SELECT value FROM STRING_SPLIT(@Specializari, ','))) AND
                      (@Semestru = 0 OR Semestru = @Semestru) AND (@TipPost = 'Toti' OR TipPost = @TipPost)
                GROUP BY NumeIntreg, ProgramStudiu
            ),
            TotalProfesor AS (SELECT NumeIntreg, SUM(OreConvProgram) AS TotalPost FROM Filtrat GROUP BY NumeIntreg)
            SELECT f.NumeIntreg, ISNULL(f.ProgramStudiu, 'Nespecificat') AS ProgramStudiu, f.OreConvProgram, t.TotalPost,
                CAST(CASE WHEN t.TotalPost = 0 THEN 0 ELSE (f.OreConvProgram / t.TotalPost) * 100 END AS DECIMAL(10,2)) AS ProcentPost
            FROM Filtrat f INNER JOIN TotalProfesor t ON f.NumeIntreg = t.NumeIntreg ORDER BY f.NumeIntreg, f.OreConvProgram DESC";

            using (var conn = new SqlConnection(_connectionString))
            {
                await conn.OpenAsync();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@AnUniv", anUniv ?? "Toti"); cmd.Parameters.AddWithValue("@Facultate", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@Specializari", specializari ?? "Toti"); cmd.Parameters.AddWithValue("@Profesor", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@Semestru", semestru); cmd.Parameters.AddWithValue("@TipPost", tipPost ?? "Toti");

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Rows.Add(reader["NumeIntreg"], reader["ProgramStudiu"], Convert.ToDouble(reader["OreConvProgram"]), Convert.ToDouble(reader["ProcentPost"]), Convert.ToDouble(reader["TotalPost"]));
                }
            }

            string fileName = "StatisticaOre_General.xlsx";
            if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
            {
                fileName = $"StatisticaOre_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.xlsx";
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Distributie Ore");

                ws.Cell(1, 1).Value = "Filtre Aplicate";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);

                ws.Cell(2, 1).Value = $"An Universitar: {anUniv} | Facultate: {facultate} | Specializare: {specializari}";
                ws.Cell(3, 1).Value = $"Profesor: {profesor} | Semestru: {(semestru == 0 ? "Toate" : semestru.ToString())} | Tip Post: {tipPost}";

                var table = ws.Cell(5, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;

                ws.Columns().AdjustToContents();

                var headerRange = ws.Range(5, 1, 5, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;

                var dataRange = ws.Range(5, 1, 5 + result.Rows.Count, result.Columns.Count);
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }

        private class PdfNormaData
        {
            public string Profesor { get; set; } = "";
            public string Specializare { get; set; } = "";
            public string Materie { get; set; } = "";
            public string Tip { get; set; } = "";
            public string Semestru { get; set; } = "";
            public string OreConv { get; set; } = "";
            public string TotalPost { get; set; } = "";
        }

        private class PdfOreData
        {
            public string Profesor { get; set; } = "";
            public string Program { get; set; } = "";
            public string OreConv { get; set; } = "";
            public string TotalPost { get; set; } = "";
            public string Procent { get; set; } = "";
        }

        [HttpGet("export/pdf/norma")]
        public IActionResult ExportNormaPdf(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti")
        {
            QuestPDF.Settings.License = LicenseType.Community;
            var data = new List<PdfNormaData>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT ppm.NumeIntreg, UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'Ș', 'S'), 'Ț', 'T')))) AS SpecializareCurata,
                    ISNULL(sf.DenumireMaterie, 'Nedefinit') AS DenumireMaterie, ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost, ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                    (ISNULL(sf.Nr_Ore_Curs, 0) * ISNULL(sf.CoefOreConvCurs, 1)) + ((ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0)) * ISNULL(sf.CoefOreConvApp, 1)) AS OreConvLinie,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) AS FacultateCurata, UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                ),
                Filtrat AS (
                    SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru, SUM(OreConvLinie) AS TotalOreConvItem FROM BaseData
                    WHERE (@an = 'Toti' OR AnCurat = @an) AND (@fac = 'Toti' OR FacultateCurata = @fac) AND (@prof = 'Toti' OR NumeIntreg = @prof) AND
                          (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ','))) AND (@semestru = 0 OR Semestru = @semestru) AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                    GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru
                ),
                TotalProfesor AS (SELECT NumeIntreg, SUM(TotalOreConvItem) AS TotalOreConvProf FROM Filtrat GROUP BY NumeIntreg)
                SELECT f.NumeIntreg, f.SpecializareCurata, f.DenumireMaterie, f.TipPost, f.Semestru, f.TotalOreConvItem, t.TotalOreConvProf
                FROM Filtrat f INNER JOIN TotalProfesor t ON f.NumeIntreg = t.NumeIntreg ORDER BY f.NumeIntreg, f.DenumireMaterie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti"); cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti"); cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari); cmd.Parameters.AddWithValue("@semestru", semestru); cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    data.Add(new PdfNormaData
                    {
                        Profesor = reader["NumeIntreg"]?.ToString() ?? "",
                        Specializare = reader["SpecializareCurata"]?.ToString() ?? "",
                        Materie = reader["DenumireMaterie"]?.ToString() ?? "",
                        Tip = reader["TipPost"]?.ToString() ?? "",
                        Semestru = reader["Semestru"]?.ToString() ?? "",
                        OreConv = reader["TotalOreConvItem"]?.ToString() ?? "",
                        TotalPost = reader["TotalOreConvProf"]?.ToString() ?? ""
                    });
                }
            }

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape()); page.Margin(2, Unit.Centimetre); page.PageColor(Colors.White); page.DefaultTextStyle(x => x.FontSize(10));

                    page.Header().Column(col =>
                    {
                        col.Item().Text("Raport Norme Profesori").SemiBold().FontSize(18).FontColor(Color.FromHex(BrandColorHex));
                        col.Item().Text($"Generat la: {DateTime.Now:dd/MM/yyyy}").FontSize(10).FontColor(Colors.Grey.Medium);
                        col.Item().PaddingTop(5).Text(text =>
                        {
                            text.Span("Filtre aplicate: ").SemiBold().FontSize(9);
                            text.Span($"An Univ: {anUniv} | Fac: {facultate} | Spec: {specializari} | Prof: {profesor} | Sem: {(semestru == 0 ? "Toate" : semestru.ToString())} | Tip Post: {tipPost}").FontSize(9);
                        });
                    });

                    page.Content().PaddingVertical(1, Unit.Centimetre).Table(table =>
                    {
                        table.ColumnsDefinition(c => { c.RelativeColumn(3); c.RelativeColumn(3); c.RelativeColumn(4); c.RelativeColumn(2); c.ConstantColumn(40); c.ConstantColumn(60); c.ConstantColumn(60); });
                        table.Header(h =>
                        {
                            h.Cell().Element(HeaderStyle).Text("Profesor"); h.Cell().Element(HeaderStyle).Text("Specializare"); h.Cell().Element(HeaderStyle).Text("Materie"); h.Cell().Element(HeaderStyle).Text("Tip Post"); h.Cell().Element(HeaderStyle).Text("Sem"); h.Cell().Element(HeaderStyle).Text("Ore Conv"); h.Cell().Element(HeaderStyle).Text("Total Post");
                            static IContainer HeaderStyle(IContainer c) => c.DefaultTextStyle(x => x.SemiBold().FontColor(Colors.White)).Background(Color.FromHex(BrandColorHex)).PaddingVertical(5).PaddingHorizontal(2);
                        });
                        for (int i = 0; i < data.Count; i++)
                        {
                            var item = data[i]; var bgColor = i % 2 == 0 ? Colors.White : Colors.Grey.Lighten4;
                            table.Cell().Element(c => CellStyle(c, bgColor)).Text(item.Profesor); table.Cell().Element(c => CellStyle(c, bgColor)).Text(item.Specializare); table.Cell().Element(c => CellStyle(c, bgColor)).Text(item.Materie); table.Cell().Element(c => CellStyle(c, bgColor)).Text(item.Tip); table.Cell().Element(c => CellStyle(c, bgColor)).AlignCenter().Text(item.Semestru); table.Cell().Element(c => CellStyle(c, bgColor)).AlignCenter().Text(item.OreConv); table.Cell().Element(c => CellStyle(c, bgColor)).AlignCenter().Text(item.TotalPost).Bold().FontColor(Color.FromHex(BrandColorHex));
                        }
                        static IContainer CellStyle(IContainer c, string bgColor) => c.Background(bgColor).BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(4);
                    });
                    page.Footer().AlignCenter().Text(x => { x.Span("Pagina "); x.CurrentPageNumber(); });
                });
            });

            string fileName = "NormaProfesori_General.pdf";
            if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
            {
                fileName = $"NormaProfesori_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.pdf";
            }
            var stream = new MemoryStream(document.GeneratePdf()); return File(stream.ToArray(), "application/pdf", fileName);
        }

        [HttpGet("export/pdf/ore-program")]
        public async Task<IActionResult> ExportOrePdf(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti")
        {
            QuestPDF.Settings.License = LicenseType.Community;
            var data = new List<PdfOreData>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                await conn.OpenAsync();
                string sql = @"
                WITH BaseData AS (
                    SELECT ppm.NumeIntreg, UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'Ș', 'S'), 'Ț', 'T')))) AS ProgramStudiu,
                    (ISNULL(sf.Nr_Ore_Curs, 0) * ISNULL(sf.CoefOreConvCurs, 1)) + ((ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0)) * ISNULL(sf.CoefOreConvApp, 1)) AS OreConvLinie,
                    ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost, ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) AS FacultateCurata, UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                ),
                Filtrat AS (
                    SELECT NumeIntreg, ProgramStudiu, SUM(OreConvLinie) AS OreConvProgram FROM BaseData
                    WHERE (@AnUniv = 'Toti' OR AnCurat = @AnUniv) AND (@Facultate = 'Toti' OR FacultateCurata = @Facultate) AND (@Profesor = 'Toti' OR NumeIntreg = @Profesor) AND
                          (@Specializari = 'Toti' OR ProgramStudiu IN (SELECT value FROM STRING_SPLIT(@Specializari, ','))) AND (@Semestru = 0 OR Semestru = @Semestru) AND (@TipPost = 'Toti' OR TipPost = @TipPost)
                    GROUP BY NumeIntreg, ProgramStudiu
                ),
                TotalProfesor AS (SELECT NumeIntreg, SUM(OreConvProgram) AS TotalPost FROM Filtrat GROUP BY NumeIntreg)
                SELECT f.NumeIntreg, ISNULL(f.ProgramStudiu, 'Nespecificat') AS ProgramStudiu, f.OreConvProgram, t.TotalPost,
                    CAST(CASE WHEN t.TotalPost = 0 THEN 0 ELSE (f.OreConvProgram / t.TotalPost) * 100 END AS DECIMAL(10,2)) AS ProcentPost
                FROM Filtrat f INNER JOIN TotalProfesor t ON f.NumeIntreg = t.NumeIntreg ORDER BY f.NumeIntreg, f.OreConvProgram DESC";

                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@AnUniv", anUniv ?? "Toti"); cmd.Parameters.AddWithValue("@Facultate", facultate ?? "Toti"); cmd.Parameters.AddWithValue("@Specializari", specializari ?? "Toti");
                cmd.Parameters.AddWithValue("@Profesor", profesor ?? "Toti"); cmd.Parameters.AddWithValue("@Semestru", semestru); cmd.Parameters.AddWithValue("@TipPost", tipPost ?? "Toti");

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    data.Add(new PdfOreData
                    {
                        Profesor = reader["NumeIntreg"]?.ToString() ?? "",
                        Program = reader["ProgramStudiu"]?.ToString() ?? "",
                        OreConv = reader["OreConvProgram"]?.ToString() ?? "",
                        TotalPost = reader["TotalPost"]?.ToString() ?? "",
                        Procent = reader["ProcentPost"]?.ToString() ?? ""
                    });
                }
            }

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4); page.Margin(2, Unit.Centimetre); page.PageColor(Colors.White); page.DefaultTextStyle(x => x.FontSize(10));

                    page.Header().Column(col =>
                    {
                        col.Item().Text("Statistica Ore pe Programe").SemiBold().FontSize(18).FontColor(Color.FromHex(BrandColorHex));
                        col.Item().Text($"Generat la: {DateTime.Now:dd/MM/yyyy}").FontSize(10).FontColor(Colors.Grey.Medium);
                        col.Item().PaddingTop(5).Text(text =>
                        {
                            text.Span("Filtre aplicate: ").SemiBold().FontSize(9);
                            text.Span($"An Univ: {anUniv} | Fac: {facultate} | Spec: {specializari} | Prof: {profesor} | Sem: {(semestru == 0 ? "Toate" : semestru.ToString())} | Tip Post: {tipPost}").FontSize(9);
                        });
                    });

                    page.Content().PaddingVertical(1, Unit.Centimetre).Table(table =>
                    {
                        table.ColumnsDefinition(c => { c.RelativeColumn(3); c.RelativeColumn(4); c.ConstantColumn(50); c.ConstantColumn(50); c.ConstantColumn(60); });
                        table.Header(h =>
                        {
                            h.Cell().Element(HeaderStyle).Text("Profesor"); h.Cell().Element(HeaderStyle).Text("Program Studiu"); h.Cell().Element(HeaderStyle).Text("Ore Conv"); h.Cell().Element(HeaderStyle).Text("% Post"); h.Cell().Element(HeaderStyle).Text("Total Post");
                            static IContainer HeaderStyle(IContainer c) => c.DefaultTextStyle(x => x.SemiBold().FontColor(Colors.White)).Background(Color.FromHex(BrandColorHex)).Padding(5);
                        });
                        for (int i = 0; i < data.Count; i++)
                        {
                            var item = data[i]; var bgColor = i % 2 == 0 ? Colors.White : Colors.Grey.Lighten4;
                            table.Cell().Element(c => CellStyle(c, bgColor)).Text(item.Profesor); table.Cell().Element(c => CellStyle(c, bgColor)).Text(item.Program); table.Cell().Element(c => CellStyle(c, bgColor)).AlignCenter().Text(item.OreConv); table.Cell().Element(c => CellStyle(c, bgColor)).AlignCenter().Text($"{item.Procent}%"); table.Cell().Element(c => CellStyle(c, bgColor)).AlignCenter().Text(item.TotalPost).Bold().FontColor(Color.FromHex(BrandColorHex));
                        }
                        static IContainer CellStyle(IContainer c, string bgColor) => c.Background(bgColor).BorderBottom(1).BorderColor(Colors.Grey.Lighten3).Padding(5);
                    });
                    page.Footer().AlignCenter().Text(x => { x.Span("Pagina "); x.CurrentPageNumber(); });
                });
            });

            string fileName = "StatisticaOre_General.pdf";
            if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
            {
                fileName = $"StatisticaOre_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.pdf";
            }
            var stream = new MemoryStream(document.GeneratePdf()); return File(stream.ToArray(), "application/pdf", fileName);
        }

        [HttpGet("export/raport-ans")]
        public IActionResult ExportRaportANS([FromQuery] int idAnUniv = 45)
        {
            var dateBrute = new List<RandSqlANS>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string query = @"
                    SELECT 
                        NumeIntreg AS NumeComplet, 
                        DenumireGradDidacticPost AS GradFunctie, 
                        ISNULL(NrOreConventionale, 0) AS NrOreConventionale, 
                        DenumireFacultate AS CriteriuMapare 
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie]
                    WHERE ID_AnUniv = @ID_AnUniv AND TitularSauSuplinitor = 1";

                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var numeComplet = reader["NumeComplet"]?.ToString() ?? "";
                            var partiNume = numeComplet.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);

                            dateBrute.Add(new RandSqlANS
                            {
                                Nume = partiNume.Length > 0 ? partiNume[0] : "",
                                Prenume = partiNume.Length > 1 ? partiNume[1] : "",
                                GradFunctie = reader["GradFunctie"]?.ToString() ?? "",
                                OreConventionale = reader["NrOreConventionale"] != DBNull.Value ? Convert.ToDecimal(reader["NrOreConventionale"]) : 0,
                                DomeniuANS = AsociazaDomeniuANS(reader["CriteriuMapare"]?.ToString() ?? "")
                            });
                        }
                    }
                }
            }

            var profesoriGrupati = dateBrute
                .GroupBy(x => new { x.Nume, x.Prenume, x.GradFunctie })
                .Select(g => new
                {
                    g.Key.Nume,
                    g.Key.Prenume,
                    g.Key.GradFunctie,
                    NormaBaza = ObtineNormaBaza(g.Key.GradFunctie),
                    Domenii = g.GroupBy(d => d.DomeniuANS)
                               .Select(dg => new { NumeDomeniu = dg.Key, TotalOreConv = dg.Sum(x => x.OreConventionale) })
                               .OrderByDescending(d => d.TotalOreConv).ToList()
                })
                .OrderBy(p => p.Nume).ThenBy(p => p.Prenume).ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Date ANS");
                worksheet.Cell(5, 1).Value = "Nr. crt."; worksheet.Cell(5, 2).Value = "Nume"; worksheet.Cell(5, 3).Value = "Prenume";
                worksheet.Cell(5, 4).Value = "Grad / funcție"; worksheet.Cell(5, 5).Value = "Total normă conform contract";

                for (int i = 1; i <= 10; i++)
                {
                    worksheet.Cell(5, 4 + (i * 2) - 1).Value = $"Domeniul ANS {i}";
                    worksheet.Cell(5, 4 + (i * 2)).Value = $"Fracțiune normă dom. ANS {i}";
                }

                worksheet.Range(5, 1, 5, 25).Style.Font.Bold = true;
                int randCurent = 6; int nrCrt = 1;

                foreach (var prof in profesoriGrupati)
                {
                    worksheet.Cell(randCurent, 1).Value = nrCrt++; worksheet.Cell(randCurent, 2).Value = prof.Nume;
                    worksheet.Cell(randCurent, 3).Value = prof.Prenume; worksheet.Cell(randCurent, 4).Value = prof.GradFunctie;
                    worksheet.Cell(randCurent, 5).Value = 1;

                    int coloanaStartDomeniu = 6;
                    for (int i = 0; i < Math.Min(10, prof.Domenii.Count); i++)
                    {
                        var domeniu = prof.Domenii[i];
                        decimal fractiune = prof.NormaBaza > 0 ? domeniu.TotalOreConv / prof.NormaBaza : 0;
                        worksheet.Cell(randCurent, coloanaStartDomeniu).Value = domeniu.NumeDomeniu;
                        worksheet.Cell(randCurent, coloanaStartDomeniu + 1).Value = Math.Round(fractiune, 3);
                        worksheet.Cell(randCurent, coloanaStartDomeniu + 1).Style.NumberFormat.Format = "0.000";
                        coloanaStartDomeniu += 2;
                    }
                    randCurent++;
                }

                worksheet.Columns().AdjustToContents();
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Date_ANS.xlsx");
                }
            }
        }

        private decimal ObtineNormaBaza(string grad)
        {
            if (string.IsNullOrWhiteSpace(grad)) return 0;
            string gradLow = grad.ToLower();
            if (gradLow.Contains("profesor")) return 11;
            if (gradLow.Contains("conferentiar") || gradLow.Contains("conferențiar")) return 12;
            if (gradLow.Contains("lector") || gradLow.Contains("sef lucrari") || gradLow.Contains("șef lucrări")) return 14;
            if (gradLow.Contains("asistent")) return 16;
            return 0;
        }

        private string AsociazaDomeniuANS(string criteriu)
        {
            if (string.IsNullOrWhiteSpace(criteriu)) return "Ştiinţe economice (fără  Cibernetică, statistică şi informatică economică)";
            string val = criteriu.ToLower();
            if (val.Contains("matematic") || val.Contains("informatic")) return "Matematică";
            if (val.Contains("fizic")) return "Fizică";
            if (val.Contains("chimie") || val.Contains("inginerie chimic")) return "Chimie şi inginerie chimică";
            if (val.Contains("civil") || val.Contains("construct")) return "Inginerie civilă";
            if (val.Contains("electric") || val.Contains("electronic") || val.Contains("telecomunica")) return "Inginerie electrică, electronică şi telecomunicaţii";
            if (val.Contains("transport")) return "Ingineria transporturilor";
            if (val.Contains("silvicultur") || val.Contains("lemn")) return "Ingineria resurselor vegetale şi animale";
            if (val.Contains("sistem") || val.Contains("calculatoare")) return "Ingineria sistemelor, calculatoare şi tehnologia informaţiei";
            if (val.Contains("mecanic") || val.Contains("mecatronic") || val.Contains("industrial")) return "Inginerie mecanică, mecatronică, inginerie industrială şi management";
            if (val.Contains("medicin")) return "Medicină";
            if (val.Contains("drept") || val.Contains("juridic")) return "Ştiinţe juridice";
            if (val.Contains("sociologie") || val.Contains("comunicar")) return "Sociologie";
            if (val.Contains("economic") || val.Contains("business")) return "Ştiinţe economice (fără  Cibernetică, statistică şi informatică economică)";
            if (val.Contains("psihologie") || val.Contains("educa")) return "Psihologie şi ştiinţe comportamentale";
            if (val.Contains("liter") || val.Contains("filologie")) return "Filologie";
            if (val.Contains("sport") || val.Contains("educatie fizic")) return "Ştiinţele Sportului şi Educaţiei Fizice";
            if (val.Contains("muzic")) return "Muzică (fără Interpretare muzicală)";
            return "Studii culturale";
        }

        private class RandSqlANS
        {
            public string Nume { get; set; } = "";
            public string Prenume { get; set; } = "";
            public string GradFunctie { get; set; } = "";
            public decimal OreConventionale { get; set; }
            public string DomeniuANS { get; set; } = "";
        }
    }
}
