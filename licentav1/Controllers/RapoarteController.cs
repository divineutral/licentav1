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
                        ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
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
                    ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
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
                        ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
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
                    ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
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
                    ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
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
                    ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
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


        // Returneaz? subdomeniul ANS pentru un id_metaspecializare dat.
        // Dac? nu exist? mapping, returneaz? 0 (record exclus din raport).

        [HttpGet("export/raport-ans")]
        public IActionResult ExportRaportANS([FromQuery] int idAnUniv = 45)
        {
            // Mapping id_metaspecializare -> ID subdomeniu ANS
            var mappingMetaspec = new Dictionary<int, int>
            {
                // Constructii -> 5 Ing. civila
                { 20, 5 }, { 34, 5 }, { 44, 5 }, { 182, 5 }, { 837, 5 }, { 847, 5 }, { 848, 5 },
                // Ing. Mecanica -> 11, transport -> 8
                { 43, 11 }, { 148, 8 }, { 171, 11 }, { 306, 11 }, { 358, 11 }, { 464, 11 },
                { 466, 11 }, { 615, 11 }, { 823, 11 }, { 828, 11 },
                // IESC -> 6 Ing. electrica
                { 35, 6 }, { 82, 6 }, { 84, 6 }, { 139, 11 }, { 162, 6 }, { 310, 6 },
                { 315, 6 }, { 316, 6 }, { 362, 6 }, { 372, 6 }, { 418, 6 }, { 597, 6 }, { 617, 6 },
                // Ing. Tehnologica + Mat. -> 11
                { 116, 11 }, { 126, 11 }, { 129, 11 }, { 156, 11 }, { 228, 11 }, { 229, 11 },
                { 404, 11 }, { 529, 11 }, { 142, 11 }, { 326, 11 }, { 531, 11 }, { 806, 11 }, { 819, 11 },
                // Design Produs -> 11
                { 53, 11 }, { 138, 11 }, { 446, 11 }, { 448, 11 }, { 449, 11 }, { 450, 11 },
                { 451, 11 }, { 496, 11 }, { 497, 11 }, { 821, 11 },
                // Alimentatie + Silvicultura -> 9 Ing. res. vegetale
                { 46, 9 }, { 122, 9 }, { 176, 9 }, { 178, 9 }, { 731, 9 },
                { 72, 9 }, { 90, 9 }, { 118, 9 }, { 226, 9 }, { 235, 9 }, { 249, 9 },
                { 307, 9 }, { 437, 9 }, { 458, 9 }, { 566, 9 }, { 845, 9 },
                // Matematica si Informatica
                { 101, 40 }, { 102, 40 }, { 104, 40 }, { 251, 1 }, { 317, 40 },
                { 340, 1 }, { 368, 40 }, { 369, 40 }, { 477, 40 },
                // Economie -> 25, Cibernetica -> 24
                { 45, 25 }, { 73, 25 }, { 93, 25 }, { 112, 24 }, { 221, 25 }, { 223, 25 },
                { 227, 25 }, { 242, 25 }, { 283, 25 }, { 288, 25 }, { 299, 25 },
                // Litere -> 27 Filologie, Inovare Culturala -> 31
                { 181, 31 }, { 196, 27 }, { 197, 27 }, { 200, 27 }, { 205, 27 }, { 207, 27 },
                { 209, 27 }, { 217, 27 }, { 218, 27 }, { 341, 27 }, { 343, 27 }, { 463, 27 },
                { 511, 27 }, { 512, 27 }, { 513, 27 }, { 514, 27 }, { 726, 27 }, { 798, 27 }, { 801, 27 },
                // Drept -> 18
                { 60, 18 }, { 64, 18 }, { 331, 18 }, { 485, 18 }, { 555, 18 },
                // Sociologie -> 21, Comunicare -> 20, Resurse Umane -> 25
                { 41, 20 }, { 98, 20 }, { 100, 25 }, { 322, 21 }, { 524, 25 }, { 579, 20 }, { 831, 20 },
                // Psihologie + DPPD -> 26
                { 276, 26 }, { 294, 26 }, { 296, 26 }, { 383, 26 }, { 384, 26 }, { 416, 26 },
                { 515, 26 }, { 813, 26 }, { 851, 26 }, { 832, 26 }, { 834, 26 },
                // Medicina -> 14
                { 394, 14 }, { 397, 14 }, { 402, 14 }, { 484, 14 }, { 585, 14 }, { 594, 14 }, { 835, 14 },
                // Muzica -> 37 (Muzica doar Interpretare muzicala, col 47)
                { 186, 37 }, { 187, 37 }, { 264, 37 }, { 332, 37 }, { 351, 37 }, { 557, 37 }, { 838, 37 },
                // Educatie Fizica -> 39 (Stiintele Sportului, col 49)
                { 78, 39 }, { 189, 39 }, { 325, 39 }, { 470, 39 }, { 783, 39 }, { 784, 39 }, { 846, 39 },
            };

            // ID ANS -> coloana Excel (1-based) - ordinea exacta din template Diana Ionita
            var ansIdToCol = new Dictionary<int, int>
            {
                { 1, 10 }, { 40, 11 }, { 2, 12 }, { 3, 13 }, { 4, 14 },   // Matematica si stiinte
                { 5, 15 }, { 6, 16 }, { 7, 17 }, { 8, 18 }, { 9, 19 }, { 10, 20 }, { 11, 21 }, // Inginerie
                { 12, 22 }, { 13, 23 }, { 14, 24 }, { 15, 25 }, { 16, 26 }, { 17, 27 },         // Biologice
                { 18, 28 }, { 19, 29 }, { 20, 30 }, { 21, 31 }, { 22, 32 }, { 23, 33 },          // Sociale
                { 24, 34 }, { 25, 35 }, { 26, 36 },
                { 27, 37 }, { 28, 38 }, { 29, 39 }, { 30, 40 }, { 31, 41 }, { 32, 42 },          // Umaniste
                { 33, 43 }, { 34, 44 }, { 35, 45 }, { 36, 46 }, { 37, 47 }, { 38, 48 }, { 39, 49 },
            };

            // Citeste date din DB
            var dateBrute = new List<RandSqlANS>();
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string query = @"
                    SELECT
                        ppm.NumeIntreg                              AS NumeComplet,
                        ppm.DenumireGradDidacticPost                AS GradFunctie,
                        ISNULL(ppm.DenumireCatedra, '')             AS Departament,
                        ISNULL(ppm.DenumireFacultate, '')           AS Facultate,
                        ISNULL(sf.NrOreConventionale, 0)            AS OreConventionale,
                        sf.id_metaspecializare                      AS IdMetaspec
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                        ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    WHERE sf.id_anuniv = @ID_AnUniv
                      AND sf.TitularSauSuplinitor = 1";

                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int idMeta = reader["IdMetaspec"] != DBNull.Value
                                ? Convert.ToInt32(reader["IdMetaspec"]) : 0;
                            int idAns;
                            if (!mappingMetaspec.TryGetValue(idMeta, out idAns)) continue;
                            if (!ansIdToCol.ContainsKey(idAns)) continue;

                            dateBrute.Add(new RandSqlANS
                            {
                                NumeComplet = reader["NumeComplet"]?.ToString() ?? "",
                                Departament = reader["Departament"]?.ToString() ?? "",
                                Facultate = reader["Facultate"]?.ToString() ?? "",
                                GradFunctie = reader["GradFunctie"]?.ToString() ?? "",
                                OreConventionale = reader["OreConventionale"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["OreConventionale"]) : 0m,
                                IdANS = idAns,
                            });
                        }
                    }
                }
            }

            // Grupeaza pe profesor: ore per coloana ANS
            var profesori = new List<ProfANS>();
            foreach (var grp in dateBrute.GroupBy(x => new { x.NumeComplet, x.Departament, x.Facultate, x.GradFunctie }))
            {
                decimal normaBaza = ObtineNormaBaza(grp.Key.GradFunctie);
                var orePerCol = new Dictionary<int, decimal>();
                foreach (var rand in grp)
                {
                    int col = ansIdToCol[rand.IdANS];
                    if (!orePerCol.ContainsKey(col)) orePerCol[col] = 0m;
                    orePerCol[col] += rand.OreConventionale;
                }
                // Converteste ore in fractiuni de norma
                var fractiuni = new Dictionary<int, decimal>();
                foreach (var kv in orePerCol)
                    fractiuni[kv.Key] = normaBaza > 0 ? Math.Round(kv.Value / normaBaza, 3) : 0m;

                profesori.Add(new ProfANS
                {
                    NumeComplet = grp.Key.NumeComplet,
                    Departament = grp.Key.Departament,
                    Facultate = grp.Key.Facultate,
                    GradFunctie = grp.Key.GradFunctie,
                    Fractiuni = fractiuni,
                });
            }
            profesori = profesori.OrderBy(p => p.NumeComplet).ToList();

            // Construieste Excel pe baza template-ului Diana
            var templatePath = System.IO.Path.Combine(
                System.AppDomain.CurrentDomain.BaseDirectory, "Templates", "Date_ANS_Diana_Ionita.xlsx");

            ClosedXML.Excel.XLWorkbook workbook;

            // Daca template-ul exista, il folosim ca baza; altfel construim de la zero
            if (System.IO.File.Exists(templatePath))
            {
                workbook = new ClosedXML.Excel.XLWorkbook(templatePath);
            }
            else
            {
                workbook = BuildANSWorkbookFromScratch();
            }

            var ws = workbook.Worksheets.First();

            // Determina randul de start al datelor (primul rand gol dupa headerele cu note)
            // In template: R1=gol, R2=Anexa, R3=Universitatea, R4=NOTE, R5=header1, R6=header2, R7=header3, R8=litere, R9+=date
            int dataStartRow = 9;

            // Sterge randurile de date existente (randuri demo din template)
            // si pastreaza doar un rand "Total general" la final
            // Gaseste randul Total general
            int totalRow = -1;
            for (int r = dataStartRow; r <= ws.LastRowUsed().RowNumber(); r++)
            {
                var cellVal = ws.Cell(r, 1).GetString();
                if (cellVal != null && cellVal.Contains("Total"))
                {
                    totalRow = r;
                    break;
                }
            }

            // Sterge randurile cu date demo (pastrand Total)
            if (totalRow > dataStartRow)
            {
                for (int r = totalRow - 1; r >= dataStartRow; r--)
                    ws.Row(r).Delete();
                totalRow = dataStartRow; // dupa stergere, Total e acum pe dataStartRow
            }
            else if (totalRow == -1)
            {
                totalRow = dataStartRow;
            }

            // Insereaza randuri noi pentru profesori inainte de randul Total
            if (profesori.Count > 0)
                ws.Row(totalRow).InsertRowsAbove(profesori.Count);

            // Completeaza datele profesorilor
            for (int i = 0; i < profesori.Count; i++)
            {
                var prof = profesori[i];
                int r = dataStartRow + i;

                ws.Cell(r, 1).Value = i + 1;                    // Nr. Crt.
                ws.Cell(r, 2).Value = prof.NumeComplet;          // Nume si prenume
                ws.Cell(r, 3).Value = "";                        // CNP - nu avem
                ws.Cell(r, 4).Value = prof.GradFunctie;          // Functie
                ws.Cell(r, 5).Value = 1;                         // Forma angajare: 1=norma intreaga
                ws.Cell(r, 6).Value = "";                        // Conducator doctorat - nu avem
                ws.Cell(r, 7).Value = "";                        // Varsta - nu avem
                ws.Cell(r, 8).Value = prof.Facultate;            // Facultate
                ws.Cell(r, 9).Value = prof.Departament;          // Departament

                // Fractiuni per subdomeniu ANS (col 10-49)
                foreach (var kv in prof.Fractiuni)
                    ws.Cell(r, kv.Key).Value = kv.Value;

                // Total (col 50) = SUM(J:AW)
                string rowStr = r.ToString();
                ws.Cell(r, 50).FormulaA1 = $"=SUM(J{rowStr}:AW{rowStr})";

                // Stil zebra
                if (i % 2 != 0)
                {
                    var rowRange = ws.Range(r, 1, r, 50);
                    rowRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#f2f7ef");
                }
            }

            // Actualizeaza formule Total general (ultimul rand)
            int newTotalRow = dataStartRow + profesori.Count;
            for (int c = 10; c <= 49; c++)
            {
                string colLetter = ColumnLetter(c);
                ws.Cell(newTotalRow, c).FormulaA1 =
                    $"=SUM({colLetter}{dataStartRow}:{colLetter}{newTotalRow - 1})";
            }
            ws.Cell(newTotalRow, 50).FormulaA1 =
                $"=SUM(J{newTotalRow}:AW{newTotalRow})";

            using (var stream = new System.IO.MemoryStream())
            {
                workbook.SaveAs(stream);
                workbook.Dispose();
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Date_ANS_{idAnUniv}.xlsx");
            }
        }

        private static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + col % 26) + result;
                col /= 26;
            }
            return result;
        }

        private ClosedXML.Excel.XLWorkbook BuildANSWorkbookFromScratch()
        {
            var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add("CD DRU");

            // R2: Titlu
            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, 50).Merge();

            // R3: Universitate
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov";
            ws.Range(3, 1, 3, 6).Merge();

            // R4: Note
            ws.Cell(4, 1).Value = "NOTA: Se includ in tabel toate cadrele didactice si de cercetare titulare (cu norma de baza in universitate), indiferent de forma de angajare.";
            ws.Range(4, 1, 4, 9).Merge();
            ws.Cell(4, 10).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta.";
            ws.Range(4, 10, 4, 27).Merge();
            ws.Cell(4, 28).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta.";
            ws.Range(4, 28, 4, 50).Merge();

            // R5: Header nivel 1 - coloane fixe + domenii fundamentale
            ws.Cell(5, 1).Value = "Nr. \nCrt.";
            ws.Cell(5, 2).Value = "Nume si prenume cadru didactic";
            ws.Cell(5, 3).Value = "CNP";
            ws.Cell(5, 4).Value = "Functie cadru didactic sau cercetare";
            ws.Cell(5, 5).Value = "Forma de angajare";
            ws.Cell(5, 6).Value = "Calitate conducator doctorat";
            ws.Cell(5, 7).Value = "Varsta";
            ws.Cell(5, 8).Value = "Facultate";
            ws.Cell(5, 9).Value = "Departament";
            ws.Cell(5, 10).Value = "Matematica si stiinte ale naturii";
            ws.Cell(5, 15).Value = "Stiinte ingineresti";
            ws.Cell(5, 22).Value = "Stiinte biologice si biomedicale";
            ws.Cell(5, 28).Value = "Stiinte sociale";
            ws.Cell(5, 37).Value = "Stiinte umaniste si arte";
            ws.Cell(5, 50).Value = "Total";

            // Merge domenii fundamentale in R5
            ws.Range(5, 1, 7, 1).Merge();   // Nr.Crt merge R5:R7
            ws.Range(5, 2, 7, 2).Merge();
            ws.Range(5, 3, 7, 3).Merge();
            ws.Range(5, 4, 7, 4).Merge();
            ws.Range(5, 5, 7, 5).Merge();
            ws.Range(5, 6, 7, 6).Merge();
            ws.Range(5, 7, 7, 7).Merge();
            ws.Range(5, 8, 7, 8).Merge();
            ws.Range(5, 9, 7, 9).Merge();
            ws.Range(5, 10, 5, 14).Merge();  // Matematica
            ws.Range(5, 15, 5, 21).Merge();  // Inginerie
            ws.Range(5, 22, 5, 27).Merge();  // Biologice
            ws.Range(5, 28, 5, 36).Merge();  // Sociale
            ws.Range(5, 37, 5, 49).Merge();  // Umaniste
            ws.Range(5, 50, 7, 50).Merge();  // Total

            // R6: Subdomenii
            string[] subdomenii = {
                "Matematica", "Informatica", "Fizica", "Chimie si inginerie chimica", "Stiintele pamantului si atmosferei",
                "Inginerie civila", "Inginerie electrica, electronica si telecomunicatii", "Inginerie geologica, mine, petrol si gaze",
                "Ingineria transporturilor", "Ingineria resurselor vegetale si animale",
                "Ingineria sistemelor, calculatoare si tehnologia informatiei",
                "Inginerie mecanica, mecatronica, inginerie industriala si management",
                "Biologie", "Biochimie", "Medicina", "Medicina veterinara", "Medicina dentara", "Farmacie",
                "Stiinte juridice", "Stiinte administrative", "Stiinte ale comunicarii", "Sociologie",
                "Stiinte politice", "Stiinte militare, informatii si ordine publica",
                "Stiinte economice (doar Cibernetica, statistica si informatica economica)",
                "Stiinte economice (fara Cibernetica, statistica si informatica economica)",
                "Psihologie si stiinte comportamentale",
                "Filologie", "Filosofie", "Istorie", "Teologie", "Studii culturale",
                "Arhitectura si urbanism", "Arte vizuale (fara Istoria si teoria artei)",
                "Arte vizuale (doar Istoria si teoria artei)", "Teatru si artele spectacolului",
                "Cinematografie si media", "Muzica (doar Interpretare muzicala)",
                "Muzica (fara Interpretare muzicala)", "Stiintele Sportului si Educatiei Fizice"
            };
            for (int i = 0; i < subdomenii.Length; i++)
            {
                ws.Cell(6, 10 + i).Value = subdomenii[i];
                ws.Range(6, 10 + i, 7, 10 + i).Merge();
            }

            // R8: Litere index
            string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            for (int i = 0; i < 9; i++) ws.Cell(8, i + 1).Value = (char)('A' + i);
            for (int i = 0; i < 41; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, 50).Value = "40";

            // Stiluri header
            var headerFill = ClosedXML.Excel.XLColor.FromHtml("#1F4E79");
            for (int r = 5; r <= 8; r++)
                for (int c = 1; c <= 50; c++)
                {
                    ws.Cell(r, c).Style.Font.Bold = true;
                    ws.Cell(r, c).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                    ws.Cell(r, c).Style.Fill.BackgroundColor = headerFill;
                    ws.Cell(r, c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                    ws.Cell(r, c).Style.Alignment.WrapText = true;
                    ws.Cell(r, c).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                }

            // Dimensiuni coloane
            ws.Column(1).Width = 5;
            ws.Column(2).Width = 30;
            ws.Column(3).Width = 14;
            ws.Column(4).Width = 22;
            ws.Column(5).Width = 10;
            ws.Column(6).Width = 12;
            ws.Column(7).Width = 8;
            ws.Column(8).Width = 28;
            ws.Column(9).Width = 28;
            for (int c = 10; c <= 50; c++) ws.Column(c).Width = 12;

            return wb;
        }

        private class ProfANS
        {
            public string NumeComplet { get; set; } = "";
            public string Departament { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string GradFunctie { get; set; } = "";
            public Dictionary<int, decimal> Fractiuni { get; set; } = new Dictionary<int, decimal>();
        }


        private decimal ObtineNormaBaza(string grad)
        {
            if (string.IsNullOrWhiteSpace(grad)) return 0;
            string gradLow = grad.ToLower();
            if (gradLow.Contains("profesor")) return 11;
            if (gradLow.Contains("conferentiar") || gradLow.Contains("conferențiar")) return 12;
            if (gradLow.Contains("lector") || gradLow.Contains("sef lucrari") || gradLow.Contains("șef lucrări") || gradLow.Contains("șef de lucrări")) return 14;
            if (gradLow.Contains("asistent")) return 16;
            return 0;
        }

        private class RandSqlANS
        {
            public string NumeComplet { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string Departament { get; set; } = "";
            public string GradFunctie { get; set; } = "";
            public decimal OreConventionale { get; set; }
            public int IdANS { get; set; }
        }

    }
}