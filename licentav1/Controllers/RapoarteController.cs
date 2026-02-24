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

        // 40 subdomenii ANS fixe - ordinea coloanelor e identica cu fisierul Diana Ionita
        private static readonly string[] _domeniiANS = new[]
        {
            "Matematică",
            "Informatică",
            "Fizică",
            "Chimie şi inginerie chimică",
            "Ştiinţele pământului şi atmosferei",
            "Inginerie civilă",
            "Inginerie electrică, electronică şi telecomunicaţii",
            "Inginerie geologică, mine, petrol şi gaze",
            "Ingineria transporturilor",
            "Ingineria resurselor vegetale şi animale",
            "Ingineria sistemelor, calculatoare şi tehnologia informaţiei",
            "Inginerie mecanică, mecatronică, inginerie industrială şi management",
            "Biologie",
            "Biochimie",
            "Medicină",
            "Medicină veterinară",
            "Medicină dentară",
            "Farmacie",
            "Ştiinţe juridice",
            "Ştiinţe administrative",
            "Ştiinţe ale comunicării",
            "Sociologie",
            "Ştiinţe politice",
            "Ştiinţe militare, informaţii şi ordine publică",
            "Ştiinţe economice (doar Cibernetică, statistică şi informatică economică)",
            "Ştiinţe economice (fără  Cibernetică, statistică şi informatică economică)",
            "Psihologie şi ştiinţe comportamentale",
            "Filologie",
            "Filosofie",
            "Istorie",
            "Teologie",
            "Studii culturale",
            "Arhitectură şi urbanism",
            "Arte vizuale (fără Istoria şi teoria artei)",
            "Arte vizuale (doar Istoria şi teoria artei)",
            "Teatru şi artele spectacolului",
            "Cinematografie şi media",
            "Muzică (doar Interpretare muzicală)",
            "Muzică (fără Interpretare muzicală)",
            "Ştiinţele Sportului şi Educaţiei Fizice"
        };

        [HttpGet("export/raport-ans")]
        public IActionResult ExportRaportANS([FromQuery] int idAnUniv = 44)
        {
            var dateBrute = new List<RandSqlANS>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                // Query corectat: adaugam Facultatea, folosim NrOreConventionale direct (nu calculam)
                string query = @"
                    SELECT 
                        ppm.NumeIntreg AS NumeComplet,
                        ISNULL(ppm.DenumireGradDidacticPost, '') AS GradFunctie,
                        ISNULL(ppm.DenumireFacultate, 'Nespecificat') AS Facultate,
                        ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                        ISNULL(sf.NrOreConventionale, 0) AS OreConventionale,
                        ISNULL(rsa.Denumire, 'Nedefinit') AS SubdomeniuANS
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm 
                        ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN [AGSIS].[dbo].[N_RAMURA_STIINTA_ANS] rsa 
                        ON sf.id_metaspecializare = rsa.ID_Element
                    WHERE sf.id_anuniv = @ID_AnUniv 
                      AND sf.TitularSauSuplinitor = 1";

                using var command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);

                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var numeComplet = reader["NumeComplet"]?.ToString() ?? "";
                    var parti = numeComplet.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);

                    dateBrute.Add(new RandSqlANS
                    {
                        Nume = parti.Length > 0 ? parti[0] : "",
                        Prenume = parti.Length > 1 ? parti[1] : "",
                        Facultate = reader["Facultate"]?.ToString() ?? "",
                        Departament = reader["Departament"]?.ToString() ?? "",
                        GradFunctie = reader["GradFunctie"]?.ToString() ?? "",
                        OreConventionale = reader["OreConventionale"] != DBNull.Value
                            ? Convert.ToDecimal(reader["OreConventionale"]) : 0,
                        DomeniuANS = reader["SubdomeniuANS"]?.ToString() ?? "Nedefinit"
                    });
                }
            }

            // Grupare pe profesor; calculare fractiuni pentru fiecare din cele 40 domenii fixe
            var profesori = dateBrute
                .GroupBy(x => new { x.Nume, x.Prenume, x.Facultate, x.Departament, x.GradFunctie })
                .Select(g =>
                {
                    decimal normaBaza = ObtineNormaBaza(g.Key.GradFunctie);
                    var orePerDomeniu = g
                        .GroupBy(r => r.DomeniuANS)
                        .ToDictionary(dg => dg.Key, dg => dg.Sum(r => r.OreConventionale));

                    var fractiuni = new decimal[40];
                    for (int i = 0; i < _domeniiANS.Length; i++)
                    {
                        if (orePerDomeniu.TryGetValue(_domeniiANS[i], out decimal ore) && normaBaza > 0)
                            fractiuni[i] = Math.Round(ore / normaBaza, 2);
                    }

                    return new
                    {
                        g.Key.Nume,
                        g.Key.Prenume,
                        g.Key.Facultate,
                        g.Key.Departament,
                        g.Key.GradFunctie,
                        NormaBaza = normaBaza,
                        Fractiuni = fractiuni
                    };
                })
                .OrderBy(p => p.Facultate).ThenBy(p => p.Departament).ThenBy(p => p.Nume).ThenBy(p => p.Prenume)
                .ToList();

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("CD DRU");

            // ── HELPER: get Excel column letter from 1-based index ──────────────
            static string ColLetter(int col)
            {
                if (col <= 26) return ((char)('A' + col - 1)).ToString();
                return ((char)('A' + (col - 1) / 26 - 1)).ToString() + ((char)('A' + (col - 1) % 26)).ToString();
            }

            // ── RÂND 1: gol (height 8.25 ca în template) ────────────────────────
            ws.Row(1).Height = 8.25;

            // ── RÂND 2: titlu Anexa 1 ────────────────────────────────────────────
            ws.Cell(2, 1).Value = "\nAnexa 1. Tabel instituţional privind normarea şi activitatea de cercetare a cadrelor didactice şi de cercetare din universitate (raportare IC2015)";
            ws.Cell(2, 1).Style.Font.Bold = true;
            ws.Cell(2, 1).Style.Font.FontName = "Times New Roman";
            ws.Cell(2, 1).Style.Font.FontSize = 11;
            ws.Cell(2, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            ws.Row(2).Height = 18.0;

            // ── RÂND 3: Universitatea (merge A3:F3) ──────────────────────────────
            ws.Cell(3, 1).Value = "Universitatea TRANSILVANIA DIN BRAŞOV";
            ws.Cell(3, 1).Style.Font.FontName = "Times New Roman";
            ws.Cell(3, 1).Style.Font.FontSize = 11;
            ws.Range(3, 1, 3, 6).Merge();
            ws.Row(3).Height = 15.75;

            // ── RÂND 4: Note mari + header grupuri ANS ───────────────────────────
            // Col A4:D4 merge - nota stanga
            ws.Cell(4, 1).Value = "NOTĂ: \nSe includ în tabel toate cadrele didactice şi de cercetare titulare (inclusiv cadrele didactice angajate cu normă întreagă, cu un contract pe perioadă determinată conform art.294, din LEN 1/2011, valid în perioada de raportare). Pentru facilitarea verificărilor interne recomandăm gruparea pe facultăţi, respectiv departamente. \nFiecare cadru didactic sau de cercetare al universităţii se raportează pe un singur rând.\nCompletarea în câmpurile aferente col.D-F din tabel se realizează prin selectarea valorii corespunzatoare din lista predefinita in col.D, respectiv completarea cu numarul corespunzator valorii din listele predefinite in col.E si col.F.\nVă rugăm să completați numai spațiile marcate cu culoarea galben.";
            ws.Cell(4, 1).Style.Font.FontName = "Times New Roman";
            ws.Cell(4, 1).Style.Font.FontSize = 8;
            ws.Cell(4, 1).Style.Alignment.WrapText = true;
            ws.Cell(4, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            ws.Range(4, 1, 4, 4).Merge();

            // Col E4:F4 merge
            ws.Range(4, 5, 4, 6).Merge();

            // Col J4:AA4 merge - grupuri Matematica + Inginerie
            ws.Cell(4, 10).Value = "Matematică şi ştiinţe ale naturii / Ştiinţe inginereşti";
            ws.Cell(4, 10).Style.Font.Bold = true;
            ws.Cell(4, 10).Style.Font.FontName = "Times New Roman";
            ws.Cell(4, 10).Style.Font.FontSize = 8;
            ws.Cell(4, 10).Style.Alignment.WrapText = true;
            ws.Cell(4, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(4, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range(4, 10, 4, 27).Merge(); // J4:AA4

            // Col AB4:AX4 merge - celelalte grupuri
            ws.Cell(4, 28).Value = "Ştiinţe sociale / Ştiinţe umaniste şi arte";
            ws.Cell(4, 28).Style.Font.Bold = true;
            ws.Cell(4, 28).Style.Font.FontName = "Times New Roman";
            ws.Cell(4, 28).Style.Font.FontSize = 8;
            ws.Cell(4, 28).Style.Alignment.WrapText = true;
            ws.Cell(4, 28).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(4, 28).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range(4, 28, 4, 50).Merge(); // AB4:AX4

            ws.Row(4).Height = 183.75;

            // ── RÂND 5: Headers coloane individuale (A5:A7, B5:B7 etc.) + grupuri domenii ──
            // Merge-uri coloane individuale pe 3 rânduri (5-7): A,B,C,D,E,F,G,H,I
            string[] hdr5 = { "Nr. \nCrt.", "Nume si prenume cadru didactic", "CNP",
                               "Funcţie cadru didactic sau cercetare", "Forma de angajare",
                               "Calitate conducator doctorat", "Varsta", "Facultate", "Departament" };
            for (int c = 1; c <= 9; c++)
            {
                ws.Cell(5, c).Value = hdr5[c - 1];
                ws.Cell(5, c).Style.Font.Bold = true;
                ws.Cell(5, c).Style.Font.FontName = "Times New Roman";
                ws.Cell(5, c).Style.Font.FontSize = 7;
                ws.Cell(5, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(5, c).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Cell(5, c).Style.Alignment.WrapText = true;
                ws.Range(5, c, 7, c).Merge(); // fiecare coloana merge pe randurile 5-7
            }

            // Grupuri domenii ANS pe rândul 5 (merge cu subdomenii pe rândul 6-7)
            // J5:N5 = Matematică şi ştiinţe ale naturii (col 10-14)
            ws.Cell(5, 10).Value = "Matematică şi ştiinţe ale naturii";
            ws.Cell(5, 10).Style.Font.Bold = true;
            ws.Cell(5, 10).Style.Font.FontName = "Times New Roman";
            ws.Cell(5, 10).Style.Font.FontSize = 8;
            ws.Cell(5, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(5, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(5, 10).Style.Alignment.WrapText = true;
            ws.Range(5, 10, 5, 14).Merge();

            // O5:U5 = Ştiinţe inginereşti (col 15-21)
            ws.Cell(5, 15).Value = "Ştiinţe inginereşti";
            ws.Cell(5, 15).Style.Font.Bold = true;
            ws.Cell(5, 15).Style.Font.FontName = "Times New Roman";
            ws.Cell(5, 15).Style.Font.FontSize = 8;
            ws.Cell(5, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(5, 15).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(5, 15).Style.Alignment.WrapText = true;
            ws.Range(5, 15, 5, 21).Merge();

            // V5:AA5 = Ştiinţe biologice şi biomedicale (col 22-27)
            ws.Cell(5, 22).Value = "Ştiinţe biologice şi biomedicale";
            ws.Cell(5, 22).Style.Font.Bold = true;
            ws.Cell(5, 22).Style.Font.FontName = "Times New Roman";
            ws.Cell(5, 22).Style.Font.FontSize = 8;
            ws.Cell(5, 22).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(5, 22).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(5, 22).Style.Alignment.WrapText = true;
            ws.Range(5, 22, 5, 27).Merge();

            // AB5:AJ5 = Ştiinţe sociale (col 28-36)
            ws.Cell(5, 28).Value = "Ştiinţe sociale";
            ws.Cell(5, 28).Style.Font.Bold = true;
            ws.Cell(5, 28).Style.Font.FontName = "Times New Roman";
            ws.Cell(5, 28).Style.Font.FontSize = 8;
            ws.Cell(5, 28).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(5, 28).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(5, 28).Style.Alignment.WrapText = true;
            ws.Range(5, 28, 5, 36).Merge();

            // AK5:AW5 = Ştiinţe umaniste şi arte (col 37-49)
            ws.Cell(5, 37).Value = "Ştiinţe umaniste şi arte";
            ws.Cell(5, 37).Style.Font.Bold = true;
            ws.Cell(5, 37).Style.Font.FontName = "Times New Roman";
            ws.Cell(5, 37).Style.Font.FontSize = 8;
            ws.Cell(5, 37).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(5, 37).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(5, 37).Style.Alignment.WrapText = true;
            ws.Range(5, 37, 5, 49).Merge();

            // AX5:AX7 = Total (col 50) - merge pe 3 rânduri
            ws.Cell(5, 50).Value = "Total";
            ws.Cell(5, 50).Style.Font.Bold = true;
            ws.Cell(5, 50).Style.Font.FontName = "Times New Roman";
            ws.Cell(5, 50).Style.Font.FontSize = 8;
            ws.Cell(5, 50).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(5, 50).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range(5, 50, 7, 50).Merge();

            ws.Row(5).Height = 24.0;

            // ── RÂND 6: Subdomenii individuale (fiecare merge cu rândul 7) ───────
            for (int i = 0; i < _domeniiANS.Length; i++)
            {
                int col = 10 + i;
                ws.Cell(6, col).Value = _domeniiANS[i];
                ws.Cell(6, col).Style.Font.FontName = "Times New Roman";
                ws.Cell(6, col).Style.Font.FontSize = 7;
                ws.Cell(6, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(6, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Cell(6, col).Style.Alignment.WrapText = true;
                ws.Range(6, col, 7, col).Merge(); // fiecare subdomeniu merge cu randul 7
            }

            ws.Row(6).Height = 19.5;

            // ── RÂND 7: parte din merge-urile de la 6 (implicit prin merge-uri) ──
            ws.Row(7).Height = 60.6;

            // ── RÂND 8: litere A, B, C, D, E, F, 40 ────────────────────────────
            string[] litere = { "A", "B", "C", "D", "E", "F" };
            for (int c = 1; c <= 6; c++)
            {
                ws.Cell(8, c).Value = litere[c - 1];
                ws.Cell(8, c).Style.Font.FontName = "Times New Roman";
                ws.Cell(8, c).Style.Font.FontSize = 7;
                ws.Cell(8, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(8, c).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
            ws.Cell(8, 50).Value = 40;
            ws.Cell(8, 50).Style.Font.FontName = "Times New Roman";
            ws.Cell(8, 50).Style.Font.FontSize = 7;
            ws.Cell(8, 50).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(8).Height = 13.5;

            // ── Borduri pe header (rânduri 5-8, col 1-50) ───────────────────────
            var borderStyle = XLBorderStyleValues.Thin;
            var headerBorderRange = ws.Range(5, 1, 8, 50);
            headerBorderRange.Style.Border.TopBorder = borderStyle;
            headerBorderRange.Style.Border.BottomBorder = borderStyle;
            headerBorderRange.Style.Border.LeftBorder = borderStyle;
            headerBorderRange.Style.Border.RightBorder = borderStyle;
            // Medium pe exterior rânduri 5 (top) si 8 (bottom)
            ws.Range(5, 1, 5, 50).Style.Border.TopBorder = XLBorderStyleValues.Medium;
            ws.Range(8, 1, 8, 50).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

            // ── DATE: rânduri de la 9 ────────────────────────────────────────────
            int rand = 9;
            int nrCrt = 1;
            // Fill galben ANS (theme 7, tint 0.8 ≈ RGB FFFF99 aproape exact)
            var fillGalben = XLColor.FromArgb(255, 255, 255, 153);

            foreach (var prof in profesori)
            {
                ws.Cell(rand, 1).Value = nrCrt++;
                ws.Cell(rand, 1).Style.Font.FontName = "Times New Roman";
                ws.Cell(rand, 1).Style.Font.FontSize = 8;
                ws.Cell(rand, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Cell(rand, 2).Value = $"{prof.Nume} {prof.Prenume}".Trim();
                ws.Cell(rand, 2).Style.Font.FontSize = 10;
                // col 3 (CNP) - gol

                ws.Cell(rand, 4).Value = prof.GradFunctie;
                ws.Cell(rand, 4).Style.Font.FontSize = 10;
                // col 5, 6, 7 - goale (forma angajare, conducator doctorat, varsta)

                ws.Cell(rand, 8).Value = prof.Facultate;
                ws.Cell(rand, 8).Style.Font.FontSize = 10;

                ws.Cell(rand, 9).Value = prof.Departament;
                ws.Cell(rand, 9).Style.Font.FontSize = 10;

                decimal sumaFractiuni = 0;
                for (int i = 0; i < 40; i++)
                {
                    if (prof.Fractiuni[i] != 0)
                    {
                        int col = 10 + i;
                        ws.Cell(rand, col).Value = prof.Fractiuni[i];
                        ws.Cell(rand, col).Style.NumberFormat.Format = "0.00";
                        ws.Cell(rand, col).Style.Font.FontSize = 10;
                        sumaFractiuni += prof.Fractiuni[i];
                    }
                }

                // Col 50: formula SUM(J_:AW_) identic cu template-ul
                ws.Cell(rand, 50).FormulaA1 = $"=SUM(J{rand}:AW{rand})";
                ws.Cell(rand, 50).Style.NumberFormat.Format = "0.00";
                ws.Cell(rand, 50).Style.Font.FontSize = 10;

                // Fill galben pe intregul rand de date (ca in template)
                ws.Range(rand, 1, rand, 50).Style.Fill.BackgroundColor = fillGalben;

                // Borduri thin pe rand
                ws.Range(rand, 1, rand, 50).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ws.Range(rand, 1, rand, 50).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ws.Range(rand, 1, rand, 50).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ws.Range(rand, 1, rand, 50).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                // Medium pe left si bottom-ul fiecarui rand (stilul template)
                ws.Cell(rand, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;

                ws.Row(rand).Height = 13.5;
                rand++;
            }

            // ── RÂND TOTAL ───────────────────────────────────────────────────────
            int randTotal = rand;
            ws.Cell(randTotal, 1).Value = "Total general:";
            ws.Cell(randTotal, 1).Style.Font.Bold = true;
            ws.Cell(randTotal, 1).Style.Font.FontName = "Times New Roman";
            ws.Cell(randTotal, 1).Style.Font.FontSize = 10;
            ws.Cell(randTotal, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(randTotal, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(randTotal, 1).Style.Alignment.WrapText = true;

            ws.Cell(randTotal, 5).Value = profesori.Count; // numar titulari

            // SUM pe fiecare coloana de domeniu (J..AX) cu referinta la randurile de date
            if (profesori.Count > 0)
            {
                for (int col = 10; col <= 50; col++)
                {
                    string cl = ColLetter(col);
                    ws.Cell(randTotal, col).FormulaA1 = $"=SUM({cl}9:{cl}{randTotal - 1})";
                    ws.Cell(randTotal, col).Style.NumberFormat.Format = "0.00";
                    ws.Cell(randTotal, col).Style.Font.Bold = true;
                    ws.Cell(randTotal, col).Style.Font.FontSize = 10;
                }
            }

            ws.Range(randTotal, 1, randTotal, 50).Style.Border.TopBorder = XLBorderStyleValues.Medium;
            ws.Range(randTotal, 1, randTotal, 50).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            ws.Range(randTotal, 1, randTotal, 50).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            ws.Range(randTotal, 1, randTotal, 50).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            ws.Row(randTotal).Height = 42.6;

            // ── Lățimi coloane exacte din template ───────────────────────────────
            ws.Column(1).Width = 7.5546875;
            ws.Column(2).Width = 30.88671875;
            ws.Column(3).Width = 19.44140625;
            ws.Column(4).Width = 18.5546875;
            ws.Column(5).Width = 5.5546875;
            ws.Column(6).Width = 4.6640625;
            ws.Column(7).Width = 6.5546875;
            ws.Column(8).Width = 25.5546875;
            ws.Column(9).Width = 60.33203125;
            ws.Column(10).Width = 5.5546875;  // J - Matematica (ingust ca in template)
            for (int col = 11; col <= 49; col++) ws.Column(col).Width = 13.0;
            ws.Column(50).Width = 6.109375;

            // ── Freeze dupa rândul 8 si dupa coloana 9 ──────────────────────────
            ws.SheetView.FreezeRows(8);
            ws.SheetView.FreezeColumns(9);

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Date_ANS_{idAnUniv}.xlsx");
        }

        private decimal ObtineNormaBaza(string grad)
        {
            if (string.IsNullOrWhiteSpace(grad)) return 0;
            string g = grad.ToLower();
            if (g.Contains("profesor")) return 11;
            if (g.Contains("conferentiar") || g.Contains("conferențiar")) return 12;
            if (g.Contains("lector") || g.Contains("sef lucrari") || g.Contains("șef lucrări") || g.Contains("șef de lucrări")) return 14;
            if (g.Contains("asistent")) return 16;
            return 0;
        }

        private class RandSqlANS
        {
            public string Nume { get; set; } = "";
            public string Prenume { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string Departament { get; set; } = "";
            public string GradFunctie { get; set; } = "";
            public decimal OreConventionale { get; set; }
            public string DomeniuANS { get; set; } = "";
        }
    }
}