using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using ClosedXML.Excel;
using System.Linq;

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

        private readonly string[] DomeniiExcel = new string[]
        {
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

        // Mapping id_metaspecializare (ID intern program studiu) -> ID subdomeniu ANS
        private static readonly Dictionary<int, int> MappingMetaspec = new Dictionary<int, int>
        {
            { 20, 5 }, { 34, 5 }, { 44, 5 }, { 182, 5 }, { 837, 5 }, { 847, 5 }, { 848, 5 },
            { 43, 9 }, { 148, 8 }, { 171, 11 }, { 306, 11 }, { 358, 11 }, { 464, 11 },
            { 466, 11 }, { 615, 11 }, { 823, 11 }, { 828, 11 },
            { 35, 11 }, { 82, 6 }, { 84, 6 }, { 139, 11 }, { 162, 6 }, { 310, 6 },
            { 315, 6 }, { 316, 6 }, { 362, 6 }, { 372, 6 }, { 418, 6 }, { 597, 6 }, { 617, 6 },
            { 116, 12 }, { 126, 11 }, { 129, 11 }, { 156, 11 }, { 228, 11 }, { 229, 12 },
            { 404, 11 }, { 529, 11 }, { 142, 11 }, { 326, 11 }, { 531, 11 }, { 806, 11 }, { 819, 11 },
            { 53, 11 }, { 138, 11 }, { 446, 11 }, { 448, 11 }, { 449, 11 }, { 450, 11 },
            { 451, 11 }, { 496, 11 }, { 497, 11 }, { 821, 11 },
            { 46, 9 }, { 122, 9 }, { 176, 9 }, { 178, 9 }, { 731, 9 },
            { 72, 9 }, { 90, 9 }, { 118, 9 }, { 226, 9 }, { 235, 9 }, { 249, 9 },
            { 307, 9 }, { 437, 9 }, { 458, 9 }, { 566, 9 }, { 845, 9 },
            { 101, 40 }, { 102, 40 }, { 104, 40 }, { 251, 1 }, { 317, 40 },
            { 340, 1 }, { 368, 40 }, { 369, 40 }, { 477, 40 },
            { 45, 25 }, { 73, 25 }, { 93, 25 }, { 112, 24 }, { 221, 25 }, { 223, 25 },
            { 227, 25 }, { 242, 25 }, { 283, 25 }, { 288, 25 }, { 299, 25 },
            { 181, 31 }, { 196, 27 }, { 197, 27 }, { 200, 27 }, { 205, 27 }, { 207, 27 },
            { 209, 27 }, { 217, 27 }, { 218, 27 }, { 341, 27 }, { 343, 27 }, { 463, 27 },
            { 511, 27 }, { 512, 27 }, { 513, 27 }, { 514, 27 }, { 726, 27 }, { 798, 27 }, { 801, 27 },
            { 60, 18 }, { 64, 18 }, { 331, 18 }, { 485, 18 }, { 555, 18 },
            { 41, 20 }, { 98, 20 }, { 100, 25 }, { 322, 21 }, { 524, 25 }, { 579, 20 }, { 831, 20 },
            { 276, 26 }, { 294, 26 }, { 296, 26 }, { 383, 26 }, { 384, 26 }, { 416, 26 },
            { 515, 26 }, { 813, 26 }, { 851, 26 }, { 832, 26 }, { 834, 26 },
            { 394, 14 }, { 397, 14 }, { 402, 14 }, { 484, 14 }, { 585, 14 }, { 594, 14 }, { 835, 14 },
            { 186, 37 }, { 187, 37 }, { 264, 37 }, { 332, 37 }, { 351, 37 }, { 557, 37 }, { 838, 37 },
            { 78, 39 }, { 189, 39 }, { 325, 39 }, { 470, 39 }, { 783, 39 }, { 784, 39 }, { 846, 39 },
        };

        // ID ANS -> coloana Excel (1-based), ordinea exacta din template Diana Ionita
        private static readonly Dictionary<int, int> AnsIdToCol = new Dictionary<int, int>
        {
            { 1, 10 }, { 40, 11 }, { 2, 12 }, { 3, 13 }, { 4, 14 },
            { 5, 15 }, { 6, 16 }, { 7, 17 }, { 8, 18 }, { 9, 19 }, { 10, 20 }, { 11, 21 },
            { 12, 22 }, { 13, 23 }, { 14, 24 }, { 15, 25 }, { 16, 26 }, { 17, 27 },
            { 18, 28 }, { 19, 29 }, { 20, 30 }, { 21, 31 }, { 22, 32 }, { 23, 33 },
            { 24, 34 }, { 25, 35 }, { 26, 36 },
            { 27, 37 }, { 28, 38 }, { 29, 39 }, { 30, 40 }, { 31, 41 }, { 32, 42 },
            { 33, 43 }, { 34, 44 }, { 35, 45 }, { 36, 46 }, { 37, 47 }, { 38, 48 }, { 39, 49 },
        };

        public RapoarteController(IConfiguration configuration, IMemoryCache cache)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        #region ================= LISTE (DROPDOWNS) =================

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
                            UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) COLLATE DATABASE_DEFAULT as FacCurata
                        FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                        WHERE ppm.DenumireFacultate IS NOT NULL
                        ORDER BY FacCurata ASC";
                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                        lista.Add(reader["FacCurata"]?.ToString() ?? "");
                }
                return lista;
            }));
        }

        [HttpGet("liste/departamente")]
        public ActionResult GetDepartamente(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT DISTINCT 
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) COLLATE DATABASE_DEFAULT as DeptCurat
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                    WHERE ISNULL(ppm.DenumireCatedra, '') <> ''
                      AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) = @fac)
                    ORDER BY DeptCurat ASC";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    lista.Add(reader["DeptCurat"]?.ToString() ?? "");
            }
            return Ok(lista);
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
                    WHERE UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) = @fac
                    GROUP BY ID_FacultateSpecializare
                    ORDER BY COUNT(*) DESC;

                    SELECT DISTINCT 
                        UPPER(LTRIM(RTRIM(
                            REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                                CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                     THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                     ELSE sf.DenumireSpecializare END, 
                            ' - CORECT', ''), ' CORECT', ''), ' - COPIE', ''), 'S', 'S'), 'T', 'T')
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
                            ' - CORECT', ''), ' CORECT', ''), ' - COPIE', ''), 'S', 'S'), 'T', 'T')
                        ))) as SpecCurata
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    WHERE sf.DenumireSpecializare IS NOT NULL
                    ORDER BY SpecCurata";
                }
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                if (!string.IsNullOrEmpty(numeFacultate) && numeFacultate != "Toti")
                    cmd.Parameters.AddWithValue("@fac", numeFacultate);
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
        public ActionResult GetProfesori(string? anUniv, string? facultate, string? specializari, string? departament)
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
                    AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) = @fac)
                    AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) = @dept)
                    AND (@allSpecs = 1 OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) IN (SELECT value FROM STRING_SPLIT(@listaSpecs, ',')))
                ORDER BY ppm.NumeIntreg";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@allSpecs", toateSpecializarile ? 1 : 0);
                cmd.Parameters.AddWithValue("@listaSpecs", specializari ?? "");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    lista.Add(reader["NumeIntreg"]?.ToString() ?? "");
            }
            return Ok(lista);
        }

        #endregion

        #region ================= RAPORT 1: NORMA PROFESORI =================

        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new List<object>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS SpecializareCurata,
                        ISNULL(sf.DenumireMaterie, 'Nedefinit') AS DenumireMaterie,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
                        ISNULL(sf.Nr_Ore_Curs, 0) AS OreCursLinie,
                        ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0) AS OreAplicatiiLinie,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                        sf.DenumireSpecializare AS NumeSpecOriginal
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                    LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                ),
                Filtrat AS (
                    SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru,
                           SUM(OreConvLinie) AS TotalOreConvItem,
                           SUM(OreCursLinie) AS TotalOreCurs,
                           SUM(OreAplicatiiLinie) AS TotalOreAplicatii
                    FROM BaseData
                    WHERE 
                        (@an = 'Toti' OR AnCurat = @an) AND (@fac = 'Toti' OR FacultateCurata = @fac) AND
                        (@prof = 'Toti' OR NumeIntreg = @prof) AND
                        (@dept = 'Toti' OR DepartamentCurat = @dept) AND
                        (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%') AND
                        (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ','))) AND
                        (@semestru = 0 OR Semestru = @semestru) AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                    GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru
                )
                SELECT f.NumeIntreg AS Profesor, f.SpecializareCurata AS Specializare, f.DenumireMaterie AS Materie,
                       f.TipPost, f.Semestru,
                       f.TotalOreCurs AS NrOreCurs,
                       f.TotalOreAplicatii AS NrOreAplicatii,
                       f.TotalOreConvItem AS NrOreConventionale,
                       SUM(f.TotalOreConvItem) OVER(PARTITION BY f.NumeIntreg, f.TipPost) AS TotalTipPost,
                       SUM(f.TotalOreConvItem) OVER(PARTITION BY f.NumeIntreg) AS TotalPost
                FROM Filtrat f 
                ORDER BY f.NumeIntreg, f.TipPost, f.DenumireMaterie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
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
                        NrOreCurs = reader["NrOreCurs"],
                        NrOreAplicatii = reader["NrOreAplicatii"],
                        NrOreConventionale = reader["NrOreConventionale"],
                        TotalTipPost = reader["TotalTipPost"],
                        TotalPost = reader["TotalPost"]
                    });
                }
            }
            return Ok(result);
        }

        [HttpGet("export/norma")]
        public IActionResult ExportNormaExcel(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new DataTable("NormaProfesori");
            result.Columns.AddRange(new[] {
                new DataColumn("Profesor"),
                new DataColumn("Specializare"),
                new DataColumn("Materie"),
                new DataColumn("Tip Post"),
                new DataColumn("Semestru"),
                new DataColumn("Nr Ore Curs", typeof(double)),
                new DataColumn("Nr Ore Aplicatii", typeof(double)),
                new DataColumn("Nr Ore Conventionale", typeof(double))
            });

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS SpecializareCurata,
                        ISNULL(sf.DenumireMaterie, 'Nedefinit') AS DenumireMaterie,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
                        ISNULL(sf.Nr_Ore_Curs, 0) AS OreCursLinie,
                        ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) + ISNULL(sf.Nr_Ore_Proiect, 0) AS OreAplicatiiLinie,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                        sf.DenumireSpecializare AS NumeSpecOriginal
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                    LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                ),
                Filtrat AS (
                    SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru,
                           SUM(OreConvLinie) AS TotalOreConvItem,
                           SUM(OreCursLinie) AS TotalOreCurs,
                           SUM(OreAplicatiiLinie) AS TotalOreAplicatii
                    FROM BaseData
                    WHERE 
                        (@an = 'Toti' OR AnCurat = @an) AND (@fac = 'Toti' OR FacultateCurata = @fac) AND
                        (@prof = 'Toti' OR NumeIntreg = @prof) AND
                        (@dept = 'Toti' OR DepartamentCurat = @dept) AND
                        (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%') AND
                        (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ','))) AND
                        (@semestru = 0 OR Semestru = @semestru) AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                    GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru
                )
                SELECT f.NumeIntreg, f.SpecializareCurata, f.DenumireMaterie, f.TipPost, f.Semestru,
                       f.TotalOreCurs AS NrOreCurs,
                       f.TotalOreAplicatii AS NrOreAplicatii,
                       f.TotalOreConvItem AS NrOreConventionale
                FROM Filtrat f 
                ORDER BY f.NumeIntreg, f.TipPost, f.DenumireMaterie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
                cmd.Parameters.AddWithValue("@semestru", semestru);
                cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    result.Rows.Add(reader["NumeIntreg"], reader["SpecializareCurata"], reader["DenumireMaterie"], reader["TipPost"], reader["Semestru"], reader["NrOreCurs"], reader["NrOreAplicatii"], reader["NrOreConventionale"]);
            }

            string fileName = "NormaProfesori_General.xlsx";
            if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                fileName = $"NormaProfesori_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.xlsx";

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Norme");
                ws.Cell(1, 1).Value = "Filtre Aplicate";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
                ws.Cell(2, 1).Value = $"An Universitar: {anUniv} | Facultate: {facultate} | Departament: {departament}";
                ws.Cell(3, 1).Value = $"Profesor: {profesor} | Semestru: {(semestru == 0 ? "Toate" : semestru.ToString())} | Tip Post: {tipPost} | Forma Inv: {formaInvatamant}";
                var table = ws.Cell(5, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;
                table.ShowTotalsRow = true;
                table.Field("Nr Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Nr Ore Aplicatii").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Nr Ore Conventionale").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
                ws.Columns().AdjustToContents();
                var headerRange = ws.Range(5, 1, 5, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                var dataRange = ws.Range(5, 1, 5 + result.Rows.Count + 1, result.Columns.Count);
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }

        #endregion

        #region ================= RAPORT 2: ORE PE PROGRAM =================

        [HttpGet("ore-profesor-program")]
        public async Task<IActionResult> GetOreProfProgram(string? anUniv = "Toti", string? facultate = "Toti", string? specializari = "Toti", string? profesor = "Toti", int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var listaResult = new List<object>();
            string sql = @"
            WITH BaseData AS (
                SELECT 
                    ppm.NumeIntreg AS Profesor,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS ProgramStudiu,
                    ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
                    ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost,
                    ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                    UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                    UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                    sf.DenumireSpecializare AS NumeSpecOriginal
                FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
            ),
            Filtrat AS (
                SELECT Profesor, ProgramStudiu, SUM(OreConvLinie) AS OreConvProgram
                FROM BaseData
                WHERE 
                    (@AnUniv = 'Toti' OR AnCurat = @AnUniv) AND (@Facultate = 'Toti' OR FacultateCurata = @Facultate) AND
                    (@Profesor = 'Toti' OR Profesor = @Profesor) AND (@dept = 'Toti' OR DepartamentCurat = @dept) AND
                    (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%') AND
                    (@Specializari = 'Toti' OR ProgramStudiu IN (SELECT value FROM STRING_SPLIT(@Specializari, ','))) AND
                    (@Semestru = 0 OR Semestru = @Semestru) AND (@TipPost = 'Toti' OR TipPost = @TipPost)
                GROUP BY Profesor, ProgramStudiu
            ),
            TotalProfesor AS (SELECT Profesor, SUM(OreConvProgram) AS TotalPost FROM Filtrat GROUP BY Profesor)
            SELECT f.Profesor, ISNULL(f.ProgramStudiu, 'Nespecificat') AS ProgramStudiu, f.OreConvProgram AS NrOreConv, t.TotalPost,
                CAST(CASE WHEN t.TotalPost = 0 THEN 0 ELSE (f.OreConvProgram / t.TotalPost) * 100 END AS DECIMAL(10,2)) AS ProcentPost
            FROM Filtrat f INNER JOIN TotalProfesor t ON f.Profesor = t.Profesor
            ORDER BY f.Profesor, f.OreConvProgram DESC";

            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                using (var command = new SqlCommand(sql, connection))
                {
                    command.Parameters.AddWithValue("@AnUniv", anUniv ?? "Toti");
                    command.Parameters.AddWithValue("@Facultate", facultate ?? "Toti");
                    command.Parameters.AddWithValue("@dept", departament ?? "Toti");
                    command.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
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

        [HttpGet("export/ore-program")]
        public async Task<IActionResult> ExportOreProgramExcel(string? anUniv, string? facultate, string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new DataTable("OreProgram");
            result.Columns.AddRange(new[] {
                new DataColumn("Profesor"), new DataColumn("Program Studiu"),
                new DataColumn("Nr Ore Conv", typeof(double)),
                new DataColumn("Procent Post", typeof(double)),
                new DataColumn("Total Post", typeof(double))
            });

            string sql = @"
            WITH BaseData AS (
                SELECT ppm.NumeIntreg,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS ProgramStudiu,
                    ISNULL(sf.NrOreConventionale, 0) AS OreConvLinie,
                    ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost, ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                    UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                    UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                    sf.DenumireSpecializare AS NumeSpecOriginal
                FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf ON ppm.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
            ),
            Filtrat AS (
                SELECT NumeIntreg, ProgramStudiu, SUM(OreConvLinie) AS OreConvProgram FROM BaseData
                WHERE (@AnUniv = 'Toti' OR AnCurat = @AnUniv) AND (@Facultate = 'Toti' OR FacultateCurata = @Facultate) AND
                      (@dept = 'Toti' OR DepartamentCurat = @dept) AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%') AND
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
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti"); cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
                cmd.Parameters.AddWithValue("@Specializari", specializari ?? "Toti"); cmd.Parameters.AddWithValue("@Profesor", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@Semestru", semestru); cmd.Parameters.AddWithValue("@TipPost", tipPost ?? "Toti");
                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                    result.Rows.Add(reader["NumeIntreg"], reader["ProgramStudiu"], Convert.ToDouble(reader["OreConvProgram"]), Convert.ToDouble(reader["ProcentPost"]), Convert.ToDouble(reader["TotalPost"]));
            }

            string fileName = "StatisticaOre_General.xlsx";
            if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                fileName = $"StatisticaOre_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.xlsx";

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Distributie Ore");
                ws.Cell(1, 1).Value = "Filtre Aplicate";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
                ws.Cell(2, 1).Value = $"An Universitar: {anUniv} | Facultate: {facultate} | Departament: {departament}";
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

        #endregion

        #region ================= RAPORT 3: NORME PROFESORI TOTALURI =================

        [HttpGet("norma-totaluri")]
        public ActionResult GetNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var result = new List<object>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                // FIX: un singur rand per profesor - ROW_NUMBER preia departamentul cu cele mai multe ore
                string sql = @"
                WITH DateProfesor AS (
                    SELECT 
                        UPPER(LTRIM(RTRIM(ppm.NumeIntreg))) AS NumeComplet,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, 'Nespecificat')))) AS Departament,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, 'Nespecificat')))) AS Facultate,
                        ISNULL(sf.NrOreConventionale, 0) AS OreConv,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                    WHERE (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                      AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, '')))) = @fac)
                      AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) = @dept)
                      AND (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ppm.NumeIntreg))) = @prof)
                ),
                TotaluriPerCombinatie AS (
                    SELECT NumeComplet, Departament, Facultate,
                           SUM(OreConv) AS TotalOreConv,
                           SUM(OreConv) * 14 AS TotalAnualOreConv,
                           ROW_NUMBER() OVER(PARTITION BY NumeComplet ORDER BY SUM(OreConv) DESC) AS Rang
                    FROM DateProfesor
                    GROUP BY NumeComplet, Departament, Facultate
                )
                SELECT NumeComplet, Departament, Facultate, TotalOreConv, TotalAnualOreConv
                FROM TotaluriPerCombinatie
                WHERE Rang = 1
                ORDER BY NumeComplet";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    decimal totalOre = reader["TotalOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalOreConv"]) : 0;
                    decimal totalAnual = reader["TotalAnualOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAnualOreConv"]) : 0;
                    result.Add(new
                    {
                        Profesor = reader["NumeComplet"],
                        Departament = reader["Departament"],
                        Facultate = reader["Facultate"],
                        TotalOreConv = Math.Round(totalOre, 2),
                        TotalAnualOreConv = Math.Round(totalAnual, 2)
                    });
                }
            }
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public IActionResult ExportNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var result = new DataTable("TotaluriNorme");
            result.Columns.AddRange(new[] {
                new DataColumn("Nume Profesor"), new DataColumn("Departament"), new DataColumn("Facultate"),
                new DataColumn("Total Ore Conv.", typeof(decimal)), new DataColumn("Total tot anul ore conv.", typeof(decimal))
            });

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                // FIX: un singur rand per profesor
                string sql = @"
                WITH DateProfesor AS (
                    SELECT 
                        UPPER(LTRIM(RTRIM(ppm.NumeIntreg))) AS NumeComplet,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, 'Nespecificat')))) AS Departament,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, 'Nespecificat')))) AS Facultate,
                        ISNULL(sf.NrOreConventionale, 0) AS OreConv
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                    WHERE (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                      AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, '')))) = @fac)
                      AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) = @dept)
                      AND (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ppm.NumeIntreg))) = @prof)
                ),
                TotaluriPerCombinatie AS (
                    SELECT NumeComplet, Departament, Facultate,
                           SUM(OreConv) AS TotalOreConv,
                           SUM(OreConv) * 14 AS TotalAnualOreConv,
                           ROW_NUMBER() OVER(PARTITION BY NumeComplet ORDER BY SUM(OreConv) DESC) AS Rang
                    FROM DateProfesor
                    GROUP BY NumeComplet, Departament, Facultate
                )
                SELECT NumeComplet, Departament, Facultate, TotalOreConv, TotalAnualOreConv
                FROM TotaluriPerCombinatie
                WHERE Rang = 1
                ORDER BY NumeComplet";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    decimal totalOre = reader["TotalOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalOreConv"]) : 0;
                    decimal totalAnual = reader["TotalAnualOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAnualOreConv"]) : 0;
                    result.Rows.Add(reader["NumeComplet"], reader["Departament"], reader["Facultate"], Math.Round(totalOre, 2), Math.Round(totalAnual, 2));
                }
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Totaluri Norme");
                var table = ws.Cell(1, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;
                table.ShowTotalsRow = true;
                table.Field("Total Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Total tot anul ore conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Nume Profesor").TotalsRowLabel = "TOTAL GENERAL";
                ws.Columns().AdjustToContents();
                var headerRange = ws.Range(1, 1, 1, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Totaluri_Norme.xlsx");
                }
            }
        }

        #endregion

        #region ================= RAPORT 4: LIMBI STRAINE =================

        [HttpGet("limbi-straine")]
        public ActionResult GetLimbiStraine(string? anUniv, string? facultate, string? departament, string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg AS NumeComplet,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        ISNULL(sf.NrOreConventionale, 0) AS OreConv,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                        sf.DenumireSpecializare AS NumeSpecOriginal,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS SpecializareCurata,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                )
                SELECT 
                    NumeComplet,
                    SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConv ELSE 0 END) * 14 AS Sem1,
                    SUM(CASE WHEN Semestru IN (2,4,6,8,10,12) THEN OreConv ELSE 0 END) * 14 AS Sem2,
                    SUM(OreConv) * 14 AS Total
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@dept = 'Toti' OR DepartamentCurat = @dept)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@prof = 'Toti' OR NumeComplet = @prof)
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                  AND (
                       NumeSpecOriginal LIKE '%englez%' OR NumeSpecOriginal LIKE '%francez%' 
                       OR NumeSpecOriginal LIKE '%german%' OR NumeSpecOriginal LIKE '%american%'
                       OR NumeSpecOriginal LIKE '%(EN)%' OR NumeSpecOriginal LIKE '%(FR)%' OR NumeSpecOriginal LIKE '%(G)%'
                       OR NumeSpecOriginal IN (
                           'Inginerie virtuala in proiectarea autovehiculelor',
                           'Metode practice integrate in ingineria sistemelor de propulsie', 
                           'Ingineria proceselor de fabricatie avansate',
                           'Managementul afacerilor industriale si antreprenoriat', 
                           'Inginerie electrica si calculatoare', 'Sisteme electrice avansate',
                           'Securitate cibernetica', 'Informatica aplicata', 'Tehnologii Internet',
                           'Cultura si discurs in spatiul anglo american', 
                           'Studii de limba si de cultura franceza',
                           'Studii de limba si literatura germana din perspectiva interculturala', 
                           'Studii lingvistice pentru comunicare interculturala',
                           'Traducere si interpretariat din limba franceza in limba romana', 
                           'Studii americane', 'Performanta umana in antrenamentul sportiv',
                           'Administrarea afacerilor', 'Managementul resurselor umane',
                           'Dezvoltarea afacerilor turistice', 'Medicina traditionala chineza'
                       )
                  )
                GROUP BY NumeComplet
                HAVING SUM(OreConv) > 0
                ORDER BY NumeComplet";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
                cmd.Parameters.AddWithValue("@semestru", semestru);
                cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");
                using var reader = cmd.ExecuteReader();
                int nrCrt = 1;
                while (reader.Read())
                {
                    result.Add(new
                    {
                        NrCrt = nrCrt++,
                        NumeProfesor = reader["NumeComplet"],
                        Sem1 = reader["Sem1"],
                        Sem2 = reader["Sem2"],
                        Total = reader["Total"]
                    });
                }
            }
            return Ok(result);
        }

        [HttpGet("export/limbi-straine")]
        public IActionResult ExportLimbiStraine(string? anUniv, string? facultate, string? departament, string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new DataTable("LimbiStraine");
            result.Columns.AddRange(new[] {
                new DataColumn("Nr. Crt.", typeof(int)), new DataColumn("Nume si prenume profesor"),
                new DataColumn("Total Sem 1", typeof(decimal)), new DataColumn("Total Sem 2", typeof(decimal)), new DataColumn("Total", typeof(decimal))
            });

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg AS NumeComplet,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        ISNULL(sf.NrOreConventionale, 0) AS OreConv,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                        sf.DenumireSpecializare AS NumeSpecOriginal,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS SpecializareCurata,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                )
                SELECT 
                    NumeComplet,
                    SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConv ELSE 0 END) * 14 AS Sem1,
                    SUM(CASE WHEN Semestru IN (2,4,6,8,10,12) THEN OreConv ELSE 0 END) * 14 AS Sem2,
                    SUM(OreConv) * 14 AS Total
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@dept = 'Toti' OR DepartamentCurat = @dept)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@prof = 'Toti' OR NumeComplet = @prof)
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                  AND (
                       NumeSpecOriginal LIKE '%englez%' OR NumeSpecOriginal LIKE '%francez%' 
                       OR NumeSpecOriginal LIKE '%german%' OR NumeSpecOriginal LIKE '%american%'
                       OR NumeSpecOriginal LIKE '%(EN)%' OR NumeSpecOriginal LIKE '%(FR)%' OR NumeSpecOriginal LIKE '%(G)%'
                       OR NumeSpecOriginal IN (
                           'Inginerie virtuala in proiectarea autovehiculelor',
                           'Metode practice integrate in ingineria sistemelor de propulsie', 
                           'Ingineria proceselor de fabricatie avansate',
                           'Managementul afacerilor industriale si antreprenoriat', 
                           'Inginerie electrica si calculatoare', 'Sisteme electrice avansate',
                           'Securitate cibernetica', 'Informatica aplicata', 'Tehnologii Internet',
                           'Cultura si discurs in spatiul anglo american', 
                           'Studii de limba si de cultura franceza',
                           'Studii de limba si literatura germana din perspectiva interculturala', 
                           'Studii lingvistice pentru comunicare interculturala',
                           'Traducere si interpretariat din limba franceza in limba romana', 
                           'Studii americane', 'Performanta umana in antrenamentul sportiv',
                           'Administrarea afacerilor', 'Managementul resurselor umane',
                           'Dezvoltarea afacerilor turistice', 'Medicina traditionala chineza'
                       )
                  )
                GROUP BY NumeComplet
                HAVING SUM(OreConv) > 0
                ORDER BY NumeComplet";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
                cmd.Parameters.AddWithValue("@semestru", semestru);
                cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");
                using var reader = cmd.ExecuteReader();
                int nrCrt = 1;
                while (reader.Read())
                    result.Rows.Add(nrCrt++, reader["NumeComplet"], reader["Sem1"], reader["Sem2"], reader["Total"]);
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Limbi Straine");
                var table = ws.Cell(1, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;
                table.ShowTotalsRow = true;
                table.Field("Total Sem 1").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Total Sem 2").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Total").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Nume si prenume profesor").TotalsRowLabel = "TOTAL GENERAL";
                ws.Columns().AdjustToContents();
                var headerRange = ws.Range(1, 1, 1, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Raport_Limbi_Straine.xlsx");
                }
            }
        }

        #endregion

        #region ================= RAPORT 5: DISCIPLINE PREDATE =================

        [HttpGet("discipline-predate")]
        public ActionResult GetDisciplinePredate(string? anUniv, string? facultate, string? departament, string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg AS NumeComplet,
                        ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                        ISNULL(sf.DenumireMaterie, 'Nespecificat') AS Materie,
                        STUFF(
                            CASE WHEN ISNULL(sf.Nr_Ore_Curs, 0) > 0 THEN ', Curs' ELSE '' END +
                            CASE WHEN ISNULL(sf.Nr_Ore_Seminar, 0) > 0 THEN ', Seminar' ELSE '' END +
                            CASE WHEN ISNULL(sf.Nr_Ore_Laborator, 0) > 0 THEN ', Laborator' ELSE '' END +
                            CASE WHEN ISNULL(sf.Nr_Ore_Proiect, 0) > 0 THEN ', Proiect' ELSE '' END
                        , 1, 2, '') AS TipActivitate,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                        sf.DenumireSpecializare AS NumeSpecOriginal,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS SpecializareCurata,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                )
                SELECT DISTINCT NumeComplet, Departament, Materie, TipActivitate
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@dept = 'Toti' OR DepartamentCurat = @dept)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@prof = 'Toti' OR NumeComplet = @prof)
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                ORDER BY NumeComplet, Materie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
                cmd.Parameters.AddWithValue("@semestru", semestru);
                cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        Profesor = reader["NumeComplet"],
                        Departament = reader["Departament"],
                        Materie = reader["Materie"],
                        TipActivitate = reader["TipActivitate"]?.ToString() ?? ""
                    });
                }
            }
            return Ok(result);
        }

        [HttpGet("export/discipline-predate")]
        public IActionResult ExportDisciplinePredate(string? anUniv, string? facultate, string? departament, string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new DataTable("Discipline");
            result.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"), new DataColumn("Departament"),
                new DataColumn("Disciplina"), new DataColumn("Tip Activitate")
            });

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg AS NumeComplet,
                        ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                        ISNULL(sf.DenumireMaterie, 'Nespecificat') AS Materie,
                        STUFF(
                            CASE WHEN ISNULL(sf.Nr_Ore_Curs, 0) > 0 THEN ', Curs' ELSE '' END +
                            CASE WHEN ISNULL(sf.Nr_Ore_Seminar, 0) > 0 THEN ', Seminar' ELSE '' END +
                            CASE WHEN ISNULL(sf.Nr_Ore_Laborator, 0) > 0 THEN ', Laborator' ELSE '' END +
                            CASE WHEN ISNULL(sf.Nr_Ore_Proiect, 0) > 0 THEN ', Proiect' ELSE '' END
                        , 1, 2, '') AS TipActivitate,
                        ISNULL(sf.NrSemestruDinAn, 0) AS Semestru,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) AS FacultateCurata,
                        UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) AS DepartamentCurat,
                        UPPER(LTRIM(RTRIM(au.Denumire))) AS AnCurat,
                        sf.DenumireSpecializare AS NumeSpecOriginal,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1) ELSE sf.DenumireSpecializare END, 'S', 'S'), 'T', 'T')))) AS SpecializareCurata,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') AS TipPost
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                )
                SELECT DISTINCT NumeComplet, Departament, Materie, TipActivitate
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@dept = 'Toti' OR DepartamentCurat = @dept)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@prof = 'Toti' OR NumeComplet = @prof)
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                ORDER BY NumeComplet, Materie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
                cmd.Parameters.AddWithValue("@semestru", semestru);
                cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    result.Rows.Add(reader["NumeComplet"], reader["Departament"], reader["Materie"], reader["TipActivitate"]?.ToString() ?? "");
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Discipline Predate");
                var table = ws.Cell(1, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;
                ws.Columns().AdjustToContents();
                var headerRange = ws.Range(1, 1, 1, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Discipline_Predate.xlsx");
                }
            }
        }

        #endregion

        #region ================= RAPORT 6: CADRE DIDACTICE TITULARE =================

        [HttpGet("titulari")]
        public ActionResult GetTitulari(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                SELECT DISTINCT 
                    ppm.NumeIntreg AS NumeComplet,
                    ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                    ISNULL(ppm.DenumireFacultate, 'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                WHERE sf.TitularSauSuplinitor = 1
                  AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, '')))) = @fac)
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) = @dept)
                ORDER BY ppm.NumeIntreg";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        Profesor = reader["NumeComplet"],
                        Departament = reader["Departament"],
                        Facultate = reader["Facultate"]
                    });
                }
            }
            return Ok(result);
        }

        [HttpGet("export/titulari")]
        public IActionResult ExportTitulari(string? anUniv, string? facultate, string? departament)
        {
            var result = new DataTable("Titulari");
            result.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"), new DataColumn("Departament"), new DataColumn("Facultate")
            });
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                SELECT DISTINCT 
                    ppm.NumeIntreg AS NumeComplet,
                    ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                    ISNULL(ppm.DenumireFacultate, 'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                WHERE sf.TitularSauSuplinitor = 1
                  AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, '')))) = @fac)
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) = @dept)
                ORDER BY ppm.NumeIntreg";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    result.Rows.Add(reader["NumeComplet"], reader["Departament"], reader["Facultate"]);
            }
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Titulari");
                var table = ws.Cell(1, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;
                ws.Columns().AdjustToContents();
                var headerRange = ws.Range(1, 1, 1, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Cadre_Didactice_Titulare.xlsx");
                }
            }
        }

        #endregion

        #region ================= RAPORT 7: ASOCIATI / COLABORATORI =================

        [HttpGet("colaboratori")]
        public ActionResult GetColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                SELECT DISTINCT 
                    ppm.NumeIntreg AS NumeComplet,
                    ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                    ISNULL(ppm.DenumireFacultate, 'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                WHERE sf.TitularSauSuplinitor = 0
                  AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, '')))) = @fac)
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) = @dept)
                  AND ppm.NumeIntreg NOT IN (
                      SELECT DISTINCT p2.NumeIntreg 
                      FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] s2
                      INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] p2 ON s2.ID_Post_Profesor_Materie = p2.ID_Post_Profesor_Materie
                      INNER JOIN [AGSIS].[dbo].[AnUniversitar] au2 ON p2.ID_AnUniv = au2.ID_AnUniv
                      WHERE s2.TitularSauSuplinitor = 1 
                        AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au2.Denumire))) = @an)
                  )
                ORDER BY ppm.NumeIntreg";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        Profesor = reader["NumeComplet"],
                        Departament = reader["Departament"],
                        Facultate = reader["Facultate"]
                    });
                }
            }
            return Ok(result);
        }

        [HttpGet("export/colaboratori")]
        public IActionResult ExportColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var result = new DataTable("Colaboratori");
            result.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"), new DataColumn("Departament"), new DataColumn("Facultate")
            });
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                SELECT DISTINCT 
                    ppm.NumeIntreg AS NumeComplet,
                    ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                    ISNULL(ppm.DenumireFacultate, 'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                WHERE sf.TitularSauSuplinitor = 0
                  AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireFacultate, '')))) = @fac)
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(ppm.DenumireCatedra, '')))) = @dept)
                  AND ppm.NumeIntreg NOT IN (
                      SELECT DISTINCT p2.NumeIntreg 
                      FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] s2
                      INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] p2 ON s2.ID_Post_Profesor_Materie = p2.ID_Post_Profesor_Materie
                      INNER JOIN [AGSIS].[dbo].[AnUniversitar] au2 ON p2.ID_AnUniv = au2.ID_AnUniv
                      WHERE s2.TitularSauSuplinitor = 1 
                        AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au2.Denumire))) = @an)
                  )
                ORDER BY ppm.NumeIntreg";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    result.Rows.Add(reader["NumeComplet"], reader["Departament"], reader["Facultate"]);
            }
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Colaboratori");
                var table = ws.Cell(1, 1).InsertTable(result);
                table.Theme = XLTableTheme.None;
                ws.Columns().AdjustToContents();
                var headerRange = ws.Range(1, 1, 1, result.Columns.Count);
                headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Cadre_Didactice_Colaboratori.xlsx");
                }
            }
        }

        #endregion


        #region ================= RAPORT 8: ANS =================

        [HttpGet("date-ans")]
        public IActionResult GetDateANS([FromQuery] int idAnUniv = 45)
        {
            var dateBrute = new List<RandSqlANS>();
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string query = @"
                    SELECT 
                        ppm.NumeIntreg                                AS NumeComplet,
                        ppm.DenumireGradDidacticPost                  AS GradFunctie,
                        ISNULL(sf.NrOreConventionale, 0)              AS OreConventionale,
                        ISNULL(ppm.DenumireFacultate, 'Nespecificat') AS Facultate,
                        ISNULL(ppm.DenumireCatedra, 'Nespecificat')   AS Departament,
                        sf.id_metaspecializare                        AS IdMetaspec
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm 
                        ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    WHERE sf.id_anuniv = @ID_AnUniv
                      AND sf.DenTitularSauSuplinitor = 'Tit'";
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int idMeta = reader["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(reader["IdMetaspec"]) : 0;
                            int idAns;
                            if (!MappingMetaspec.TryGetValue(idMeta, out idAns)) continue;
                            if (!AnsIdToCol.ContainsKey(idAns)) continue;
                            dateBrute.Add(new RandSqlANS
                            {
                                NumeComplet = reader["NumeComplet"]?.ToString() ?? "",
                                Facultate = reader["Facultate"]?.ToString() ?? "",
                                Departament = reader["Departament"]?.ToString() ?? "",
                                GradFunctie = reader["GradFunctie"]?.ToString() ?? "",
                                OreConventionale = reader["OreConventionale"] != DBNull.Value ? Convert.ToDecimal(reader["OreConventionale"]) : 0m,
                                IdANS = idAns,
                            });
                        }
                    }
                }
            }

            // Grupare dupa NumeComplet - un profesor = un singur rand
            var profesori = dateBrute
                .GroupBy(x => x.NumeComplet)
                .Select(g =>
                {
                    var grupDept = g
                        .GroupBy(x => new { x.Departament, x.Facultate })
                        .Select(dg => new {
                            dg.Key.Departament,
                            dg.Key.Facultate,
                            TotalOre = dg.Sum(x => x.OreConventionale),
                            Grad = dg.OrderByDescending(x => x.OreConventionale).First().GradFunctie
                        })
                        .OrderByDescending(d => d.TotalOre)
                        .First();

                    var orePerAns = g.GroupBy(x => x.IdANS)
                                     .ToDictionary(ag => ag.Key, ag => ag.Sum(x => x.OreConventionale));
                    decimal totalOre = orePerAns.Values.Sum();
                    var fractiuni = new Dictionary<string, decimal>();
                    if (totalOre > 0)
                    {
                        int maxKey = orePerAns.OrderByDescending(x => x.Value).First().Key;
                        decimal sum = 0;
                        foreach (var kv in orePerAns)
                        {
                            if (kv.Key == maxKey) continue;
                            decimal frac = Math.Round(kv.Value / totalOre, 2);
                            fractiuni[DomeniiExcel[AnsIdToCol[kv.Key] - 10]] = frac;
                            sum += frac;
                        }
                        fractiuni[DomeniiExcel[AnsIdToCol[maxKey] - 10]] = Math.Round(1m - sum, 2);
                    }
                    return new
                    {
                        NumeComplet = g.Key,
                        Facultate = grupDept.Facultate,
                        Departament = grupDept.Departament,
                        GradFunctie = MapareFunctieANS(grupDept.Grad),
                        DomeniiMapate = fractiuni
                    };
                })
                .OrderBy(p => p.NumeComplet)
                .ToList();

            return Ok(profesori);
        }

        [HttpGet("export/raport-ans")]
        public IActionResult ExportRaportANS([FromQuery] int idAnUniv = 45)
        {
            var dateBrute = new List<RandSqlANS>();
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                // Filtru: DenTitularSauSuplinitor = 'Tit' EXACT
                // 'Tit'    = titular -> INCLUS
                // 'SupTit' = suplinitor al unui titular -> EXCLUS (cauza dublurilor!)
                // 'Sup'    = suplinitor -> EXCLUS
                // Gradul (Profesor, Lector, Asistent etc.) vine din DenumireGradDidacticPost
                string query = @"
                    SELECT
                        ppm.NumeIntreg                              AS NumeComplet,
                        ppm.DenumireGradDidacticPost                AS GradFunctie,
                        ISNULL(ppm.DenumireCatedra, 'Nespecificat') AS Departament,
                        ISNULL(ppm.DenumireFacultate, '')           AS Facultate,
                        ISNULL(sf.NrOreConventionale, 0)            AS OreConventionale,
                        sf.id_metaspecializare                      AS IdMetaspec
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                        ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    WHERE sf.id_anuniv = @ID_AnUniv
                      AND sf.DenTitularSauSuplinitor = 'Tit'";

                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int idMeta = reader["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(reader["IdMetaspec"]) : 0;
                            int idAns;
                            if (!MappingMetaspec.TryGetValue(idMeta, out idAns)) continue;
                            if (!AnsIdToCol.ContainsKey(idAns)) continue;
                            dateBrute.Add(new RandSqlANS
                            {
                                NumeComplet = reader["NumeComplet"]?.ToString() ?? "",
                                Departament = reader["Departament"]?.ToString() ?? "",
                                Facultate = reader["Facultate"]?.ToString() ?? "",
                                GradFunctie = reader["GradFunctie"]?.ToString() ?? "",
                                OreConventionale = reader["OreConventionale"] != DBNull.Value ? Convert.ToDecimal(reader["OreConventionale"]) : 0m,
                                IdANS = idAns,
                            });
                        }
                    }
                }
            }

            // Grupare dupa NumeComplet - un profesor = un singur rand in ANS.
            // Departament/Facultate = cel cu SUM(ore) maxim (postul principal al profesorului).
            // GradFunctie = luat din acelasi departament principal.
            // Dublurile (Ungureanu Camelia asistent+conferentiar, etc.) sunt cauzate de
            // inregistrari cu DenTitularSauSuplinitor diferit in BD - filtrul '%tit%' le elimina.
            var profesori = new List<ProfANS>();
            foreach (var grp in dateBrute.GroupBy(x => x.NumeComplet))
            {
                // Gasim departamentul principal = cel cu cele mai multe ore totale
                var grupDept = grp
                    .GroupBy(x => new { x.Departament, x.Facultate })
                    .Select(dg => new {
                        dg.Key.Departament,
                        dg.Key.Facultate,
                        TotalOre = dg.Sum(x => x.OreConventionale),
                        // Gradul: luam primul rand din acest departament (deja filtrat cu '%tit%')
                        Grad = dg.OrderByDescending(x => x.OreConventionale).First().GradFunctie
                    })
                    .OrderByDescending(d => d.TotalOre)
                    .First();

                // Calculam orele per coloana ANS (suma tuturor orelor profesorului, indiferent de dept)
                var orePerCol = new Dictionary<int, decimal>();
                foreach (var rand in grp)
                {
                    int col = AnsIdToCol[rand.IdANS];
                    if (!orePerCol.ContainsKey(col)) orePerCol[col] = 0m;
                    orePerCol[col] += rand.OreConventionale;
                }

                decimal totalOre = orePerCol.Values.Sum();
                var fractiuni = new Dictionary<int, decimal>();
                if (totalOre > 0)
                {
                    // Coloana cu cele mai multe ore primeste restul pana la 1.00
                    // (pentru a evita erori de rotunjire)
                    int maxKey = orePerCol.OrderByDescending(x => x.Value).First().Key;
                    decimal sum = 0;
                    foreach (var kv in orePerCol)
                    {
                        if (kv.Key == maxKey) continue;
                        decimal frac = Math.Round(kv.Value / totalOre, 2);
                        fractiuni[kv.Key] = frac;
                        sum += frac;
                    }
                    fractiuni[maxKey] = Math.Round(1m - sum, 2);
                }

                profesori.Add(new ProfANS
                {
                    NumeComplet = grp.Key,
                    Departament = grupDept.Departament,
                    Facultate = grupDept.Facultate,
                    GradFunctie = MapareFunctieANS(grupDept.Grad),
                    Fractiuni = fractiuni,
                });
            }
            profesori = profesori.OrderBy(p => p.NumeComplet).ToList();

            var wb = BuildANSWorkbookFromScratch();
            var ws = wb.Worksheets.First();
            int dataStartRow = 9;

            if (profesori.Count > 0)
                ws.Row(dataStartRow).InsertRowsBelow(profesori.Count - 1);

            for (int i = 0; i < profesori.Count; i++)
            {
                var prof = profesori[i];
                int r = dataStartRow + i;

                ws.Cell(r, 1).Value = i + 1;
                ws.Cell(r, 2).Value = prof.NumeComplet;
                ws.Cell(r, 3).Value = "";
                ws.Cell(r, 4).Value = prof.GradFunctie;
                ws.Cell(r, 5).Value = 1;
                ws.Cell(r, 6).Value = 0;
                ws.Cell(r, 7).Value = "";
                ws.Cell(r, 8).Value = prof.Facultate;
                ws.Cell(r, 9).Value = prof.Departament;

                foreach (var kv in prof.Fractiuni)
                    ws.Cell(r, kv.Key).Value = kv.Value;

                string rowStr = r.ToString();
                ws.Cell(r, 50).FormulaA1 = $"=SUM(J{rowStr}:AW{rowStr})";

                if (i % 2 != 0)
                    for (int c = 1; c <= 50; c++)
                        ws.Cell(r, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5f5f5");
            }

            int newTotalRow = dataStartRow + profesori.Count;
            ws.Cell(newTotalRow, 1).Value = "Total general:";
            ws.Cell(newTotalRow, 1).Style.Font.Bold = true;
            for (int c = 10; c <= 49; c++)
            {
                string colLetter = ColumnLetter(c);
                ws.Cell(newTotalRow, c).FormulaA1 = $"=SUM({colLetter}{dataStartRow}:{colLetter}{newTotalRow - 1})";
                ws.Cell(newTotalRow, c).Style.Font.Bold = true;
            }
            ws.Cell(newTotalRow, 50).FormulaA1 = $"=SUM(J{newTotalRow}:AW{newTotalRow})";
            ws.Cell(newTotalRow, 50).Style.Font.Bold = true;

            using (var stream = new System.IO.MemoryStream())
            {
                wb.SaveAs(stream);
                wb.Dispose();
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Date_ANS_{idAnUniv}.xlsx");
            }
        }

        #endregion

        #region ================= HELPER METHODS =================

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

        private XLWorkbook BuildANSWorkbookFromScratch()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("CD DRU");

            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, 50).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov";
            ws.Range(3, 1, 3, 6).Merge();
            ws.Cell(4, 1).Value = "NOTA: Se includ in tabel toate cadrele didactice si de cercetare titulare (cu norma de baza in universitate), indiferent de forma de angajare.";
            ws.Range(4, 1, 4, 9).Merge();
            ws.Cell(4, 10).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta. Va rugam sa completati numai spatiile marcate cu culoarea verde.";
            ws.Range(4, 10, 4, 27).Merge();
            ws.Cell(4, 28).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta. Va rugam sa completati numai spatiile marcate cu culoarea verde.";
            ws.Range(4, 28, 4, 50).Merge();

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

            ws.Range(5, 1, 7, 1).Merge(); ws.Range(5, 2, 7, 2).Merge(); ws.Range(5, 3, 7, 3).Merge();
            ws.Range(5, 4, 7, 4).Merge(); ws.Range(5, 5, 7, 5).Merge(); ws.Range(5, 6, 7, 6).Merge();
            ws.Range(5, 7, 7, 7).Merge(); ws.Range(5, 8, 7, 8).Merge(); ws.Range(5, 9, 7, 9).Merge();
            ws.Range(5, 10, 5, 14).Merge(); ws.Range(5, 15, 5, 21).Merge();
            ws.Range(5, 22, 5, 27).Merge(); ws.Range(5, 28, 5, 36).Merge();
            ws.Range(5, 37, 5, 49).Merge(); ws.Range(5, 50, 7, 50).Merge();

            string[] subdomenii = {
                "Matematica", "Informatica", "Fizica", "Chimie si inginerie chimica", "Stiintele pamantului si atmosferei",
                "Inginerie civila", "Inginerie electrica, electronica si telecomunicatii",
                "Inginerie geologica, mine, petrol si gaze", "Ingineria transporturilor",
                "Ingineria resurselor vegetale si animale",
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

            for (int i = 0; i < 9; i++) ws.Cell(8, i + 1).Value = ((char)('A' + i)).ToString();
            for (int i = 0; i < 41; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, 50).Value = "40";

            var headerFill = XLColor.FromHtml(BrandColorHex);
            for (int r = 5; r <= 8; r++)
                for (int c = 1; c <= 50; c++)
                {
                    ws.Cell(r, c).Style.Font.Bold = true;
                    ws.Cell(r, c).Style.Font.FontColor = XLColor.Black;
                    ws.Cell(r, c).Style.Fill.BackgroundColor = XLColor.White;
                    ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(r, c).Style.Alignment.WrapText = true;
                    ws.Cell(r, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(r, c).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

            ws.Range(5, 1, 5, 50).Style.Fill.BackgroundColor = headerFill;
            ws.Range(5, 1, 5, 50).Style.Font.FontColor = XLColor.White;

            ws.Column(1).Width = 5; ws.Column(2).Width = 30; ws.Column(3).Width = 14;
            ws.Column(4).Width = 22; ws.Column(5).Width = 10; ws.Column(6).Width = 12;
            ws.Column(7).Width = 8; ws.Column(8).Width = 28; ws.Column(9).Width = 28;
            for (int c = 10; c <= 50; c++) ws.Column(c).Width = 12;

            return wb;
        }

        private string MapareFunctieANS(string grad)
        {
            if (string.IsNullOrWhiteSpace(grad)) return "Asistent";
            string g = grad.ToLower();
            if (g.Contains("profesor")) return "Profesor";
            if (g.Contains("conferentiar") || g.Contains("conferentiar")) return "Conferentiar";
            if (g.Contains("lector") || g.Contains("sef lucrari") || g.Contains("sl")) return "Lector/Sef de lucrari (SL)";
            if (g.Contains("asistent de cercetare")) return "Asistent de cercetare";
            if (g.Contains("asistent")) return "Asistent";
            if (g.Contains("preparator")) return "Preparator";
            if (g.Contains("cercetator stiintific i") || g.Contains("cs i")) return "Cercetator stiintific I (CS I)";
            if (g.Contains("cercetator stiintific ii") || g.Contains("cs ii")) return "Cercetator stiintific II (CS II)";
            if (g.Contains("cercetator stiintific iii") || g.Contains("cs iii")) return "Cercetator stiintific III (CS III)";
            if (g.Contains("cercetator")) return "Cercetator";
            return "Asistent";
        }

        private class ProfANS
        {
            public string NumeComplet { get; set; } = "";
            public string Departament { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string GradFunctie { get; set; } = "";
            public Dictionary<int, decimal> Fractiuni { get; set; } = new Dictionary<int, decimal>();
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

        #endregion
    }
}