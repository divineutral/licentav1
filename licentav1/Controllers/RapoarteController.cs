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

        // =====================================================================
        // SURSA UNICA DE DATE:
        //   [AGSIS].[pi].[StatDeFunctiiPeSpecializare]  sf   -> ore, materii, specializari
        //   [agsis_dw].[dbo].[View_ProfesoriActivi_CF]  prof -> NumeIntreg, Grad, Facultate, Catedra
        //   [AGSIS].[dbo].[AnUniversitar]               au   -> denumire an universitar
        //
        // JOIN: sf.ID_Profesor = prof.ID_Profesor  (nu pe ID_Post_Profesor_Materie!)
        // FILTRU AN: sf.id_anuniv = au.ID_AnUniv   (nu ppm.ID_AnUniv!)
        // =====================================================================

        // Constante SQL reutilizate
        private const string SqlJoinProf = @"
            FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
            LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
            LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv";

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

        private static readonly Dictionary<int, int> MappingMetaspec = new Dictionary<int, int>
        {
            { 20, 5 }, { 34, 5 }, { 44, 5 }, { 182, 5 }, { 837, 5 }, { 847, 5 }, { 848, 5 },
            { 43, 9 }, { 148, 8 }, { 171, 11 }, { 306, 11 }, { 358, 11 }, { 464, 11 },
            { 466, 11 }, { 615, 11 }, { 823, 11 }, { 828, 11 },
            { 35, 11 }, { 82, 7 }, { 84, 7 }, { 139, 11 }, { 162, 7 }, { 310, 7 },
            { 315, 7 }, { 316, 7 }, { 362, 7 }, { 372, 7 }, { 418, 7 }, { 597, 7 }, { 617, 7 },
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

        private static readonly Dictionary<int, int> AnsIdToCol = new Dictionary<int, int>
        {
            { 1, 10 }, { 2, 11 }, { 3, 12 }, { 4, 13 }, { 5, 14 },
            { 6, 15 }, { 7, 16 }, { 8, 17 }, { 9, 18 }, { 10, 19 }, { 11, 20 }, { 12, 21 },
            { 13, 22 }, { 14, 23 }, { 15, 24 }, { 16, 25 }, { 17, 26 }, { 18, 27 },
            { 19, 28 }, { 20, 29 }, { 21, 30 }, { 22, 31 }, { 23, 32 }, { 24, 33 },
            { 25, 34 }, { 26, 35 }, { 27, 36 },
            { 28, 37 }, { 29, 38 }, { 30, 39 }, { 31, 40 }, { 32, 41 }, { 33, 42 },
            { 34, 43 }, { 35, 44 }, { 36, 45 }, { 37, 46 }, { 38, 47 }, { 39, 48 }, { 40, 49 },
        };

        public RapoarteController(IConfiguration configuration, IMemoryCache cache)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        // =====================================================================
        // HELPER: adauga parametrii comuni de filtrare
        // =====================================================================
        private void AdaugaParametriFiltru(SqlCommand cmd, string anUniv, string facultate, string departament, string profesor, string specializari, int semestru, string tipPost, string formaInvatamant)
        {
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
            cmd.Parameters.AddWithValue("@sem", semestru);
            cmd.Parameters.AddWithValue("@tip", tipPost ?? "Toti");
            cmd.Parameters.AddWithValue("@forma", formaInvatamant ?? "Toti");
        }

        // Clauza WHERE comuna pentru toate rapoartele
        private const string SqlWhereFiltru = @"
            WHERE
                (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
            AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
            AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
            AND (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.NumeIntreg,'')))) = UPPER(LTRIM(RTRIM(@prof))))
            AND (@specs = 'Toti' OR UPPER(LTRIM(RTRIM(sf.DenumireSpecializare))) IN (SELECT UPPER(LTRIM(RTRIM(value))) FROM STRING_SPLIT(@specs,',')))
            AND (@sem  = 0    OR sf.NrSemestruDinAn = @sem)
            AND (@tip  = 'Toti' OR sf.DenTitularSauSuplinitor = @tip)
            AND (@forma = 'Toti' OR sf.DenumireSpecializare LIKE '% ' + @forma + '%' OR sf.DenumireSpecializare LIKE '%-' + @forma + '%')";

        #region ================= LISTE (DROPDOWNS) =================

        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAniUniversitari()
        {
            var lista = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT au.ID_AnUniv AS id, au.Denumire AS nume
                FROM [AGSIS].[dbo].[AnUniversitar] au
                WHERE au.ID_AnUniv IN (
                    SELECT DISTINCT id_anuniv FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare]
                )
                ORDER BY au.ID_AnUniv DESC", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                lista.Add(new { id = r["id"].ToString(), nume = r["nume"].ToString() });
            return Ok(lista);
        }

        [HttpGet("liste/facultati")]
        public ActionResult GetFacultati(string? anUniv)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) AS Fac
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE prof.DenumireFacultate IS NOT NULL
                  AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                ORDER BY Fac", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r["Fac"]?.ToString(); if (!string.IsNullOrWhiteSpace(v)) lista.Add(v); }
            return Ok(lista);
        }

        [HttpGet("liste/departamente")]
        public ActionResult GetDepartamente(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) AS Dept
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE prof.DenumireCatedra IS NOT NULL
                  AND (@an  = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                ORDER BY Dept", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r["Dept"]?.ToString(); if (!string.IsNullOrWhiteSpace(v)) lista.Add(v); }
            return Ok(lista);
        }

        [HttpGet("liste/specializari-per-facultate")]
        public ActionResult GetSpecializari(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT UPPER(LTRIM(RTRIM(
                    CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0
                         THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                         ELSE sf.DenumireSpecializare END
                ))) AS Spec
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE sf.DenumireSpecializare IS NOT NULL
                  AND (@an  = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                ORDER BY Spec", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r["Spec"]?.ToString(); if (!string.IsNullOrWhiteSpace(v) && !lista.Contains(v)) lista.Add(v); }
            return Ok(lista);
        }

        [HttpGet("liste/profesori-per-specializari")]
        public ActionResult GetProfesoriPerSpecializari(string? anUniv, string? facultate, string? specializari, string? departament)
        {
            var lista = new List<string>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT ISNULL(prof.NumeIntreg, '') AS NumeIntreg
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE prof.NumeIntreg IS NOT NULL
                  AND (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                ORDER BY NumeIntreg", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r["NumeIntreg"]?.ToString(); if (!string.IsNullOrWhiteSpace(v)) lista.Add(v); }
            return Ok(lista);
        }

        #endregion

        #region ================= RAPORT 1: NORMA PROFESORI =================

        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string? anUniv, string? facultate, string? specializari,
            string? profesor, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT
                    ISNULL(prof.NumeIntreg,'Prof. ID ' + CAST(sf.ID_Profesor AS VARCHAR)) AS NumeProfesor,
                    sf.DenumireSpecializare,
                    sf.DenumireMaterie,
                    sf.DenTitularSauSuplinitor AS TipPost,
                    sf.NrSemestruDinAn,
                    ISNULL(sf.Nr_Ore_Curs,0)       AS Nr_Ore_Curs,
                    ISNULL(sf.Nr_Ore_Seminar,0) + ISNULL(sf.Nr_Ore_Laborator,0) + ISNULL(sf.Nr_Ore_Proiect,0) + ISNULL(sf.Nr_Ore_SF,0) AS OreAplic,
                    ISNULL(sf.NrOreConventionale,0) AS NrOreConventionale"
                + SqlJoinFiltru() + SqlWhereFiltru + @"
                ORDER BY NumeProfesor, sf.DenumireSpecializare, sf.DenumireMaterie";

            using var cmd = new SqlCommand(sql, conn);
            AdaugaParametriFiltru(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add(new
                {
                    Profesor = r["NumeProfesor"].ToString(),
                    Specializare = r["DenumireSpecializare"].ToString(),
                    Materie = r["DenumireMaterie"].ToString(),
                    TipPost = r["TipPost"].ToString(),
                    Semestru = Convert.ToInt32(r["NrSemestruDinAn"]),
                    NrOreCurs = Convert.ToDouble(r["Nr_Ore_Curs"]),
                    NrOreAplicatii = Convert.ToDouble(r["OreAplic"]),
                    NrOreConventionale = Convert.ToDouble(r["NrOreConventionale"])
                });
            return Ok(result);
        }

        [HttpGet("export/norma")]
        public IActionResult ExportNormaExcel(string? anUniv, string? facultate, string? specializari,
            string? profesor, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"), new DataColumn("Specializare"), new DataColumn("Materie"),
                new DataColumn("Tip Post"), new DataColumn("Semestru"),
                new DataColumn("Nr Ore Curs", typeof(double)),
                new DataColumn("Nr Ore Aplicatii", typeof(double)),
                new DataColumn("Nr Ore Conventionale", typeof(double))
            });

            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT
                    ISNULL(prof.NumeIntreg,'Prof. ID ' + CAST(sf.ID_Profesor AS VARCHAR)) AS NumeProfesor,
                    sf.DenumireSpecializare,
                    ISNULL(sf.DenumireMaterie,'Nedefinit') AS DenumireMaterie,
                    ISNULL(sf.DenTitularSauSuplinitor,'Nespecificat') AS TipPost,
                    ISNULL(sf.NrSemestruDinAn,0) AS NrSemestruDinAn,
                    ISNULL(sf.Nr_Ore_Curs,0) AS Nr_Ore_Curs,
                    ISNULL(sf.Nr_Ore_Seminar,0)+ISNULL(sf.Nr_Ore_Laborator,0)+ISNULL(sf.Nr_Ore_Proiect,0)+ISNULL(sf.Nr_Ore_SF,0) AS OreAplic,
                    ISNULL(sf.NrOreConventionale,0) AS NrOreConventionale"
                + SqlJoinFiltru() + SqlWhereFiltru + @"
                ORDER BY NumeProfesor, sf.DenumireSpecializare, sf.DenumireMaterie";

            using var cmd = new SqlCommand(sql, conn);
            AdaugaParametriFiltru(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(r["NumeProfesor"], r["DenumireSpecializare"], r["DenumireMaterie"], r["TipPost"],
                    r["NrSemestruDinAn"], Convert.ToDouble(r["Nr_Ore_Curs"]),
                    Convert.ToDouble(r["OreAplic"]), Convert.ToDouble(r["NrOreConventionale"]));

            string fileName = (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                ? $"NormaProfesori_{SanitizeFileName(profesor)}.xlsx"
                : "NormaProfesori_General.xlsx";

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme");
            ws.Cell(1, 1).Value = $"An: {anUniv} | Facultate: {facultate} | Departament: {departament} | Profesor: {profesor}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            StileazaHeader(ws, 3, dt.Columns.Count);
            tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Nr Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Aplicatii").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Conventionale").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            return ExcelResponse(wb, fileName);
        }

        #endregion

        #region ================= RAPORT 2: ORE PE PROGRAM =================

        [HttpGet("ore-profesor-program")]
        public ActionResult GetOreProfProgram(string? anUniv, string? facultate, string? specializari,
            string? profesor, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                WITH Filtrat AS (
                    SELECT
                        ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                        UPPER(LTRIM(RTRIM(
                            CASE WHEN CHARINDEX('+',sf.DenumireSpecializare)>0
                                 THEN LEFT(sf.DenumireSpecializare,CHARINDEX('+',sf.DenumireSpecializare)-1)
                                 ELSE sf.DenumireSpecializare END))) AS ProgramStudiu,
                        ISNULL(sf.NrOreConventionale,0) AS OreConv"
                + SqlJoinFiltru() + SqlWhereFiltru + @"
                ),
                Grupat AS (
                    SELECT Profesor, ProgramStudiu, SUM(OreConv) AS OreConvProgram
                    FROM Filtrat GROUP BY Profesor, ProgramStudiu
                ),
                Total AS (SELECT Profesor, SUM(OreConvProgram) AS TotalPost FROM Grupat GROUP BY Profesor)
                SELECT g.Profesor, ISNULL(g.ProgramStudiu,'Nespecificat') AS ProgramStudiu,
                    g.OreConvProgram AS NrOreConv, t.TotalPost,
                    CAST(CASE WHEN t.TotalPost=0 THEN 0 ELSE (g.OreConvProgram/t.TotalPost)*100 END AS DECIMAL(10,2)) AS ProcentPost
                FROM Grupat g INNER JOIN Total t ON g.Profesor=t.Profesor
                ORDER BY g.Profesor, g.OreConvProgram DESC";

            using var cmd = new SqlCommand(sql, conn);
            AdaugaParametriFiltru(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add(new
                {
                    Profesor = r["Profesor"].ToString(),
                    ProgramStudiu = r["ProgramStudiu"].ToString(),
                    NrOreConv = Convert.ToDouble(r["NrOreConv"]),
                    TotalPost = Convert.ToDouble(r["TotalPost"]),
                    ProcentPost = Convert.ToDouble(r["ProcentPost"])
                });
            return Ok(result);
        }

        [HttpGet("export/ore-program")]
        public IActionResult ExportOreProgramExcel(string? anUniv, string? facultate, string? specializari,
            string? profesor, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"), new DataColumn("Program Studiu"),
                new DataColumn("Nr Ore Conv", typeof(double)),
                new DataColumn("Procent Post", typeof(double)),
                new DataColumn("Total Post", typeof(double))
            });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                WITH Filtrat AS (
                    SELECT
                        ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                        UPPER(LTRIM(RTRIM(
                            CASE WHEN CHARINDEX('+',sf.DenumireSpecializare)>0
                                 THEN LEFT(sf.DenumireSpecializare,CHARINDEX('+',sf.DenumireSpecializare)-1)
                                 ELSE sf.DenumireSpecializare END))) AS ProgramStudiu,
                        ISNULL(sf.NrOreConventionale,0) AS OreConv"
                + SqlJoinFiltru() + SqlWhereFiltru + @"
                ),
                Grupat AS (
                    SELECT Profesor, ProgramStudiu, SUM(OreConv) AS OreConvProgram
                    FROM Filtrat GROUP BY Profesor, ProgramStudiu
                ),
                Total AS (SELECT Profesor, SUM(OreConvProgram) AS TotalPost FROM Grupat GROUP BY Profesor)
                SELECT g.Profesor, ISNULL(g.ProgramStudiu,'Nespecificat') AS ProgramStudiu,
                    g.OreConvProgram, t.TotalPost,
                    CAST(CASE WHEN t.TotalPost=0 THEN 0 ELSE (g.OreConvProgram/t.TotalPost)*100 END AS DECIMAL(10,2)) AS ProcentPost
                FROM Grupat g INNER JOIN Total t ON g.Profesor=t.Profesor
                ORDER BY g.Profesor, g.OreConvProgram DESC";
            using var cmd = new SqlCommand(sql, conn);
            AdaugaParametriFiltru(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(r["Profesor"], r["ProgramStudiu"], Convert.ToDouble(r["OreConvProgram"]),
                    Convert.ToDouble(r["ProcentPost"]), Convert.ToDouble(r["TotalPost"]));

            string fileName = (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                ? $"StatisticaOre_{SanitizeFileName(profesor)}.xlsx"
                : "StatisticaOre_General.xlsx";
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Distributie Ore");
            var tbl = ws.Cell(1, 1).InsertTable(dt);
            StileazaHeader(ws, 1, dt.Columns.Count);
            tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            return ExcelResponse(wb, fileName);
        }

        #endregion

        #region ================= RAPORT 3: TOTALURI NORME =================

        [HttpGet("norma-totaluri")]
        public ActionResult GetNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                WITH DateProf AS (
                    SELECT
                        ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS NumeComplet,
                        ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                        ISNULL(prof.DenumireFacultate,'Nespecificat') AS Facultate,
                        ISNULL(sf.NrOreConventionale,0) AS OreConv
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                    LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                    WHERE
                        (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                    AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                    AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                    AND (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.NumeIntreg,'')))) = UPPER(LTRIM(RTRIM(@prof))))
                ),
                Totale AS (
                    SELECT NumeComplet, SUM(OreConv) AS TotalOreConv, SUM(OreConv)*14 AS TotalAnual
                    FROM DateProf GROUP BY NumeComplet
                ),
                DeptPrinc AS (
                    SELECT NumeComplet, Departament, Facultate,
                           ROW_NUMBER() OVER(PARTITION BY NumeComplet ORDER BY SUM(OreConv) DESC) AS Rang
                    FROM DateProf GROUP BY NumeComplet, Departament, Facultate
                )
                SELECT t.NumeComplet, d.Departament, d.Facultate, t.TotalOreConv, t.TotalAnual
                FROM Totale t INNER JOIN DeptPrinc d ON t.NumeComplet=d.NumeComplet AND d.Rang=1
                ORDER BY t.NumeComplet";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add(new
                {
                    Profesor = r["NumeComplet"].ToString(),
                    Departament = r["Departament"].ToString(),
                    Facultate = r["Facultate"].ToString(),
                    TotalOreConv = Math.Round(Convert.ToDecimal(r["TotalOreConv"]), 2),
                    TotalAnualOreConv = Math.Round(Convert.ToDecimal(r["TotalAnual"]), 2)
                });
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public IActionResult ExportNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume Profesor"), new DataColumn("Departament"), new DataColumn("Facultate"),
                new DataColumn("Total Ore Conv.", typeof(decimal)),
                new DataColumn("Total tot anul ore conv.", typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                WITH DateProf AS (
                    SELECT
                        ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS NumeComplet,
                        ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                        ISNULL(prof.DenumireFacultate,'Nespecificat') AS Facultate,
                        ISNULL(sf.NrOreConventionale,0) AS OreConv
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                    LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                    WHERE
                        (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                    AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                    AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                    AND (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.NumeIntreg,'')))) = UPPER(LTRIM(RTRIM(@prof))))
                ),
                Totale AS (
                    SELECT NumeComplet, SUM(OreConv) AS TotalOreConv, SUM(OreConv)*14 AS TotalAnual
                    FROM DateProf GROUP BY NumeComplet
                ),
                DeptPrinc AS (
                    SELECT NumeComplet, Departament, Facultate,
                           ROW_NUMBER() OVER(PARTITION BY NumeComplet ORDER BY SUM(OreConv) DESC) AS Rang
                    FROM DateProf GROUP BY NumeComplet, Departament, Facultate
                )
                SELECT t.NumeComplet, d.Departament, d.Facultate, t.TotalOreConv, t.TotalAnual
                FROM Totale t INNER JOIN DeptPrinc d ON t.NumeComplet=d.NumeComplet AND d.Rang=1
                ORDER BY t.NumeComplet";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(r["NumeComplet"], r["Departament"], r["Facultate"],
                    Math.Round(Convert.ToDecimal(r["TotalOreConv"]), 2),
                    Math.Round(Convert.ToDecimal(r["TotalAnual"]), 2));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Totaluri Norme");
            var tbl = ws.Cell(1, 1).InsertTable(dt);
            StileazaHeader(ws, 1, dt.Columns.Count);
            tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Total Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total tot anul ore conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nume Profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            return ExcelResponse(wb, "Totaluri_Norme.xlsx");
        }

        #endregion

        #region ================= RAPORT 4: LIMBI STRAINE =================

        [HttpGet("limbi-straine")]
        public ActionResult GetLimbiStraine(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                WITH Baza AS (
                    SELECT
                        ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS NumeComplet,
                        sf.NrSemestruDinAn AS Semestru,
                        ISNULL(sf.NrOreConventionale,0) AS OreConv,
                        sf.DenumireSpecializare
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                    LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                    WHERE
                        (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                    AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                    AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                    AND (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.NumeIntreg,'')))) = UPPER(LTRIM(RTRIM(@prof))))
                    AND (@specs = 'Toti' OR UPPER(LTRIM(RTRIM(sf.DenumireSpecializare))) IN (SELECT UPPER(LTRIM(RTRIM(value))) FROM STRING_SPLIT(@specs,',')))
                    AND (@sem  = 0 OR sf.NrSemestruDinAn = @sem)
                    AND (@tip  = 'Toti' OR sf.DenTitularSauSuplinitor = @tip)
                    AND (
                        sf.DenumireSpecializare LIKE '%englez%' OR sf.DenumireSpecializare LIKE '%francez%'
                        OR sf.DenumireSpecializare LIKE '%german%' OR sf.DenumireSpecializare LIKE '%american%'
                        OR sf.DenumireSpecializare LIKE '%(EN)%'  OR sf.DenumireSpecializare LIKE '%(FR)%'
                        OR sf.DenumireSpecializare LIKE '%(G)%'
                        OR sf.DenumireSpecializare IN (
                            'Inginerie virtuala in proiectarea autovehiculelor',
                            'Metode practice integrate in ingineria sistemelor de propulsie',
                            'Ingineria proceselor de fabricatie avansate',
                            'Managementul afacerilor industriale si antreprenoriat',
                            'Inginerie electrica si calculatoare','Sisteme electrice avansate',
                            'Securitate cibernetica','Informatica aplicata','Tehnologii Internet',
                            'Cultura si discurs in spatiul anglo american',
                            'Studii de limba si de cultura franceza',
                            'Studii de limba si literatura germana din perspectiva interculturala',
                            'Studii lingvistice pentru comunicare interculturala',
                            'Traducere si interpretariat din limba franceza in limba romana',
                            'Studii americane','Performanta umana in antrenamentul sportiv',
                            'Administrarea afacerilor','Managementul resurselor umane',
                            'Dezvoltarea afacerilor turistice','Medicina traditionala chineza'
                        )
                    )
                )
                SELECT NumeComplet,
                    SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConv ELSE 0 END)*14 AS Sem1,
                    SUM(CASE WHEN Semestru IN (2,4,6,8,10,12) THEN OreConv ELSE 0 END)*14 AS Sem2,
                    SUM(OreConv)*14 AS Total
                FROM Baza
                GROUP BY NumeComplet
                HAVING SUM(OreConv) > 0
                ORDER BY NumeComplet";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
            cmd.Parameters.AddWithValue("@sem", semestru);
            cmd.Parameters.AddWithValue("@tip", tipPost ?? "Toti");
            using var r = cmd.ExecuteReader();
            int nrCrt = 1;
            while (r.Read())
                result.Add(new
                {
                    NrCrt = nrCrt++,
                    NumeProfesor = r["NumeComplet"].ToString(),
                    Sem1 = r["Sem1"],
                    Sem2 = r["Sem2"],
                    Total = r["Total"]
                });
            return Ok(result);
        }

        [HttpGet("export/limbi-straine")]
        public IActionResult ExportLimbiStraine(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nr. Crt.", typeof(int)),
                new DataColumn("Nume si prenume profesor"),
                new DataColumn("Total Sem 1", typeof(decimal)),
                new DataColumn("Total Sem 2", typeof(decimal)),
                new DataColumn("Total", typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                WITH Baza AS (
                    SELECT
                        ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS NumeComplet,
                        sf.NrSemestruDinAn AS Semestru,
                        ISNULL(sf.NrOreConventionale,0) AS OreConv,
                        sf.DenumireSpecializare
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                    LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                    WHERE
                        (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                    AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                    AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                    AND (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.NumeIntreg,'')))) = UPPER(LTRIM(RTRIM(@prof))))
                    AND (@specs = 'Toti' OR UPPER(LTRIM(RTRIM(sf.DenumireSpecializare))) IN (SELECT UPPER(LTRIM(RTRIM(value))) FROM STRING_SPLIT(@specs,',')))
                    AND (@sem  = 0 OR sf.NrSemestruDinAn = @sem)
                    AND (@tip  = 'Toti' OR sf.DenTitularSauSuplinitor = @tip)
                    AND (
                        sf.DenumireSpecializare LIKE '%englez%' OR sf.DenumireSpecializare LIKE '%francez%'
                        OR sf.DenumireSpecializare LIKE '%german%' OR sf.DenumireSpecializare LIKE '%american%'
                        OR sf.DenumireSpecializare LIKE '%(EN)%'  OR sf.DenumireSpecializare LIKE '%(FR)%'
                        OR sf.DenumireSpecializare LIKE '%(G)%'
                        OR sf.DenumireSpecializare IN (
                            'Inginerie virtuala in proiectarea autovehiculelor',
                            'Metode practice integrate in ingineria sistemelor de propulsie',
                            'Ingineria proceselor de fabricatie avansate',
                            'Managementul afacerilor industriale si antreprenoriat',
                            'Inginerie electrica si calculatoare','Sisteme electrice avansate',
                            'Securitate cibernetica','Informatica aplicata','Tehnologii Internet',
                            'Cultura si discurs in spatiul anglo american',
                            'Studii de limba si de cultura franceza',
                            'Studii de limba si literatura germana din perspectiva interculturala',
                            'Studii lingvistice pentru comunicare interculturala',
                            'Traducere si interpretariat din limba franceza in limba romana',
                            'Studii americane','Performanta umana in antrenamentul sportiv',
                            'Administrarea afacerilor','Managementul resurselor umane',
                            'Dezvoltarea afacerilor turistice','Medicina traditionala chineza'
                        )
                    )
                )
                SELECT NumeComplet,
                    CAST(SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConv ELSE 0 END)*14 AS DECIMAL(10,2)) AS Sem1,
                    CAST(SUM(CASE WHEN Semestru IN (2,4,6,8,10,12) THEN OreConv ELSE 0 END)*14 AS DECIMAL(10,2)) AS Sem2,
                    CAST(SUM(OreConv)*14 AS DECIMAL(10,2)) AS Total
                FROM Baza
                GROUP BY NumeComplet
                HAVING SUM(OreConv) > 0
                ORDER BY NumeComplet";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
            cmd.Parameters.AddWithValue("@sem", semestru);
            cmd.Parameters.AddWithValue("@tip", tipPost ?? "Toti");
            using var r = cmd.ExecuteReader();
            int nrCrt = 1;
            while (r.Read())
                dt.Rows.Add(nrCrt++, r["NumeComplet"], r["Sem1"], r["Sem2"], r["Total"]);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Limbi Straine");
            var tbl = ws.Cell(1, 1).InsertTable(dt);
            StileazaHeader(ws, 1, dt.Columns.Count);
            tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Total Sem 1").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total Sem 2").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nume si prenume profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            return ExcelResponse(wb, "Raport_Limbi_Straine.xlsx");
        }

        #endregion

        #region ================= RAPORT 5: DISCIPLINE PREDATE =================

        [HttpGet("discipline-predate")]
        public ActionResult GetDisciplinePredate(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT DISTINCT
                    ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(sf.DenumireMaterie,'Nespecificat') AS Materie,
                    LTRIM(STUFF(
                        CASE WHEN ISNULL(sf.Nr_Ore_Curs,0)>0      THEN ', Curs'      ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_Seminar,0)>0   THEN ', Seminar'   ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_Laborator,0)>0 THEN ', Laborator' ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_Proiect,0)>0   THEN ', Proiect'   ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_SF,0)>0        THEN ', SF'        ELSE '' END
                    ,1,2,'')) AS TipActivitate"
                + SqlJoinFiltru() + SqlWhereFiltru + @"
                ORDER BY Profesor, Materie";

            using var cmd = new SqlCommand(sql, conn);
            AdaugaParametriFiltru(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add(new
                {
                    Profesor = r["Profesor"].ToString(),
                    Departament = r["Departament"].ToString(),
                    Materie = r["Materie"].ToString(),
                    TipActivitate = r["TipActivitate"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/discipline-predate")]
        public IActionResult ExportDisciplinePredate(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0, string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"), new DataColumn("Departament"),
                new DataColumn("Disciplina"), new DataColumn("Tip Activitate")
            });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT DISTINCT
                    ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(sf.DenumireMaterie,'Nespecificat') AS Materie,
                    LTRIM(STUFF(
                        CASE WHEN ISNULL(sf.Nr_Ore_Curs,0)>0      THEN ', Curs'      ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_Seminar,0)>0   THEN ', Seminar'   ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_Laborator,0)>0 THEN ', Laborator' ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_Proiect,0)>0   THEN ', Proiect'   ELSE '' END +
                        CASE WHEN ISNULL(sf.Nr_Ore_SF,0)>0        THEN ', SF'        ELSE '' END
                    ,1,2,'')) AS TipActivitate"
                + SqlJoinFiltru() + SqlWhereFiltru + @"
                ORDER BY Profesor, Materie";
            using var cmd = new SqlCommand(sql, conn);
            AdaugaParametriFiltru(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(r["Profesor"], r["Departament"], r["Materie"], r["TipActivitate"]?.ToString() ?? "");

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Discipline Predate");
            var tbl = ws.Cell(1, 1).InsertTable(dt);
            StileazaHeader(ws, 1, dt.Columns.Count);
            tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            return ExcelResponse(wb, "Discipline_Predate.xlsx");
        }

        #endregion

        #region ================= RAPORT 6: TITULARI =================

        [HttpGet("titulari")]
        public ActionResult GetTitulari(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT DISTINCT
                    ISNULL(prof.NumeIntreg,'') AS NumeComplet,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(prof.DenumireFacultate,'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE sf.DenTitularSauSuplinitor = 'Tit'
                  AND prof.NumeIntreg IS NOT NULL
                  AND (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                ORDER BY NumeComplet";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add(new { Profesor = r["NumeComplet"].ToString(), Departament = r["Departament"].ToString(), Facultate = r["Facultate"].ToString() });
            return Ok(result);
        }

        [HttpGet("export/titulari")]
        public IActionResult ExportTitulari(string? anUniv, string? facultate, string? departament)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] { new DataColumn("Nume si prenume"), new DataColumn("Departament"), new DataColumn("Facultate") });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT DISTINCT
                    ISNULL(prof.NumeIntreg,'') AS NumeComplet,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(prof.DenumireFacultate,'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE sf.DenTitularSauSuplinitor = 'Tit'
                  AND prof.NumeIntreg IS NOT NULL
                  AND (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                ORDER BY NumeComplet";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) dt.Rows.Add(r["NumeComplet"], r["Departament"], r["Facultate"]);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Titulari");
            var tbl = ws.Cell(1, 1).InsertTable(dt);
            StileazaHeader(ws, 1, dt.Columns.Count);
            tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            return ExcelResponse(wb, "Cadre_Didactice_Titulare.xlsx");
        }

        #endregion

        #region ================= RAPORT 7: COLABORATORI =================

        [HttpGet("colaboratori")]
        public ActionResult GetColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT DISTINCT
                    ISNULL(prof.NumeIntreg,'') AS NumeComplet,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(prof.DenumireFacultate,'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE sf.DenTitularSauSuplinitor = 'Sup'
                  AND prof.NumeIntreg IS NOT NULL
                  AND (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                  AND sf.ID_Profesor NOT IN (
                      SELECT DISTINCT sf2.ID_Profesor
                      FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf2
                      LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au2 ON sf2.id_anuniv = au2.ID_AnUniv
                      WHERE sf2.DenTitularSauSuplinitor = 'Tit'
                        AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au2.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  )
                ORDER BY NumeComplet";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add(new { Profesor = r["NumeComplet"].ToString(), Departament = r["Departament"].ToString(), Facultate = r["Facultate"].ToString() });
            return Ok(result);
        }

        [HttpGet("export/colaboratori")]
        public IActionResult ExportColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] { new DataColumn("Nume si prenume"), new DataColumn("Departament"), new DataColumn("Facultate") });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string sql = @"
                SELECT DISTINCT
                    ISNULL(prof.NumeIntreg,'') AS NumeComplet,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(prof.DenumireFacultate,'Nespecificat') AS Facultate
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE sf.DenTitularSauSuplinitor = 'Sup'
                  AND prof.NumeIntreg IS NOT NULL
                  AND (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                  AND sf.ID_Profesor NOT IN (
                      SELECT DISTINCT sf2.ID_Profesor
                      FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf2
                      LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au2 ON sf2.id_anuniv = au2.ID_AnUniv
                      WHERE sf2.DenTitularSauSuplinitor = 'Tit'
                        AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au2.Denumire))) = UPPER(LTRIM(RTRIM(@an))))
                  )
                ORDER BY NumeComplet";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) dt.Rows.Add(r["NumeComplet"], r["Departament"], r["Facultate"]);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Colaboratori");
            var tbl = ws.Cell(1, 1).InsertTable(dt);
            StileazaHeader(ws, 1, dt.Columns.Count);
            tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            return ExcelResponse(wb, "Cadre_Didactice_Colaboratori.xlsx");
        }

        #endregion

        #region ================= RAPORT 8: ANS =================

        [HttpGet("date-ans")]
        public IActionResult GetDateANS([FromQuery] int idAnUniv = 45)
        {
            var dateBrute = new List<RandSqlANS>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT
                    ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS NumeComplet,
                    ISNULL(prof.DenumireGradDidactic,'') AS GradFunctie,
                    ISNULL(prof.DenumireFacultate,'Nespecificat') AS Facultate,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(sf.NrOreConventionale,0) AS OreConventionale,
                    sf.id_metaspecializare AS IdMetaspec
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                WHERE sf.id_anuniv = @ID_AnUniv
                  AND sf.DenTitularSauSuplinitor = 'Tit'", conn);
            cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                int idMeta = r["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(r["IdMetaspec"]) : 0;
                if (!MappingMetaspec.TryGetValue(idMeta, out int idAns)) continue;
                if (!AnsIdToCol.ContainsKey(idAns)) continue;
                dateBrute.Add(new RandSqlANS
                {
                    NumeComplet = r["NumeComplet"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    GradFunctie = r["GradFunctie"]?.ToString() ?? "",
                    OreConventionale = Convert.ToDecimal(r["OreConventionale"]),
                    IdANS = idAns
                });
            }

            var profesori = dateBrute.GroupBy(x => x.NumeComplet).Select(g => {
                var grupDept = g.GroupBy(x => new { x.Departament, x.Facultate })
                    .Select(dg => new { dg.Key.Departament, dg.Key.Facultate, TotalOre = dg.Sum(x => x.OreConventionale), Grad = dg.OrderByDescending(x => x.OreConventionale).First().GradFunctie })
                    .OrderByDescending(d => d.TotalOre).First();
                var orePerAns = g.GroupBy(x => x.IdANS).ToDictionary(ag => ag.Key, ag => ag.Sum(x => x.OreConventionale));
                decimal totalOre = orePerAns.Values.Sum();
                var fractiuni = new Dictionary<string, decimal>();
                if (totalOre > 0)
                {
                    int maxKey = orePerAns.OrderByDescending(x => x.Value).First().Key;
                    decimal sum = 0;
                    foreach (var kv in orePerAns) { if (kv.Key == maxKey) continue; decimal frac = Math.Round(kv.Value / totalOre, 2); fractiuni[DomeniiExcel[AnsIdToCol[kv.Key] - 10]] = frac; sum += frac; }
                    fractiuni[DomeniiExcel[AnsIdToCol[maxKey] - 10]] = Math.Round(1m - sum, 2);
                }
                return new { NumeComplet = g.Key, Facultate = grupDept.Facultate, Departament = grupDept.Departament, GradFunctie = MapareFunctieANS(grupDept.Grad), DomeniiMapate = fractiuni };
            }).OrderBy(p => p.NumeComplet).ToList();

            return Ok(profesori);
        }

        [HttpGet("export/raport-ans")]
        public IActionResult ExportRaportANS([FromQuery] int idAnUniv = 45)
        {
            var dateBrute = new List<RandSqlANS>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT
                    ISNULL(prof.NumeIntreg,'Prof. ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS NumeComplet,
                    ISNULL(prof.DenumireGradDidactic,'') AS GradFunctie,
                    ISNULL(prof.DenumireCatedra,'Nespecificat') AS Departament,
                    ISNULL(prof.DenumireFacultate,'') AS Facultate,
                    ISNULL(sf.NrOreConventionale,0) AS OreConventionale,
                    sf.id_metaspecializare AS IdMetaspec
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                WHERE sf.id_anuniv = @ID_AnUniv
                  AND sf.DenTitularSauSuplinitor = 'Tit'", conn);
            cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                int idMeta = r["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(r["IdMetaspec"]) : 0;
                if (!MappingMetaspec.TryGetValue(idMeta, out int idAns)) continue;
                if (!AnsIdToCol.ContainsKey(idAns)) continue;
                dateBrute.Add(new RandSqlANS
                {
                    NumeComplet = r["NumeComplet"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    GradFunctie = r["GradFunctie"]?.ToString() ?? "",
                    OreConventionale = Convert.ToDecimal(r["OreConventionale"]),
                    IdANS = idAns
                });
            }

            var profesori = new List<ProfANS>();
            foreach (var grp in dateBrute.GroupBy(x => x.NumeComplet))
            {
                var grupDept = grp.GroupBy(x => new { x.Departament, x.Facultate })
                    .Select(dg => new { dg.Key.Departament, dg.Key.Facultate, TotalOre = dg.Sum(x => x.OreConventionale), Grad = dg.OrderByDescending(x => x.OreConventionale).First().GradFunctie })
                    .OrderByDescending(d => d.TotalOre).First();
                var orePerCol = new Dictionary<int, decimal>();
                foreach (var rand in grp) { int col = AnsIdToCol[rand.IdANS]; if (!orePerCol.ContainsKey(col)) orePerCol[col] = 0m; orePerCol[col] += rand.OreConventionale; }
                decimal totalOre = orePerCol.Values.Sum();
                var fractiuni = new Dictionary<int, decimal>();
                if (totalOre > 0)
                {
                    int maxKey = orePerCol.OrderByDescending(x => x.Value).First().Key;
                    decimal sum = 0;
                    foreach (var kv in orePerCol) { if (kv.Key == maxKey) continue; decimal frac = Math.Round(kv.Value / totalOre, 2); fractiuni[kv.Key] = frac; sum += frac; }
                    fractiuni[maxKey] = Math.Round(1m - sum, 2);
                }
                profesori.Add(new ProfANS { NumeComplet = grp.Key, Departament = grupDept.Departament, Facultate = grupDept.Facultate, GradFunctie = MapareFunctieANS(grupDept.Grad), Fractiuni = fractiuni });
            }

            // Override manual valori confirmate Diana Ionita
            var overrides = new Dictionary<string, Dictionary<int, decimal>>
            {
                ["VOLMER MARIUS"] = new Dictionary<int, decimal> { { AnsIdToCol[7], 0.83m }, { AnsIdToCol[12], 0.17m } },
                ["ZAHARIA SEBASTIAN MARIAN"] = new Dictionary<int, decimal> { { AnsIdToCol[12], 0.74m }, { AnsIdToCol[9], 0.27m } },
            };
            foreach (var prof in profesori)
                if (overrides.TryGetValue(prof.NumeComplet, out var ov)) prof.Fractiuni = ov;

            profesori = profesori.OrderBy(p => p.NumeComplet).ToList();

            var wb = BuildANSWorkbookFromScratch();
            var ws = wb.Worksheets.First();
            int dataStartRow = 9;
            if (profesori.Count > 0) ws.Row(dataStartRow).InsertRowsBelow(profesori.Count - 1);

            for (int i = 0; i < profesori.Count; i++)
            {
                var prof = profesori[i];
                int rr = dataStartRow + i;
                ws.Cell(rr, 1).Value = i + 1;
                ws.Cell(rr, 2).Value = prof.NumeComplet;
                ws.Cell(rr, 3).Value = "";
                ws.Cell(rr, 4).Value = prof.GradFunctie;
                ws.Cell(rr, 5).Value = 1;
                ws.Cell(rr, 6).Value = 0;
                ws.Cell(rr, 7).Value = "";
                ws.Cell(rr, 8).Value = prof.Facultate;
                ws.Cell(rr, 9).Value = prof.Departament;
                foreach (var kv in prof.Fractiuni) ws.Cell(rr, kv.Key).Value = kv.Value;
                ws.Cell(rr, 50).FormulaA1 = $"=SUM(J{rr}:AW{rr})";
                if (i % 2 != 0)
                    for (int c = 1; c <= 50; c++) ws.Cell(rr, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5f5f5");
            }

            int totalRow = dataStartRow + profesori.Count;
            ws.Cell(totalRow, 1).Value = "Total general:";
            ws.Cell(totalRow, 1).Style.Font.Bold = true;
            for (int c = 10; c <= 49; c++)
            {
                string col = ColumnLetter(c);
                ws.Cell(totalRow, c).FormulaA1 = $"=SUM({col}{dataStartRow}:{col}{totalRow - 1})";
                ws.Cell(totalRow, c).Style.Font.Bold = true;
            }
            ws.Cell(totalRow, 50).FormulaA1 = $"=SUM(J{totalRow}:AW{totalRow})";
            ws.Cell(totalRow, 50).Style.Font.Bold = true;

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            wb.Dispose();
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Date_ANS_{idAnUniv}.xlsx");
        }

        #endregion

        #region ================= HELPER METHODS =================

        // Returneaza clauza FROM+JOIN reutilizabila cu aliasurile standard sf/prof/au
        private static string SqlJoinFiltru() => @"
            FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
            LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
            LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv";

        private void StileazaHeader(IXLWorksheet ws, int headerRow, int nrColoane)
        {
            var range = ws.Range(headerRow, 1, headerRow, nrColoane);
            range.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
            range.Style.Font.FontColor = XLColor.White;
            range.Style.Font.Bold = true;
            var dataRange = ws.Range(headerRow, 1, headerRow + 1000, nrColoane);
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        private IActionResult ExcelResponse(XLWorkbook wb, string fileName)
        {
            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        private static string SanitizeFileName(string s) =>
            string.Join("_", s.Split(Path.GetInvalidFileNameChars()));

        private static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }
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
            ws.Cell(4, 10).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta.";
            ws.Range(4, 10, 4, 27).Merge();
            ws.Cell(4, 28).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta.";
            ws.Range(4, 28, 4, 50).Merge();

            ws.Cell(5, 1).Value = "Nr. \nCrt."; ws.Cell(5, 2).Value = "Nume si prenume cadru didactic";
            ws.Cell(5, 3).Value = "CNP"; ws.Cell(5, 4).Value = "Functie cadru didactic sau cercetare";
            ws.Cell(5, 5).Value = "Forma de angajare"; ws.Cell(5, 6).Value = "Calitate conducator doctorat";
            ws.Cell(5, 7).Value = "Varsta"; ws.Cell(5, 8).Value = "Facultate"; ws.Cell(5, 9).Value = "Departament";
            ws.Cell(5, 10).Value = "Matematica si stiinte ale naturii"; ws.Cell(5, 15).Value = "Stiinte ingineresti";
            ws.Cell(5, 22).Value = "Stiinte biologice si biomedicale"; ws.Cell(5, 28).Value = "Stiinte sociale";
            ws.Cell(5, 37).Value = "Stiinte umaniste si arte"; ws.Cell(5, 50).Value = "Total";

            foreach (var (r1, c1, r2, c2) in new (int, int, int, int)[] {
                (5,1,7,1),(5,2,7,2),(5,3,7,3),(5,4,7,4),(5,5,7,5),(5,6,7,6),(5,7,7,7),(5,8,7,8),(5,9,7,9),
                (5,10,5,14),(5,15,5,21),(5,22,5,27),(5,28,5,36),(5,37,5,49),(5,50,7,50) })
                ws.Range(r1, c1, r2, c2).Merge();

            string[] subdomenii = {
                "Matematica","Informatica","Fizica","Chimie si inginerie chimica","Stiintele pamantului si atmosferei",
                "Inginerie civila","Inginerie electrica, electronica si telecomunicatii",
                "Inginerie geologica, mine, petrol si gaze","Ingineria transporturilor",
                "Ingineria resurselor vegetale si animale",
                "Ingineria sistemelor, calculatoare si tehnologia informatiei",
                "Inginerie mecanica, mecatronica, inginerie industriala si management",
                "Biologie","Biochimie","Medicina","Medicina veterinara","Medicina dentara","Farmacie",
                "Stiinte juridice","Stiinte administrative","Stiinte ale comunicarii","Sociologie",
                "Stiinte politice","Stiinte militare, informatii si ordine publica",
                "Stiinte economice (doar Cibernetica, statistica si informatica economica)",
                "Stiinte economice (fara Cibernetica, statistica si informatica economica)",
                "Psihologie si stiinte comportamentale",
                "Filologie","Filosofie","Istorie","Teologie","Studii culturale",
                "Arhitectura si urbanism","Arte vizuale (fara Istoria si teoria artei)",
                "Arte vizuale (doar Istoria si teoria artei)","Teatru si artele spectacolului",
                "Cinematografie si media","Muzica (doar Interpretare muzicala)",
                "Muzica (fara Interpretare muzicala)","Stiintele Sportului si Educatiei Fizice"
            };
            for (int i = 0; i < subdomenii.Length; i++) { ws.Cell(6, 10 + i).Value = subdomenii[i]; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }

            for (int i = 0; i < 9; i++) ws.Cell(8, i + 1).Value = ((char)('A' + i)).ToString();
            for (int i = 0; i < 41; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, 50).Value = "40";

            var headerFill = XLColor.FromHtml(BrandColorHex);
            for (int row = 5; row <= 8; row++)
                for (int c = 1; c <= 50; c++)
                {
                    ws.Cell(row, c).Style.Font.Bold = true;
                    ws.Cell(row, c).Style.Fill.BackgroundColor = XLColor.White;
                    ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(row, c).Style.Alignment.WrapText = true;
                    ws.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(row, c).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
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
            if (g.Contains("conferentiar")) return "Conferentiar";
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