using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using ClosedXML.Excel;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RapoarteController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;
        private readonly IMemoryCache _cache;
        private const string BrandColor = "#56723e";

        // ============================================================
        // SURSA UNICA:
        //   sf   = [AGSIS].[pi].[StatDeFunctiiPeSpecializare]
        //   prof = [agsis_dw].[dbo].[View_ProfesoriActivi_CF]
        //   au   = [AGSIS].[dbo].[AnUniversitar]
        //
        // CUPLAJE: ROW_NUMBER() OVER(PARTITION BY ID_Post_Profesor_Materie)
        //   -> RnPost=1 pastreaza un singur rand per post.
        //   Mentiunile (specializarile cuplate) se extrag separat din toate
        //   randurile cu acelasi ID_Post_Profesor_Materie.
        //
        // FILTRARE CUPLAJE IN NORME:
        //   Daca utilizatorul filtreaza pe specializare X, iar un curs e cuplatat
        //   X+Y, cursul apare totusi (join pe ID_Post_Profesor_Materie).
        //   Asta inseamna ca filtrul pe specs se aplica la nivel de POST,
        //   nu la nivel de rand individual.
        //
        // ID_TipFormaInv: 1=IF, 2=IFR, 3=ID
        // prof.Titular=1 -> titular; prof.Titular=0 -> colaborator
        // ============================================================

        // CTE 1: deduplicare cuplaje (un rand per post)
        private const string CteDedup = @"
            SfDedup AS (
                SELECT *,
                    ROW_NUMBER() OVER(
                        PARTITION BY sf.ID_Post_Profesor_Materie
                        ORDER BY sf.ID_PlanMaterie
                    ) AS RnPost
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                WHERE sf.xTipCuplaj <> 'CuplajeCareNuMaiExista'
            )";

        // CTE 2: mentiuni cuplaj - toate specializarile pentru un post
        // Returneaza lista COMPLETA de specializari pt orice post cuplatat.
        // Astfel, la ambele/toate randurile cuplate apare aceeasi mentiune simetrica.
        // Ex: "Pachete software" cuplatat cu CIG+IE -> mentiunea = "CIG / IE" la AMBELE randuri.
        private const string CteMentiuni = @"
            MentiuniCuplaj AS (
                SELECT
                    sf2.ID_Post_Profesor_Materie,
                    STUFF((
                        SELECT DISTINCT ' / ' + LTRIM(RTRIM(sf3.DenumireSpecializare))
                        FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf3
                        WHERE sf3.ID_Post_Profesor_Materie = sf2.ID_Post_Profesor_Materie
                          AND sf3.xTipCuplaj <> 'CuplajeCareNuMaiExista'
                        FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'), 1, 3, '') AS ToateSpecializarile,
                    COUNT(DISTINCT sf2.DenumireSpecializare) AS NrSpec
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf2
                WHERE sf2.xTipCuplaj <> 'CuplajeCareNuMaiExista'
                GROUP BY sf2.ID_Post_Profesor_Materie
            )";

        // JOIN principal:
        //   prof = View_ProfesoriActivi_CF (titulari + colaboratori cu date complete)
        // NOTA: Post_Profesor_Materie NU este in SqlJoin global - contine intrari din toti anii
        //       si produce duplicate masive (Mizgaciu 41->144 ore, etc.)
        //       Este folosit DOAR in GetProfesori dropdown cu filtru strict pe id_anuniv.
        private const string SqlJoin = @"
            FROM SfDedup sf
            LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof
                   ON sf.ID_Profesor = prof.ID_Profesor
            LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au
                   ON sf.id_anuniv = au.ID_AnUniv";

        // WHERE standard (RnPost=1 + filtre)
        // NOTA: filtrul @specs se verifica si in cadrul postului (pt cuplaje)
        // adica daca o specializare din cuplaj e in @specs, postul e inclus
        // COLLATE Romanian_CI_AS adaugat explicit pe coloanele din AGSIS (cp1250)
        // pentru a evita conflicte cu agsis_dw (Romanian_CI_AS nvarchar)
        private const string SqlWhere = @"
            WHERE  sf.RnPost = 1
              AND (@an    = 'Toti' OR UPPER(LTRIM(RTRIM(CAST(au.Denumire AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
              AND (@fac   = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
              AND (@dept  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
              AND (@prof  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@prof))))
              AND (@sem   = 0      OR sf.NrSemestruDinAn = @sem)
              AND (@tip   = 'Toti' OR sf.DenTitularSauSuplinitor = @tip)
              AND (@formaId = 0    OR sf.ID_TipFormaInv = @formaId)
              AND (@specs = 'Toti' OR EXISTS (
                    SELECT 1 FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf_all
                    WHERE sf_all.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                      AND sf_all.xTipCuplaj <> 'CuplajeCareNuMaiExista'
                      AND UPPER(LTRIM(RTRIM(CAST(sf_all.DenumireSpecializare AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                          IN (SELECT UPPER(LTRIM(RTRIM(CAST(value AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) FROM STRING_SPLIT(@specs,','))))";

        private const string SqlOreCols = @"
                    ISNULL(sf.Nr_Ore_Curs, 0) AS OreCurs,
                    ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0)
                    + ISNULL(sf.Nr_Ore_Proiect, 0) + ISNULL(sf.Nr_Ore_SF, 0) AS OreAplic,
                    ISNULL(sf.NrOreConventionale, 0) AS OreConv";

        private static int FormaToId(string? f) => f?.ToUpper() switch
        { "IF" => 1, "IFR" => 2, "ID" => 3, _ => 0 };

        // Normalizeaza "Catedra de X" / "Departamentul X" -> "Departamentul X"
        // Valoarea din BD ramane neschimbata, doar afisarea e normalizata
        private static string NormalizeDept(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "Nespecificat";
            s = s.Trim();
            // Elimina prefixe cunoscute
            string[] prefixes = { "Catedra de ", "Catedra De ", "CATEDRA DE ",
                                  "Departamentul ", "DEPARTAMENTUL ", "Departament " };
            foreach (var p in prefixes)
                if (s.StartsWith(p, StringComparison.OrdinalIgnoreCase))
                    s = s.Substring(p.Length).Trim();
            return "Departamentul " + s;
        }

        private static readonly Dictionary<int, int> MappingMetaspec = new()
        {
            // Matematica (1), Informatica (2), Fizica (3)
            {20,5},{34,5},{44,5},{182,5},{837,5},{847,5},{848,5},
            {43,9},{148,8},{171,11},{306,11},{358,11},{464,11},{466,11},{615,11},{823,11},{828,11},
            {35,11},{82,7},{84,7},{139,11},{162,7},{310,7},{315,7},{316,7},{362,7},{372,7},
            {418,7},{597,7},{617,7},{116,12},{126,11},{129,11},{156,11},{228,11},{229,12},
            {404,11},{529,11},{142,11},{326,11},{531,11},{806,11},{819,11},{53,11},{138,11},
            {446,11},{448,11},{449,11},{450,11},{451,11},{496,11},{497,11},{821,11},
            {46,9},{122,9},{176,9},{178,9},{731,9},{72,9},{90,9},{118,9},{226,9},{235,9},
            {249,9},{307,9},{437,9},{458,9},{566,9},{845,9},
            {101,40},{102,40},{104,40},{251,1},{317,40},{340,1},{368,40},{369,40},{477,40},
            {45,25},{73,25},{93,25},{112,24},{221,25},{223,25},{227,25},{242,25},{283,25},
            {288,25},{299,25},{181,31},{196,27},{197,27},{200,27},{205,27},{207,27},{209,27},
            {217,27},{218,27},{341,27},{343,27},{463,27},{511,27},{512,27},{513,27},{514,27},
            {726,27},{798,27},{801,27},{60,18},{64,18},{331,18},{485,18},{555,18},
            {41,20},{98,20},{100,25},{322,21},{524,25},{579,20},{831,20},
            {276,26},{294,26},{296,26},{383,26},{384,26},{416,26},{515,26},{813,26},
            {851,26},{832,26},{834,26},{394,14},{397,14},{402,14},{484,14},{585,14},
            {594,14},{835,14},{186,37},{187,37},{264,37},{332,37},{351,37},{557,37},{838,37},
            {78,39},{189,39},{325,39},{470,39},{783,39},{784,39},{846,39},
            // ID-uri prezente in date dar nemapate anterior -> adaugate din Q7
            {297,11},{517,11},{859,40},{863,40},{864,40},{865,40},
            // metaspecializari suplimentare identificate
            {7,16},{9,18},{10,19},{13,22},{15,24},{16,25},{23,32},{28,37},
        };

        private static readonly Dictionary<int, int> AnsIdToCol = new()
        {
            {1,10},{2,11},{3,12},{4,13},{5,14},{6,15},{7,16},{8,17},{9,18},{10,19},
            {11,20},{12,21},{13,22},{14,23},{15,24},{16,25},{17,26},{18,27},{19,28},
            {20,29},{21,30},{22,31},{23,32},{24,33},{25,34},{26,35},{27,36},{28,37},
            {29,38},{30,39},{31,40},{32,41},{33,42},{34,43},{35,44},{36,45},{37,46},
            {38,47},{39,48},{40,49},
        };

        private readonly string[] DomeniiExcel = {
            "Matematica","Informatica","Fizica","Chimie si inginerie chimica",
            "Stiintele pamantului si atmosferei","Inginerie civila",
            "Inginerie electrica, electronica si telecomunicatii",
            "Inginerie geologica, mine, petrol si gaze","Ingineria transporturilor",
            "Ingineria resurselor vegetale si animale",
            "Ingineria sistemelor, calculatoare si tehnologia informatiei",
            "Inginerie mecanica, mecatronica, inginerie industriala si management",
            "Biologie","Biochimie","Medicina","Medicina veterinara","Medicina dentara","Farmacie",
            "Stiinte juridice","Stiinte administrative","Stiinte ale comunicarii","Sociologie",
            "Stiinte politice","Stiinte militare, informatii si ordine publica",
            "Stiinte economice (doar Cibernetica, statistica si informatica economica)",
            "Stiinte economice (fara Cibernetica, statistica si informatica economica)",
            "Psihologie si stiinte comportamentale","Filologie","Filosofie","Istorie",
            "Teologie","Studii culturale","Arhitectura si urbanism",
            "Arte vizuale (fara Istoria si teoria artei)","Arte vizuale (doar Istoria si teoria artei)",
            "Teatru si artele spectacolului","Cinematografie si media",
            "Muzica (doar Interpretare muzicala)","Muzica (fara Interpretare muzicala)",
            "Stiintele Sportului si Educatiei Fizice"
        };

        public RapoarteController(IConfiguration configuration, IMemoryCache cache)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        // ── helpers ───────────────────────────────────────────────────────────

        private void P(SqlCommand cmd, string? an, string? fac, string? dept,
                       string? prof, string? specs, int sem, string? tip, string? forma)
        {
            cmd.Parameters.AddWithValue("@an", an ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", fac ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", dept ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", prof ?? "Toti");
            cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specs) ? "Toti" : specs);
            cmd.Parameters.AddWithValue("@sem", sem);
            cmd.Parameters.AddWithValue("@tip", tip ?? "Toti");
            cmd.Parameters.AddWithValue("@formaId", FormaToId(forma));
        }

        private IActionResult Xlsx(XLWorkbook wb, string name)
        {
            using var ms = new MemoryStream();
            wb.SaveAs(ms); wb.Dispose();
            return File(ms.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", name);
        }

        private static void HdrStyle(IXLWorksheet ws, int row, int cols)
        {
            var r = ws.Range(row, 1, row, cols);
            r.Style.Fill.BackgroundColor = XLColor.FromHtml("#56723e");
            r.Style.Font.FontColor = XLColor.White;
            r.Style.Font.Bold = true;
            r.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            r.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        // Merge celule consecutive cu aceeasi valoare in coloana col (1-based)
        private static void MergeColumn(IXLWorksheet ws, int startRow, int endRow, int col)
        {
            int mergeStart = startRow;
            for (int i = startRow + 1; i <= endRow + 1; i++)
            {
                string cur = i <= endRow ? ws.Cell(i, col).GetString() : null!;
                string prev = ws.Cell(i - 1, col).GetString();
                if (cur != prev || i == endRow + 1)
                {
                    if (i - 1 > mergeStart)
                    {
                        ws.Range(mergeStart, col, i - 1, col).Merge();
                        ws.Cell(mergeStart, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }
                    mergeStart = i;
                }
            }
        }

        private static string Safe(string s) =>
            string.Join("_", s.Split(Path.GetInvalidFileNameChars())).Trim('_');

        private static string ColLetter(int col)
        {
            string res = "";
            while (col > 0) { col--; res = (char)('A' + col % 26) + res; col /= 26; }
            return res;
        }

        // ═════════════════════════════════════════════════════════════════════
        #region LISTE DROPDOWNS (cascada: fac -> dept -> spec -> prof)
        // ═════════════════════════════════════════════════════════════════════

        [HttpGet("liste/ani-universitari")]
        public IActionResult GetAni()
        {
            var list = new List<object>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT au.ID_AnUniv AS id, au.Denumire COLLATE Romanian_CI_AS AS nume
                FROM   [AGSIS].[dbo].[AnUniversitar] au
                WHERE  au.ID_AnUniv IN (
                       SELECT DISTINCT id_anuniv FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare])
                ORDER BY au.ID_AnUniv DESC", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(new { id = r["id"].ToString(), nume = r["nume"].ToString() });
            return Ok(list);
        }

        [HttpGet("liste/facultati")]
        public IActionResult GetFacultati(string? anUniv)
        {
            var list = new List<string>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand($@"
                WITH {CteDedup}
                SELECT DISTINCT UPPER(LTRIM(RTRIM(prof.DenumireFacultate COLLATE Romanian_CI_AS))) AS Fac
                {SqlJoin}
                WHERE  sf.RnPost = 1 AND prof.DenumireFacultate COLLATE Romanian_CI_AS IS NOT NULL
                  AND  (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
                ORDER BY Fac", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r[0]?.ToString(); if (!string.IsNullOrWhiteSpace(v)) list.Add(v); }
            return Ok(list);
        }

        [HttpGet("liste/departamente")]
        public IActionResult GetDepartamente(string? anUniv, string? numeFacultate)
        {
            var list = new List<string>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand($@"
                WITH {CteDedup}
                SELECT DISTINCT UPPER(LTRIM(RTRIM(prof.DenumireCatedra COLLATE Romanian_CI_AS))) AS Dept
                {SqlJoin}
                WHERE  sf.RnPost = 1 AND prof.DenumireCatedra COLLATE Romanian_CI_AS IS NOT NULL
                  AND  (@an  = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
                  AND  (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                ORDER BY Dept", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r[0]?.ToString(); if (!string.IsNullOrWhiteSpace(v)) list.Add(v); }
            return Ok(list);
        }

        // Specializari filtrate dupa AN + FACULTATE + DEPARTAMENT + CICLU (cascada completa)
        // Ciclu: 'Licenta' = NrAnStudii >= 3, 'Master' = NrAnStudii <= 2
        // Nu avem acces la view_metaspecializare, folosim NrAnStudii din sf
        [HttpGet("liste/specializari-per-facultate")]
        public IActionResult GetSpecializari(string? anUniv, string? numeFacultate,
            string? numeDepartament, string? ciclu)
        {
            var list = new List<string>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand($@"
                WITH {CteDedup},
                SpecCiclu AS (
                    SELECT DISTINCT
                        UPPER(LTRIM(RTRIM(sf.DenumireSpecializare COLLATE Romanian_CI_AS))) AS Spec,
                        MAX(sf.NrAnStudii) AS MaxAn
                    {SqlJoin}
                    WHERE  sf.RnPost = 1 AND sf.DenumireSpecializare COLLATE Romanian_CI_AS IS NOT NULL
                      AND  (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
                      AND  (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                      AND  (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                    GROUP BY UPPER(LTRIM(RTRIM(sf.DenumireSpecializare COLLATE Romanian_CI_AS)))
                )
                SELECT Spec FROM SpecCiclu
                WHERE (@ciclu = 'Toti'
                    OR (@ciclu = 'Licenta' AND MaxAn >= 3)
                    OR (@ciclu = 'Master'  AND MaxAn <= 2))
                ORDER BY Spec", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", numeDepartament ?? "Toti");
            cmd.Parameters.AddWithValue("@ciclu", ciclu ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var v = r[0]?.ToString();
                if (!string.IsNullOrWhiteSpace(v) && !list.Contains(v)) list.Add(v);
            }
            return Ok(list);
        }

        // Profesori filtrati dupa AN + FACULTATE + DEPARTAMENT + SPECIALIZARI
        // SURSA NUME: agsis_dw.dbo.Post_Profesor_Materie care are NumeIntreg pentru TOTI
        // (inclusiv suplinitorii care lipsesc din View_ProfesoriActivi_CF)
        [HttpGet("liste/profesori-per-specializari")]
        public IActionResult GetProfesori(string? anUniv, string? facultate,
                                          string? specializari, string? departament)
        {
            var list = new List<string>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(@"
                WITH SfFilt AS (
                    SELECT DISTINCT sf.ID_Profesor, sf.ID_Post_Profesor_Materie,
                           CAST(sf.DenumireSpecializare AS NVARCHAR(500)) COLLATE Romanian_CI_AS AS DenumireSpecializare,
                           sf.xTipCuplaj
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                    WHERE sf.xTipCuplaj <> 'CuplajeCareNuMaiExista'
                      AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(CAST(au.Denumire AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(CAST(@an AS NVARCHAR(500)) COLLATE Romanian_CI_AS))))
                )
                SELECT DISTINCT CAST(ppm.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS AS Profesor
                FROM SfFilt sf
                INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Profesor = ppm.ID_Profesor
                LEFT  JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor = prof.ID_Profesor
                WHERE ppm.NumeIntreg IS NOT NULL
                  AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(
                        COALESCE(
                            CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                            CAST(ppm.DenumireFacultate  AS NVARCHAR(500)) COLLATE Romanian_CI_AS
                        ), '')))) = UPPER(LTRIM(RTRIM(CAST(@fac AS NVARCHAR(500)) COLLATE Romanian_CI_AS))))
                  AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(
                        COALESCE(
                            CAST(prof.DenumireCatedra        AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                            CAST(ppm.DenumireCatedraProfesor AS NVARCHAR(500)) COLLATE Romanian_CI_AS
                        ), '')))) = UPPER(LTRIM(RTRIM(CAST(@dept AS NVARCHAR(500)) COLLATE Romanian_CI_AS))))
                  AND (@specs = 'Toti' OR EXISTS (
                        SELECT 1 FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf_all
                        WHERE sf_all.ID_Post_Profesor_Materie = sf.ID_Post_Profesor_Materie
                          AND sf_all.xTipCuplaj <> 'CuplajeCareNuMaiExista'
                          AND UPPER(LTRIM(RTRIM(CAST(sf_all.DenumireSpecializare AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                              IN (SELECT UPPER(LTRIM(RTRIM(CAST(value AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) FROM STRING_SPLIT(@specs,','))))
                ORDER BY Profesor", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r[0]?.ToString(); if (!string.IsNullOrWhiteSpace(v)) list.Add(v); }
            return Ok(list);
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 1 – NORMA PROFESORI (cu coloana Mentiuni cuplaje)
        // ═════════════════════════════════════════════════════════════════════

        // RAPORT 1: Norme - afiseaza TOATE randurile cuplate (nu dedupliact pe post)
        // dar exclude copiile AplicDin* (care sunt generate automat de sistem).
        // Astfel: "Pachete software" cuplatat CIG+IE -> apare de 2 ori (un rand per specializare)
        // cu mentiunea "CONTABILITATE... / INFORMATICĂ ECONOMICĂ" la AMBELE randuri.
        // Deduplicarea RnPost=1 se face DOAR pentru calculele de ore (totaluri), nu pentru afisaj.
        private string NormaQ => $@"
            WITH {CteDedup},
                 {CteMentiuni},
            SfNorma AS (
                SELECT *
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf_inner
                WHERE sf_inner.xTipCuplaj NOT IN ('AplicDinCuplajCurs','AplicDinCuplajApp','CuplajeCareNuMaiExista')
            )
            SELECT
                ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                       'Prof.ID ' + CAST(sfn.ID_Profesor AS VARCHAR))  AS Profesor,
                CAST(sfn.DenumireSpecializare AS NVARCHAR(500))
                    COLLATE Romanian_CI_AS                              AS DenumireSpecializare,
                LTRIM(RTRIM(ISNULL(sfn.DenumireMaterie, '')))          AS Materie,
                sfn.DenTitularSauSuplinitor                            AS TipPost,
                sfn.NrSemestruDinAn                                    AS Semestru,
                CASE sfn.ID_TipFormaInv
                    WHEN 1 THEN 'IF' WHEN 2 THEN 'IFR'
                    WHEN 3 THEN 'ID' ELSE '' END                       AS Forma,
                ISNULL(sfn.Nr_Ore_Curs, 0)                            AS OreCurs,
                ISNULL(sfn.Nr_Ore_Seminar, 0)
                    + ISNULL(sfn.Nr_Ore_Laborator, 0)
                    + ISNULL(sfn.Nr_Ore_Proiect, 0)
                    + ISNULL(sfn.Nr_Ore_SF, 0)                         AS OreAplic,
                ISNULL(sfn.NrOreConventionale, 0)                     AS OreConv,
                CASE WHEN mc.NrSpec > 1 THEN mc.ToateSpecializarile
                     ELSE NULL END                                     AS Mentiuni
            FROM SfNorma sfn
            LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof
                   ON sfn.ID_Profesor = prof.ID_Profesor
            LEFT JOIN [AGSIS].[dbo].[AnUniversitar] au
                   ON sfn.id_anuniv = au.ID_AnUniv
            LEFT JOIN MentiuniCuplaj mc
                   ON sfn.ID_Post_Profesor_Materie = mc.ID_Post_Profesor_Materie
            WHERE  (@an    = 'Toti' OR UPPER(LTRIM(RTRIM(CAST(au.Denumire AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
              AND  (@fac   = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
              AND  (@dept  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
              AND  (@prof  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@prof))))
              AND  (@sem   = 0      OR sfn.NrSemestruDinAn = @sem)
              AND  (@tip   = 'Toti' OR sfn.DenTitularSauSuplinitor = @tip)
              AND  (@formaId = 0    OR sfn.ID_TipFormaInv = @formaId)
              AND  (@specs = 'Toti' OR EXISTS (
                    SELECT 1 FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf_all
                    WHERE sf_all.ID_Post_Profesor_Materie = sfn.ID_Post_Profesor_Materie
                      AND sf_all.xTipCuplaj <> 'CuplajeCareNuMaiExista'
                      AND UPPER(LTRIM(RTRIM(CAST(sf_all.DenumireSpecializare AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                          IN (SELECT UPPER(LTRIM(RTRIM(CAST(value AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                              FROM STRING_SPLIT(@specs,','))))
            ORDER BY Profesor,
                     CAST(sfn.DenumireSpecializare AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                     Materie";

        [HttpGet("norma-profesori")]
        public IActionResult GetNorma(string? anUniv, string? facultate, string? departament,
            string? specializari, string? profesor, int semestru = 0,
            string? tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var res = new List<object>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(NormaQ, conn);
            P(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                res.Add(new
                {
                    Profesor = r["Profesor"].ToString(),
                    Specializare = r["DenumireSpecializare"].ToString(),
                    Materie = r["Materie"].ToString(),
                    TipPost = r["TipPost"].ToString(),
                    Semestru = Convert.ToInt32(r["Semestru"]),
                    FormaInvatamant = r["Forma"].ToString(),
                    NrOreCurs = Convert.ToDouble(r["OreCurs"]),
                    NrOreAplicatii = Convert.ToDouble(r["OreAplic"]),
                    NrOreConventionale = Convert.ToDouble(r["OreConv"]),
                    Mentiuni = r["Mentiuni"] == DBNull.Value ? null : r["Mentiuni"].ToString()
                });
            return Ok(res);
        }

        [HttpGet("export/norma")]
        public IActionResult ExportNorma(string? anUniv, string? facultate, string? departament,
            string? specializari, string? profesor, int semestru = 0,
            string? tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            // Coloane conforme cu formatul corect (fara Forma Inv, fara Mentiuni cuplaj)
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"),
                new DataColumn("Specializare"),
                new DataColumn("Materie"),
                new DataColumn("Tip Post"),
                new DataColumn("Semestru"),
                new DataColumn("Nr Ore Curs",           typeof(double)),
                new DataColumn("Nr Ore Aplicatii",      typeof(double)),
                new DataColumn("Nr Ore Conventionale",  typeof(double))
            });
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(NormaQ, conn);
            P(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(
                    r["Profesor"].ToString(),
                    r["DenumireSpecializare"].ToString(),
                    r["Materie"].ToString(),
                    r["TipPost"].ToString(),
                    r["Semestru"].ToString(),
                    Convert.ToDouble(r["OreCurs"]),
                    Convert.ToDouble(r["OreAplic"]),
                    Convert.ToDouble(r["OreConv"]));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme");

            // 3 randuri header (exact ca formatul corect)
            string semStr = semestru == 0 ? "Toate" : semestru.ToString();
            string tipStr = string.IsNullOrEmpty(tipPost) ? "Toti" : tipPost;
            string formaStr = string.IsNullOrEmpty(formaInvatamant) ? "Toti" : formaInvatamant;

            ws.Cell(1, 1).Value = "Filtre Aplicate";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);

            ws.Cell(2, 1).Value = $"An Universitar: {anUniv?.ToUpper() ?? "Toti"} | " +
                                  $"Facultate: {facultate?.ToUpper() ?? "Toti"} | " +
                                  $"Departament: {departament?.ToUpper() ?? "Toti"}";
            ws.Cell(2, 1).Style.Font.Italic = true;

            ws.Cell(3, 1).Value = $"Profesor: {profesor?.ToUpper() ?? "Toti"} | " +
                                  $"Semestru: {semStr} | " +
                                  $"Tip Post: {tipStr} | " +
                                  $"Forma Inv: {formaStr}";
            ws.Cell(3, 1).Style.Font.Italic = true;

            // Merge header pe toata latimea
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            ws.Range(2, 1, 2, dt.Columns.Count).Merge();
            ws.Range(3, 1, 3, dt.Columns.Count).Merge();

            // Tabel de date incepe la randul 5 (rand 4 = blank)
            int dataRow = 5;
            var tbl = ws.Cell(dataRow, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None;
            HdrStyle(ws, dataRow, dt.Columns.Count);
            tbl.ShowTotalsRow = true;
            tbl.Field("Nr Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Aplicatii").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Conventionale").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";

            ws.Columns().AdjustToContents();
            string fn = (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                ? $"NormaProfesori_{Safe(profesor)}.xlsx" : "NormaProfesori_General.xlsx";
            return Xlsx(wb, fn);
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 2 – DISTRIBUTIE ORE PE PROGRAM
        // ═════════════════════════════════════════════════════════════════════

        private string OreProgramQ => $@"
            WITH {CteDedup},
            Filtrat AS (
                SELECT
                    ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Prof.ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                    UPPER(LTRIM(RTRIM(sf.DenumireSpecializare COLLATE Romanian_CI_AS))) AS Program,
                    sf.DenTitularSauSuplinitor AS TipPost,
                    ISNULL(sf.NrOreConventionale,0) AS OreConv
                {SqlJoin} {SqlWhere}
            ),
            Grupat AS (
                SELECT Profesor, Program, TipPost, SUM(OreConv) AS OreProgram
                FROM   Filtrat GROUP BY Profesor, Program, TipPost
            ),
            TotalPost AS (
                SELECT Profesor, TipPost, SUM(OreProgram) AS TotPost
                FROM   Grupat GROUP BY Profesor, TipPost
            )
            SELECT g.Profesor, g.Program AS ProgramStudiu, g.TipPost,
                   g.OreProgram AS NrOreConv, t.TotPost AS TotalPost,
                   CAST(CASE WHEN t.TotPost=0 THEN 0
                        ELSE (g.OreProgram/t.TotPost)*100 END AS DECIMAL(10,2)) AS ProcentPost
            FROM   Grupat g
            INNER  JOIN TotalPost t ON g.Profesor=t.Profesor AND g.TipPost=t.TipPost
            ORDER BY g.Profesor, g.TipPost DESC, g.OreProgram DESC";

        [HttpGet("ore-profesor-program")]
        public IActionResult GetOreProgram(string? anUniv, string? facultate, string? departament,
            string? specializari, string? profesor, int semestru = 0,
            string? tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var res = new List<object>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(OreProgramQ, conn);
            P(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                res.Add(new
                {
                    Profesor = r["Profesor"].ToString(),
                    ProgramStudiu = r["ProgramStudiu"].ToString(),
                    TipPost = r["TipPost"].ToString(),
                    NrOreConv = Convert.ToDouble(r["NrOreConv"]),
                    TotalPost = Convert.ToDouble(r["TotalPost"]),
                    ProcentPost = Convert.ToDouble(r["ProcentPost"])
                });
            return Ok(res);
        }

        [HttpGet("export/ore-program")]
        public IActionResult ExportOreProgram(string? anUniv, string? facultate, string? departament,
            string? specializari, string? profesor, int semestru = 0,
            string? tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            // Format corect: 5 coloane fara Tip Post, cu 3 randuri header
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"),
                new DataColumn("Program Studiu"),
                new DataColumn("Nr Ore Conv",    typeof(double)),
                new DataColumn("Procent Post",   typeof(double)),
                new DataColumn("Total Post",     typeof(double))
            });
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(OreProgramQ, conn);
            P(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(
                    r["Profesor"].ToString(),
                    r["ProgramStudiu"].ToString(),
                    Convert.ToDouble(r["NrOreConv"]),
                    Convert.ToDouble(r["ProcentPost"]),
                    Convert.ToDouble(r["TotalPost"]));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Distributie Ore");

            string semStr = semestru == 0 ? "Toate" : semestru.ToString();
            string tipStr = string.IsNullOrEmpty(tipPost) ? "Toti" : tipPost;
            ws.Cell(1, 1).Value = "Filtre Aplicate";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
            ws.Cell(2, 1).Value = $"An Universitar: {anUniv?.ToUpper() ?? "Toti"} | " +
                                  $"Facultate: {facultate?.ToUpper() ?? "Toti"} | " +
                                  $"Departament: {departament?.ToUpper() ?? "Toti"}";
            ws.Cell(2, 1).Style.Font.Italic = true;
            ws.Cell(3, 1).Value = $"Profesor: {profesor?.ToUpper() ?? "Toti"} | " +
                                  $"Semestru: {semStr} | Tip Post: {tipStr}";
            ws.Cell(3, 1).Style.Font.Italic = true;
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            ws.Range(2, 1, 2, dt.Columns.Count).Merge();
            ws.Range(3, 1, 3, dt.Columns.Count).Merge();

            int dataRow = 5;
            var tbl = ws.Cell(dataRow, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            HdrStyle(ws, dataRow, dt.Columns.Count);
            if (dt.Rows.Count > 1) MergeColumn(ws, dataRow + 1, dataRow + dt.Rows.Count - 1, 1);
            ws.Columns().AdjustToContents();
            string fn = (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                ? $"StatisticaOre_{Safe(profesor)}.xlsx" : "StatisticaOre_General.xlsx";
            return Xlsx(wb, fn);
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 3 – TOTALURI NORME
        // ═════════════════════════════════════════════════════════════════════

        private string TotaluriQ => $@"
            WITH {CteDedup},
            Baza AS (
                SELECT
                    ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Prof.ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                    ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Departament,
                    ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Facultate,
                    ISNULL(sf.NrOreConventionale,0) AS OreConv
                {SqlJoin}
                WHERE  sf.RnPost = 1
                  AND  (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(CAST(au.Denumire AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))                        = UPPER(LTRIM(RTRIM(@an))))
                  AND  (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND  (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                  AND  (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.NumeIntreg        AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@prof))))
            )
            SELECT Profesor, Departament, Facultate,
                   CAST(SUM(OreConv)      AS DECIMAL(10,2)) AS TotalSapt,
                   CAST(SUM(OreConv) * 14 AS DECIMAL(10,2)) AS TotalAn
            FROM   Baza
            GROUP BY Profesor, Departament, Facultate
            ORDER BY Profesor";

        [HttpGet("norma-totaluri")]
        public IActionResult GetNormaTotaluri(string? anUniv, string? facultate,
            string? departament, string? profesor)
        {
            var res = new List<object>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(TotaluriQ, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                res.Add(new
                {
                    Profesor = r["Profesor"].ToString(),
                    Departament = NormalizeDept(r["Departament"].ToString()),
                    Facultate = r["Facultate"].ToString(),
                    TotalOreConv = Convert.ToDecimal(r["TotalSapt"]),
                    TotalAnualOreConv = Convert.ToDecimal(r["TotalAn"])
                });
            return Ok(res);
        }

        [HttpGet("export/norma-totaluri")]
        public IActionResult ExportNormaTotaluri(string? anUniv, string? facultate,
            string? departament, string? profesor)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume Profesor"),
                new DataColumn("Departament"),
                new DataColumn("Facultate"),
                new DataColumn("Total Ore Conv.",          typeof(decimal)),
                new DataColumn("Total tot anul ore conv.", typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(TotaluriQ, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(
                    r["Profesor"].ToString(),
                    NormalizeDept(r["Departament"].ToString()),
                    r["Facultate"].ToString(),
                    Convert.ToDecimal(r["TotalSapt"]),
                    Convert.ToDecimal(r["TotalAn"]));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Totaluri Norme");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            HdrStyle(ws, 1, dt.Columns.Count);
            tbl.ShowTotalsRow = true;
            tbl.Field("Total Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total tot anul ore conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nume Profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            return Xlsx(wb, "Totaluri_Norme.xlsx");
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 4 – LIMBI STRAINE
        // ═════════════════════════════════════════════════════════════════════

        private string LimbiQ => $@"
            WITH {CteDedup},
            Baza AS (
                SELECT
                    ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Prof.ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                    sf.NrSemestruDinAn AS Sem,
                    ISNULL(sf.NrOreConventionale,0) AS OreConv,
                    au.Denumire COLLATE Romanian_CI_AS AS AnUniv
                {SqlJoin}
                WHERE  sf.RnPost = 1
                  AND  (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire COLLATE Romanian_CI_AS)))                        = UPPER(LTRIM(RTRIM(@an))))
                  AND  (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireFacultate COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND  (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(prof.DenumireCatedra COLLATE Romanian_CI_AS,''))))   = UPPER(LTRIM(RTRIM(@dept))))
                  AND  (@prof = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS,''))))        = UPPER(LTRIM(RTRIM(@prof))))
                  AND  (@sem  = 0      OR sf.NrSemestruDinAn = @sem)
                  AND  (@tip  = 'Toti' OR sf.DenTitularSauSuplinitor = @tip)
                  AND  (
                         sf.DenumireSpecializare COLLATE Romanian_CI_AS LIKE '%englez%'
                      OR sf.DenumireSpecializare COLLATE Romanian_CI_AS LIKE '%francez%'
                      OR sf.DenumireSpecializare COLLATE Romanian_CI_AS LIKE '%german%'
                      OR sf.DenumireSpecializare COLLATE Romanian_CI_AS LIKE '%american%'
                      OR sf.DenumireSpecializare COLLATE Romanian_CI_AS LIKE '%(EN)%'
                      OR sf.DenumireSpecializare COLLATE Romanian_CI_AS LIKE '%(FR)%'
                      OR sf.DenumireSpecializare COLLATE Romanian_CI_AS LIKE '%(G)%'
                  )
            )
            SELECT Profesor,
                   CAST(SUM(CASE WHEN Sem % 2 = 1 THEN OreConv ELSE 0 END)*14 AS DECIMAL(10,2)) AS Sem1,
                   CAST(SUM(CASE WHEN Sem % 2 = 0 THEN OreConv ELSE 0 END)*14 AS DECIMAL(10,2)) AS Sem2,
                   CAST(SUM(OreConv)*14 AS DECIMAL(10,2)) AS Total,
                   MIN(AnUniv) AS AnUniv
            FROM   Baza
            GROUP BY Profesor
            HAVING SUM(OreConv) > 0
            ORDER BY Profesor";

        private void PLimbi(SqlCommand cmd, string? an, string? fac, string? dept,
                            string? prof, int sem, string? tip)
        {
            cmd.Parameters.AddWithValue("@an", an ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", fac ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", dept ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", prof ?? "Toti");
            cmd.Parameters.AddWithValue("@sem", sem);
            cmd.Parameters.AddWithValue("@tip", tip ?? "Toti");
        }

        [HttpGet("limbi-straine")]
        public IActionResult GetLimbiStraine(string? anUniv, string? facultate, string? departament,
            string? profesor, int semestru = 0, string? tipPost = "Toti")
        {
            var res = new List<object>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(LimbiQ, conn);
            PLimbi(cmd, anUniv, facultate, departament, profesor, semestru, tipPost);
            using var r = cmd.ExecuteReader();
            int nr = 1;
            while (r.Read())
                res.Add(new
                {
                    NrCrt = nr++,
                    NumeProfesor = r["Profesor"].ToString(),
                    AnUniv = r["AnUniv"].ToString(),
                    Sem1 = Convert.ToDecimal(r["Sem1"]),
                    Sem2 = Convert.ToDecimal(r["Sem2"]),
                    Total = Convert.ToDecimal(r["Total"])
                });
            return Ok(res);
        }

        [HttpGet("export/limbi-straine")]
        public IActionResult ExportLimbiStraine(string? anUniv, string? facultate, string? departament,
            string? profesor, int semestru = 0, string? tipPost = "Toti")
        {
            // Format corect: 5 coloane, fara An Universitar, cu 1 rand header "An universitar: ..."
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nr. Crt.",       typeof(int)),
                new DataColumn("Nume si prenume profesor"),
                new DataColumn("Total Sem 1",    typeof(decimal)),
                new DataColumn("Total Sem 2",    typeof(decimal)),
                new DataColumn("Total",          typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(LimbiQ, conn);
            PLimbi(cmd, anUniv, facultate, departament, profesor, semestru, tipPost);
            using var r = cmd.ExecuteReader();
            int nr = 1;
            while (r.Read())
                dt.Rows.Add(
                    nr++,
                    r["Profesor"].ToString(),
                    Convert.ToDecimal(r["Sem1"]),
                    Convert.ToDecimal(r["Sem2"]),
                    Convert.ToDecimal(r["Total"]));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Limbi Straine");

            // Header: un rand cu anul universitar
            string anDisplay = string.IsNullOrEmpty(anUniv) || anUniv == "Toti"
                ? "Toti" : anUniv.ToUpper();
            ws.Cell(1, 1).Value = $"An universitar: {anDisplay}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();

            // Tabel la randul 3
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            HdrStyle(ws, 3, dt.Columns.Count);
            tbl.ShowTotalsRow = true;
            tbl.Field("Total Sem 1").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total Sem 2").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr. Crt.").TotalsRowLabel = "TOTAL";
            ws.Columns().AdjustToContents();
            return Xlsx(wb, "Raport_Limbi_Straine.xlsx");
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 5 – DISCIPLINE PREDATE
        // FIX: STRING_AGG(DISTINCT) nu merge in SQL Server -> subquery STUFF/FOR XML
        // Export: 3 endpoint-uri separate IF / IFR / ID + unul combinat
        // ═════════════════════════════════════════════════════════════════════

        // Disciplinele per (profesor, forma) folosind STUFF + FOR XML in loc de STRING_AGG(DISTINCT)
        private string BuildDisciplineQ(int? formaId = null) => $@"
            WITH {CteDedup},
            -- Pre-agrega disciplinele o singura data (evita subquery corelat lent)
            DiscAgregate AS (
                SELECT
                    sf2.ID_Profesor,
                    sf2.id_anuniv,
                    sf2.ID_TipFormaInv,
                    STUFF((
                        SELECT DISTINCT ' | ' + LTRIM(RTRIM(sf3.DenumireMaterie))
                        FROM SfDedup sf3
                        WHERE sf3.ID_Profesor      = sf2.ID_Profesor
                          AND sf3.id_anuniv        = sf2.id_anuniv
                          AND sf3.ID_TipFormaInv   = sf2.ID_TipFormaInv
                          AND sf3.RnPost           = 1
                          AND sf3.DenumireMaterie  IS NOT NULL
                        FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'), 1, 3, '') AS Discipline
                FROM SfDedup sf2
                WHERE sf2.RnPost = 1
                GROUP BY sf2.ID_Profesor, sf2.id_anuniv, sf2.ID_TipFormaInv
            )
            SELECT
                ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                       'Prof.ID ' + CAST(sf.ID_Profesor AS VARCHAR))   AS Profesor,
                ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Departament,
                ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Facultate,
                CASE sf.ID_TipFormaInv
                    WHEN 1 THEN 'IF' WHEN 2 THEN 'IFR'
                    WHEN 3 THEN 'ID' ELSE 'Alte' END                   AS Forma,
                da.Discipline
            {SqlJoin}
            INNER JOIN DiscAgregate da
                    ON da.ID_Profesor    = sf.ID_Profesor
                   AND da.id_anuniv     = sf.id_anuniv
                   AND da.ID_TipFormaInv = sf.ID_TipFormaInv
            {SqlWhere}
            {(formaId.HasValue ? $"AND sf.ID_TipFormaInv = {formaId}" : "")}
            GROUP BY sf.ID_Profesor, sf.ID_TipFormaInv, sf.id_anuniv,
                     CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                     CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                     CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,
                     da.Discipline
            ORDER BY Forma, Profesor";

        [HttpGet("discipline-predate")]
        public IActionResult GetDiscipline(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string? tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var res = new List<object>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(BuildDisciplineQ(), conn);
            P(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, formaInvatamant);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                res.Add(new
                {
                    Profesor = r["Profesor"].ToString(),
                    Departament = NormalizeDept(r["Departament"].ToString()),
                    Facultate = r["Facultate"].ToString(),
                    Forma = r["Forma"].ToString(),
                    Discipline = r["Discipline"]?.ToString() ?? ""
                });
            return Ok(res);
        }

        // Export combinat (toate formele, sheet-uri separate)
        [HttpGet("export/discipline-predate")]
        public IActionResult ExportDiscipline(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string? tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            return ExportDisciplineIntern(anUniv, facultate, departament, profesor,
                specializari, semestru, tipPost, null, "Discipline_Predate_Toate.xlsx");
        }

        // Export doar IF
        [HttpGet("export/discipline-if")]
        public IActionResult ExportDisciplineIF(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0, string? tipPost = "Toti")
        {
            return ExportDisciplineIntern(anUniv, facultate, departament, profesor,
                specializari, semestru, tipPost, 1, "Discipline_Predate_IF.xlsx");
        }

        // Export doar IFR
        [HttpGet("export/discipline-ifr")]
        public IActionResult ExportDisciplineIFR(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0, string? tipPost = "Toti")
        {
            return ExportDisciplineIntern(anUniv, facultate, departament, profesor,
                specializari, semestru, tipPost, 2, "Discipline_Predate_IFR.xlsx");
        }

        // Export doar ID
        [HttpGet("export/discipline-id")]
        public IActionResult ExportDisciplineID(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0, string? tipPost = "Toti")
        {
            return ExportDisciplineIntern(anUniv, facultate, departament, profesor,
                specializari, semestru, tipPost, 3, "Discipline_Predate_ID.xlsx");
        }

        private IActionResult ExportDisciplineIntern(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru, string? tipPost,
            int? formaFilter, string fileName)
        {
            using var conn = new SqlConnection(_connectionString); conn.Open();
            string sql = BuildDisciplineQ(formaFilter);
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120; // 2 minute - query complex cu agregari
            P(cmd, anUniv, facultate, departament, profesor, specializari, semestru, tipPost, "Toti");
            var rows = new List<(string Profesor, string Dept, string Fac, string Forma, string Disc)>();
            using var r = cmd.ExecuteReader();
            while (r.Read())
                rows.Add((r["Profesor"].ToString()!, NormalizeDept(r["Departament"].ToString()),
                          r["Facultate"].ToString()!, r["Forma"].ToString()!,
                          r["Discipline"]?.ToString() ?? ""));

            using var wb = new XLWorkbook();
            var formeToWrite = formaFilter.HasValue
                ? new[] { formaFilter == 1 ? "IF" : formaFilter == 2 ? "IFR" : "ID" }
                : new[] { "IF", "IFR", "ID" };

            foreach (var forma in formeToWrite)
            {
                var sub = rows.Where(x => x.Forma == forma).ToList();
                if (!sub.Any()) continue;
                var ws = wb.Worksheets.Add(forma);
                ws.Cell(1, 1).Value = "Forma"; ws.Cell(1, 2).Value = "Profesor";
                ws.Cell(1, 3).Value = "Departament"; ws.Cell(1, 4).Value = "Facultate";
                ws.Cell(1, 5).Value = "Discipline predate";
                HdrStyle(ws, 1, 5);
                for (int i = 0; i < sub.Count; i++)
                {
                    ws.Cell(i + 2, 1).Value = forma;
                    ws.Cell(i + 2, 2).Value = sub[i].Profesor;
                    ws.Cell(i + 2, 3).Value = sub[i].Dept;
                    ws.Cell(i + 2, 4).Value = sub[i].Fac;
                    ws.Cell(i + 2, 5).Value = sub[i].Disc;
                    ws.Cell(i + 2, 5).Style.Alignment.WrapText = false;
                    if (i % 2 != 0)
                        for (int c = 1; c <= 5; c++)
                            ws.Cell(i + 2, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#f9f9f9");
                }
                ws.Column(1).Width = 6; ws.Column(2).Width = 32;
                ws.Column(3).Width = 38; ws.Column(4).Width = 38; ws.Column(5).Width = 90;
            }
            if (!wb.Worksheets.Any()) wb.Worksheets.Add("Fara date");
            return Xlsx(wb, fileName);
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 6 – TITULARI
        // ═════════════════════════════════════════════════════════════════════
        // LISTA TITULARI HARDCODATA (furnizata de prof. Ionita, 740 persoane)
        // Aceasta lista este sursa de adevar pentru raportul Titulari SI pentru ANS.
        // Matching: normalizare upper + collapse whitespace + fuzzy pe numele BD.
        // ═════════════════════════════════════════════════════════════════════

        private static readonly HashSet<string> TitulariHardcoded = new(StringComparer.OrdinalIgnoreCase)
        {
            "ABAITANCEI HORIA",
            "ABRUDAN IOAN VASILE",
            "ACIU LIA ELENA",
            "ADAM MIHAI-SORIN",
            "ADOCHIȚE (GĂLBĂU) CRISTINA-ȘTEFANIA",
            "AGACHE IOANA-OCTAVIA",
            "ALBU RUXANDRA GABRIELA",
            "ALDEA ADRIAN",
            "ALDEA CODRUTA NICOLETA",
            "ALDEA CONSTANTIN LUCIAN",
            "ALECU STEFAN",
            "ALEXANDRESCU DANA SORINA",
            "ALEXANDRU CATALIN",
            "ALEXANDRU MARIAN",
            "ALEXE RALUCA MONICA",
            "ANASTASIU ALEXANDRU-RAZVAN",
            "ANASTASIU COSTIN VLAD",
            "ANDREESCU OANA",
            "ANDRONIC LUMINITA CAMELIA",
            "ANDRONIC MARIA LETITIA",
            "ANGHELINA BOGDAN-CRISTIAN",
            "ANTON CARMEN ELENA",
            "ANTONARU CARMEN ELENA",
            "ANTONYA CSABA",
            "APOSTOAIE MIRELA",
            "ARBANAŞ IOANA",
            "ARGĂSEALĂ GEORGIANA",
            "ARHIRE MONA BRIGITTE",
            "ARMĂSAR IOANA PAULA",
            "ARMASELU ANCA",
            "ARON IOAN",
            "ARVATESCU CRISTIAN",
            "ATUDOREI IOANA ANISA",
            "BABA MARIUS NICOLAE",
            "BABA MIRELA CAMELIA",
            "BĂDĂRĂU CARMEN LILIANA",
            "BADAU ADELA",
            "BADAU DANA",
            "BADEA ANAMARIA RALUCA",
            "BADEA MIHAELA",
            "BADICU GEORGIAN",
            "BAICOIANU ALEXANDRA",
            "BALAN FLORIN",
            "BALAN TITUS CONSTANTIN",
            "BALAS MONICA LOREDANA",
            "BALASESCU MARIUS",
            "BALASESCU SIMONA",
            "BALINT ELENA",
            "BALINT LORAND",
            "BALTES LIANA SANDA",
            "BALTESCU CODRUTA ADINA",
            "BARABAS BARNA",
            "BARACAN ADRIAN",
            "BARBU DANIELA MARIANA",
            "BARBU ION",
            "BARBU MAGDALENA",
            "BARBU MARIUS CATALIN",
            "BARBU SILVIU GABRIEL",
            "BARBULESCU ALINA",
            "BARBULESCU OANA",
            "BAROTE LUMINITA",
            "BÂRSAN MARIA IONELA",
            "BÂRSAN MARIA MAGDALENA",
            "BASALIC ELENA-BIANCA",
            "BATRANU PINTEA VLAD",
            "BAZGAN MARIUS",
            "BEDELEAN IOAN BOGDAN",
            "BEDO TIBOR",
            "BEGU TEODORA-MARIA",
            "BELDEAN EMANUELA CARMEN",
            "BELDEAN NICOLAE LAURENTIU",
            "BELDIANU IOLANDA FELICIA",
            "BELIBOU ALEXANDRA",
            "BENCZE ANDREI",
            "BENEA BOGDAN CORNEL",
            "BEȘCHEA ANDREI-GEORGE",
            "BIGIU NICUSOR FLORIN",
            "BILDEA TEODOR STEFAN",
            "BISOC ALINA",
            "BOBESCU ELENA",
            "BOBOC RAZVAN GABRIEL",
            "BOCA LIANA LUMINITA",
            "BOCU RAZVAN",
            "BODI DIANA CRISTINA",
            "BODOC ALICE MAGDALENA",
            "BOER ATTILA LASZLO",
            "BOGATU CRISTINA AURICA",
            "BOGDAN IOANA CORINA",
            "BOLBORICI ANA MARIA",
            "BOLDISOR CRISTIAN NICOLAE",
            "BOLOCAN SORIN IONUT",
            "BONDOC IONESCU ALEXANDRU",
            "BORCAN VIRGIL",
            "BORCOMAN MARIANA",
            "BORZ STELIAN ALEXANDRU",
            "BOSCOIANU MIRCEA",
            "BOSCOR DANA",
            "BOTA OANA ALINA",
            "BOTESCU-SIREŢEANU ILEANA-AURORA",
            "BOTEZATU DAN GEORGE",
            "BOTIANU ANA MARIA",
            "BOTIS MARIUS FLORIN",
            "BOTIS SORINA",
            "BRANEA (ȚACĂ) IOANA - ANTONIA",
            "BRĂNESCU GERONIMO-RĂDUCU",
            "BRATU CIPRIAN",
            "BRATU CONSTANTIN ALEXANDRU",
            "BRATU DRAGOS-VASILE",
            "BRATU MARIA-ALEXANDRA",
            "BRATUCU GABRIEL",
            "BRAUN BARBU CRISTIAN",
            "BRENCI LUMINITA MARIA",
            "BREZEANU ALIN IONUȚ",
            "BRICIU GABRIELA ARABELA",
            "BRICIU VICTOR ALEXANDRU",
            "BUCS LORANT",
            "BUCUR ROMULUS LADISLAU",
            "BUDALA ADRIAN",
            "BUGA CRISTINA MARIA",
            "BUHAICIUC MIHAELA",
            "BUICAN GEORGE RAZVAN",
            "BUJA ELENA",
            "BULARCA ANCA ROXANA",
            "BULARCA MARIA-CRISTINA",
            "BULARCA RAZVAN",
            "BULMEZ ALEXANDRU MIHAI",
            "BURADA MARINELA",
            "BURBEA GEORGIANA-MIHAELA",
            "BURDUHOS BOGDAN GABRIEL",
            "BURLACU MIHAI",
            "BUSUIOCEANU STELIANA",
            "BUTNARIU SILVIU LUIS",
            "BUVNARIU LAVINIA",
            "BUZDUGAN IOANA DIANA",
            "BUZEA CARMEN",
            "CĂLIN (COMȘIȚ) ANDREEA-MIHAELA",
            "CALIN MARIUS DANIEL",
            "CAMPEAN MIHAELA",
            "CÂMPEAN STEFAN-IOAN",
            "CAMPU ADINA",
            "CAMPU VASILE RAZVAN",
            "CANDREA ADINA NICOLETA",
            "CANJA CRISTINA MARIA",
            "CARP MARIUS CATALIN",
            "CATANA DORIN IOAN",
            "CATANESCU ANDREEA CORINA",
            "CATARON ANGEL DORU",
            "CATEANU MIHNEA",
            "CAZACU CHRISTIANA EMILIA",
            "CAZAN ANA MARIA",
            "CAZAN CRISTINA",
            "CERBU CAMELIA",
            "CERNEA NICOLETA",
            "CHEFNEUX GABRIELA",
            "CHELMEA LIGIA",
            "CHEȘCĂ ANTONELLA ELISA",
            "CHICOMBAN CARMEN MIHAELA",
            "CHICOS LUCIA ANTONETA",
            "CHIHAIA GABRIELA-NICOLETA",
            "CHIRCAN ELIZA",
            "CHIRILA ADINA",
            "CHIS ALEXANDRU",
            "CHISALITA DUMITRU",
            "CHITONU GABRIELA CRISTINA",
            "CHITU IOANA BIANCA",
            "CHIVU CATALIN IULIAN",
            "CHIVU CATRINA",
            "CIOARA GHEORGHE ROMEO",
            "CIOBANU CATALIN",
            "CIOBANU DANIELA",
            "CIOBANU ELIZA",
            "CIOBANU RAMONA",
            "CIOCIRLAN ELENA",
            "CIOLOCA ANASTASIA MALINA",
            "CIOPLEIAS BOGDAN-NICOLAE",
            "CIOROIU SILVIU GABRIEL",
            "CIRSTOLOVEAN IOAN LUCIAN",
            "CISMARU LAURA",
            "CIUPALA LAURA ANCA",
            "CIUREA ANDREEA CĂTĂLINA",
            "CIUREA CODRUT IOAN",
            "CIURESCU DANIEL",
            "CLINCIU MIHAELA RODICA",
            "CLINCIU RAMONA",
            "CLOTEA LUMINITA ROXANA",
            "COBELSCHI CĂLIN PAVEL",
            "COCIAS TIBERIU",
            "COCUZ MARIA ELENA",
            "CODREAN CODRIN LEONID",
            "COLIBAN RADU MIHAI",
            "COMAN ALINA",
            "COMAN CLAUDIU",
            "COMAN ECATERINA",
            "COMAN SIMONA",
            "COMĂNESCU IOANA SONIA",
            "COMSIT MIHAI",
            "CONDREA EMILIA-GABRIELA",
            "CONSTANTIN BOGDAN",
            "CONSTANTIN DAN ALEXANDRU",
            "CONSTANTIN CRISTINEL PETRISOR",
            "CONSTANTIN SANDA",
            "CONSTANTINESCU CRISTIAN ADRIAN",
            "CONSTANTINESCU ELENA MIHAELA",
            "CONTIU MIRCEA",
            "CORA IRINGO",
            "COROIU PETRUTA MARIA",
            "COSEREANU CAMELIA",
            "COSTACHE CRISTEA",
            "COSTACHE DELIA",
            "COSTIUC IULIANA",
            "COSTIUC LIVIU",
            "COTARLEA DELIA ANCA",
            "COTFAS DANIEL TUDOR",
            "COTFAS PETRU ADRIAN",
            "COVEI MARIA",
            "CRACIUN ADRIAN VIRGIL",
            "CRETESCU NADIA RAMONA",
            "CRISBĂȘAN ANDREEA-MARIA",
            "CRISTEA DANIEL",
            "CRISTEA LUCIANA",
            "CROITORU CATALIN",
            "CSESZNEK CODRINA",
            "CUCULEA DAN-CRISTIAN",
            "CURTU ALEXANDRU LUCIAN",
            "CUSEN GABRIELA",
            "DAMŞESCU ADRIAN",
            "DANCIU GABRIEL MIHAIL",
            "DANILA ADRIAN",
            "DĂNILĂ DANIEL MIHAI",
            "DAVID LAURA TEODORA",
            "DEACONESCU ANDREA CATALINA",
            "DEACONESCU TUDOR ION",
            "DEACONU ADRIAN MARIUS",
            "DEACONU OVIDIU",
            "DEAKY BOGDAN ALEXANDRU",
            "DEMETER ROBERT",
            "DERCZENI RUDOLF ALEXANDRU",
            "DIACONU IOANA ANDREA",
            "DIACONU LAURENTIU IONEL",
            "DIACONU STEFANIA-ROXANA",
            "DIMA DRAGOS SORIN",
            "DIMA GABRIELA",
            "DIMA LORENA",
            "DIMIENESCU OANA GABRIELA",
            "DIMITRIU MARIA",
            "DIMULESCU CRISTINA",
            "DINCA GHEORGHITA",
            "DINCA MARIUS SORIN",
            "DINU ALEXANDRU",
            "DINU CĂTĂLINA GEORGETA",
            "DINU CRISTINA",
            "DINU ELEONORA ANTOANETA",
            "DINULICA FLORIN",
            "DOBRESCU ADA IOANA",
            "DRACEA LAURA LARISA",
            "DRAGHICI CAMELIA LUCIA",
            "DRAGOI MIRCEA VIOREL",
            "DRAGOMIR GEORGE",
            "DRAGOMIR PÂNZARU CAMELIA CRISTINA",
            "DRUGA CORNELIU NICOLAE",
            "DRUGĂU SORIN",
            "DRUMEA CRISTINA",
            "DUCA LILIANA",
            "DUGULEANA CONSTANTIN",
            "DUGULEANA LILIANA",
            "DUGULEANA MIHAI",
            "DUICU SIMONA SOFIA",
            "DUMITRASCU ADELA-ELIZA",
            "DUMITRASCU DORIN ION",
            "DUMITRESCU FLORIN",
            "DUMITRESCU SILVIU RAZVAN",
            "DUTCA IOAN",
            "EFTIMIE NICOLAE",
            "ELEKES ROBERT GABRIEL",
            "ENACHE DORIN VALTER",
            "ENACHE-DAVID NICOLETA",
            "ENE ANA",
            "ENESCA IOAN ALEXANDRU",
            "ENESCU ADRIAN-GABRIEL",
            "ENESCU IOANA-CLARA",
            "ENESCU RALUCA ELENA",
            "ENOIU RAZVAN SANDU",
            "FALUP PECURARIU CRISTIAN GAVRIL",
            "FALUP PECURARIU OANA GABRIELA",
            "FECHETE FLAVIA",
            "FELEA ALINA SILVANA",
            "FILIP ALEXANDRU CATALIN",
            "FILIP IGNAC - CSABA",
            "FILIP OVIDIU",
            "FÎNTÎNĂ IOANA MARIA",
            "FIRASTRAU IOANA",
            "FLOREA OLIVIA ANA",
            "FLORESCU ADRIANA",
            "FLORESCU MONICA",
            "FLOROIAN LAURA",
            "FOLEA MILENA FLAVIA",
            "FORIS DIANA",
            "FORIS TIBERIU",
            "FRATU MARIANA",
            "FRIEDL ANNAMARIA",
            "FRINCU MADALINA ILEANA",
            "FUGARETU COSMINA",
            "FULGA ANDREEA ILEANA",
            "GABOR CAMELIA",
            "GACEU LIVIU",
            "GALATANU TEOFIL FLORIN",
            "GALMEANU HONORIUS CEZAR",
            "GAROIU STEFAN LUCIAN",
            "GAVRILA CORNEL CATALIN",
            "GAVRIS CLAUDIA MIHAELA",
            "GAVRUS CRISTINA",
            "GHEORGHE CARMEN",
            "GHEORGHE CARMEN ADRIANA",
            "GHEORGHE CATALIN",
            "GHEORGHE DANA MIHAELA",
            "GHEORGHE VASILE",
            "GHEORGHITA (LICHIOIU) IULIANA",
            "GHIGHECI COSTEL CRISTINEL",
            "GHITA DANA ELENA",
            "GHIŢĂ-PÎRNUŢĂ OANA-ANDREEA",
            "GINERICA COSMIN",
            "GIRBACIA FLORIN STELIAN",
            "GIRDAN LAURA",
            "GLIGA CONSTANTIN IOAN",
            "GOTEA MIHAELA",
            "GREŞITĂ CONSTANTIN IRINEL",
            "GRIGORESCU OVIDIU DAN",
            "GRIGORESCU SIMONA",
            "GRIGORESCU SORIN MIHAI",
            "GROSZ WILHELM ROBERT",
            "GUIMAN MARIA VIOLETA",
            "GURAU LIDIA",
            "GUREAN DAN MARIAN",
            "HABA SEVER",
            "HALALISAN AURELIU FLORIN",
            "HENTER RAMONA",
            "HLIPCĂ PETRU",
            "HOGEA MIRCEA DANIEL",
            "HUMINIC GABRIELA",
            "HUMINIC TRAIAN ANGEL",
            "IACOB ANDREEA-BIANCA",
            "IBANESCU DANIELA CORINA",
            "ICHIM TRAIAN",
            "IDOMIR MIHAELA ELENA",
            "IFTENE LIVIU",
            "IFTENI PETRU IULIAN",
            "IGNAT MIHAI",
            "ILEA ANCA-MARIA",
            "ILIE RODICA MARIA",
            "INDREICA ELENA SIMONA",
            "INDREICA VICTOR ADRIAN",
            "ION CATALIN PETREA",
            "ION LAURENTIU-MIHAIL",
            "IONAŞ DIANA GEANINA",
            "IONESCU ALEXANDRU CODRIN",
            "IONESCU ANA MARIA",
            "IONESCU DAN TRAIAN",
            "IONESCU OVIDIU",
            "IORDACHE DANIEL",
            "IORDACHE EUGEN",
            "IORDAN NICOLAE FANI",
            "IOVANAS DANIELA MARIA",
            "IRIMIE CLAUDIA-ALEXANDRINA",
            "IRIMIE IOANA VIOLETA",
            "IRIMIE MARIUS",
            "ISAC IULIANA",
            "ISAC LUMINITA ANISOARA",
            "ISAIA FLORIN",
            "ISAIA GABRIELA AURORA",
            "ISBASOIU ANDREEA",
            "ISOP LAURA-MIHAELA",
            "ISPAS ANA",
            "ISPAS MIHAI",
            "ISPAS NICOLAE",
            "ITU ALINA",
            "ITU CALIN",
            "ITU LUCIAN MIHAI",
            "IVANCESCU RUXANDRA",
            "IVANOVICI LAURENTIU MIHAI",
            "IVASCIUC IOANA SIMONA",
            "IVASCU IRINA MIHAELA",
            "JALIU CODRUTA ILEANA",
            "KAKUCS CRISTIAN",
            "KARACSONY NOEMI",
            "KERTESZ CSABA ZOLTAN",
            "KOLAR VASUDEVA LAURA",
            "KOVACS ATTILA",
            "KRISTALY DOMINIC MIRCEA",
            "LACATUS ADRIAN",
            "LĂCĂTUȘ ANCA MARIA",
            "LACHE SIMONA",
            "LACULICEANU ALEXANDRU-GEORGIAN",
            "LANCEA CAMIL TRAIAN SORIN",
            "LAPTES RAMONA",
            "LATES MIHAI TIBERIU",
            "LAZAR ANAMARIA",
            "LAZAR CORNELIA MAGDALENA",
            "LEAHU CRISTIAN IOAN",
            "LEASU FLORIN GABRIEL",
            "LELUTIU LAURA MIHAELA",
            "LIMBASAN ILEANA GEORGIANA",
            "LINDEMANN SOFIANA IULIA",
            "LITRA ADRIANA VERONICA",
            "LIXANDROIU RADU CONSTANTIN",
            "LORINCZ SIMINA",
            "LOSTUN ALEXANDRA",
            "LUCA MIHAI ALEXANDRU",
            "LUCULESCU MARIUS CRISTIAN",
            "LUNGOCI CARMEN MIHAELA",
            "LUNGU ANTONELA CRISTINA",
            "LUNGULEASA AUREL",
            "LUPSA TATARU DANA ADRIANA",
            "LUPSA TATARU LUCIAN",
            "LUPU DACIANA ANGELICA",
            "LUPU DRAGOȘ",
            "LUPU MIRABELA IOANA",
            "LUPU NICOLETA RALUCA",
            "MACESANU GIGEL",
            "MACHEDON PISU MIHAI",
            "MĂDA STANCA",
            "MAFTEI CARMEN",
            "MAICAN CATALIN IOAN",
            "MAICAN MARIA ANCA",
            "MAIER ALINA",
            "MAJERCSIK LUCIANA",
            "MANCIULEA ILEANA CARMEN",
            "MANDRU LIDIA",
            "MANEA ADELINA LOREDANA",
            "MANEA ELENA LAURA",
            "MANEA EMILIA ADELA",
            "MANEA ROSANA MIHAELA",
            "MANOLICĂ ANA-MARIA",
            "MANTULESCU MARIUS MIHAIL",
            "MARCEANU LUIGI GEO",
            "MARCU MARINA VIORELA",
            "MARDACHE ANDREEA CLAUDIA",
            "MARINESCU DANIELA",
            "MARINESCU NICOLAE ION",
            "MARTOMA ALINA MIRELA",
            "MATEFI ROXANA",
            "MATEI ALEXANDRU",
            "MATEI FLORENTINA",
            "MATEI MADALINA GEORGIANA",
            "MAZAREL ADRIAN",
            "MESESAN SCHMITZ LUIZA IULIANA",
            "MICLAUS STELIANA ROXANA",
            "MICU CORINA SILVIA",
            "MICULESCU RADU",
            "MIHAIL LAURENTIU AUREL",
            "MIHAILESCU MARIA-MIRABELA",
            "MIHAILESCU TEOFIL",
            "MIHALCICA MIRCEA",
            "MIJAICA RALUCA DACIA",
            "MILESAN MIHAELA",
            "MILOSAN IOAN",
            "MINCULETE NICUSOR",
            "MINDRESCU VERONICA",
            "MIRON ( MIOC ) ANA-ALIANA",
            "MISARCA CATALIN",
            "MITREA NICOLETA",
            "MITRICA MARIA",
            "MITU LEONARD",
            "MITU SEBASTIAN-RĂZVAN",
            "MIZGACIU CAMELIA",
            "MOARCĂS GEORGETA",
            "MOASA HORIA",
            "MODRAN HORIA ALEXANDRU",
            "MOGA MARIUS ALEXANDRU",
            "MOJA ADELINA - IOANA",
            "MOLDOVAN (TANTAU) MARA-STEFANIA",
            "MOLDOVAN EDIT ROXANA",
            "MOLDOVAN MACEDON DUMITRU",
            "MONESCU VLAD",
            "MORARIU CRISTIN OLIMPIU",
            "MORARU SORIN AUREL",
            "MOSOI ADRIAN",
            "MOSOIU DANIELA VIORICA",
            "MOTOASCA SEPTIMIU DANIEL",
            "MOTOC DANA",
            "MUNTEAN LIVIU-IULIU",
            "MUNTEAN RADU MIRCEA",
            "MUNTEANU DANIEL",
            "MUNTEANU MIHAELA VIOLETA",
            "MUNTEANU-ICHIM ROXANA ANDREEA",
            "MURESAN VALENTIN",
            "MUSAT ELENA CAMELIA",
            "MUSUROI CRISTIAN LEONARD",
            "NANAU CORINA STEFANIA",
            "NASTAC DORIN CRISTIAN",
            "NASTASA LAURA ELENA",
            "NASTASE GABRIEL",
            "NASTASOIU MIRCEA",
            "NASULEA MARIUS DANIEL",
            "NAUNCEF ALINA MARIA",
            "NEACSU NICOLETA ANDREEA",
            "NEAGOE MIRCEA",
            "NEAGU MIRCEA",
            "NECHIFOR BIANCA ANDREEA",
            "NECHITA FLORENTINA",
            "NECHITA FLORIN MIHAI",
            "NECSOI DANIELA VERONICA",
            "NECULA RADU DAN",
            "NECULA VALENTIN",
            "NECULAU ANDREA ELENA",
            "NECULOIU DANIELA",
            "NECULOIU MARIUS",
            "NEDELOIU TIBERIU",
            "NEGULESCU ORIANA HELENA",
            "NEPOTU GABRIEL LUCIAN",
            "NICOLAE IOANA",
            "NICOLAU ANDRADA CAMELIA",
            "NICOLAU LIANA CRISTINA",
            "NICOLESCU VALERIU NOROCEL",
            "NICULA DAN",
            "NISTOR-ȘERBAN ANDREEA ELENA",
            "NITA MIHAI DANIEL",
            "NIȚOIU LORENA GABRIELA",
            "NUTU MARIA",
            "OANA ALEXANDRU",
            "OANCEA BOGDAN MARIAN",
            "OANCEA GHEORGHE",
            "OGREZEANU IULIAN ALEXANDRU",
            "OGRUTAN PETRE LUCIAN",
            "OLA DANIEL CALIN",
            "OLAH ARTHUR",
            "OLARESCU ALIN",
            "OLTEANU MIRCEA IONUȚ",
            "ONEA GHEORGHE ADRIAN",
            "OPRISESCU SERBAN",
            "ORMENISAN ALEXE NICOLAE",
            "PACURAR CRISTINA MARIA",
            "PACURAR VICTOR DAN",
            "PĂDUREANU VASILE",
            "PANAITE MARA",
            "PANTEA ILEANA",
            "PARV AURICA LUMINITA",
            "PASCU ALEXANDRU",
            "PASCU ALINA MIHAELA",
            "PASCU MIHAI LUCIAN",
            "PASCU MIHAI NICOLAE",
            "PAUN LAURIAN",
            "PAVALACHE ILIE MARIELA",
            "PAVEL ECATERINA",
            "PAVEL GINA MIHAELA",
            "PELIN BOGDAN IULIAN",
            "PERNIU DANA",
            "PETRE ANDREEA",
            "PETRE IOANA",
            "PETRIC PAULA",
            "PETRICI ANDREI VICTOR",
            "PETRITAN ION CATALIN",
            "PISARCIUC CRISTIAN",
            "PIUARU BRENDA-ANDREEA",
            "PLAJER IOANA CRISTINA",
            "PLESCAN COSTEL",
            "PLUMBOTA LAVINIA",
            "PODASCA PETRU CEZARIO",
            "POJALĂ CIPRIAN-VASILE",
            "POLEXA ALEXANDRU-CRIȘAN",
            "POP DANA MIHAELA",
            "POPA IULIAN",
            "POPA BOGDAN",
            "POPA DANIELA (EFS)",
            "POPA DANIELA (PSE)",
            "POPA GEORGE-BOGDAN",
            "POPA LIOARA RALUCA",
            "POPA LUMINITA",
            "POPA ROXANA",
            "POPA STEFAN",
            "POPESCU (GHIUTA) IOANA",
            "POPESCU ANCA",
            "POPESCU MIHAELA VIRGINIA",
            "POPESCU OVIDIU",
            "POPESCU VLAD",
            "POPOVICI BIANCA ELENA",
            "POPOVICI-POPESCU ELENA",
            "POROJAN MIHAELA",
            "POSTELNICU CRISTIAN CEZAR",
            "POTINCU LAURA",
            "POZNA CLAUDIU RADU",
            "PRALEA CRISTIAN",
            "PREDA ULITA ANCA",
            "PROCA ALEXANDRINA MARIA",
            "PUIU ANDREI",
            "PURCARU IOANA MADALINA",
            "RACASAN SERGIU",
            "RĂDOI-ENCEA RALUCA-STEFANIA",
            "RADU CRISTINA IOANA",
            "RADU (MATEI) SIMONA CORINA",
            "RADU ALEXANDRU IONUT",
            "RADU DORIN",
            "RADU FLORIN",
            "RADU LUCIAN",
            "RADU SEBASTIAN",
            "RADUCANU DORINA",
            "RAILEANU SZELES MONICA",
            "RATULEA GEORGETA GABRIELA",
            "RAUTIA IOAN CALIN",
            "REPANOVICI ANGELA",
            "ROATA IONUT CLAUDIU",
            "ROBU DAN NICOLAE",
            "ROGOZEA LILIANA MARCELA",
            "ROMAN NADINNE ALEXANDRA",
            "ROSCA IOAN CALIN",
            "ROSENBERG DAN",
            "RUCSANDA MADALINA",
            "RUNCEANU-ALBU CARMEN CRISTINA",
            "RUS HORATIU",
            "RUSU IULIAN",
            "SABOU FLORIN-LUCIAN-PETRICĂ",
            "SĂFTOIU RĂZVAN GEORGIAN",
            "SARAMET OANA",
            "SARBU FLAVIUS AURELIAN",
            "SASU ADELA",
            "SASU LAURA ELENA",
            "SASU LUCIAN-MIRCEA",
            "SAULESCU RADU GABRIEL",
            "SAVIN DIANA-CRISTINA",
            "SAVU CODRUŢ NICOLAE",
            "SAVU ELENA CRISTINA",
            "SCARNECI-DOMNISORU FLORENTINA",
            "SCARNECIU CAMELIA CORNELIA",
            "SCARNECIU IOAN",
            "SCHWAB-FRÎNCU ANAMARIA",
            "SCRIBA CEZAR",
            "SCUTARU MARIA LUMINITA",
            "SECHEL GABRIELA",
            "SERBAN IOAN",
            "SERBAN IONEL",
            "SERBU CLAUDIA GABRIELA",
            "SIBISAN AURA DANIELA",
            "SIMION GABRIEL",
            "SIMON MARINELA CRISTINA",
            "SINU RALUCA GEORGIANA",
            "SISMAN VIOREL",
            "SITOIU ANDREEA",
            "SOICA ADRIAN",
            "SOICA SIMONA",
            "SOREA DANIELA",
            "SOREA GHEORGHE DAN",
            "SOVA DANIELA",
            "SOVAILA SILVIA",
            "SPÎRCHEZ GEORGETA BIANCA",
            "SPIRCHEZ GHEORGHE COSMIN",
            "SPRIDON DELIA - ELENA",
            "STAN ALEXANDRA",
            "STAN ION GABRIEL",
            "STANCA AUREL CORNEL",
            "STANCIOIU PETRU TUDOR",
            "STANCIU ANCA ELENA",
            "STANCIU ELENA MANUELA",
            "STANCIU MARIANA DOMNICA",
            "STANESCU RUXANDRA",
            "STARETU IONEL",
            "STOICA ROXANA ELENA",
            "STOICANESCU MARIA",
            "STROE FANEL",
            "SUCIU CONSTANTIN",
            "SUCIU MARIA-MAGDALENA",
            "SUCIU TITUS",
            "SUMEDREA SILVIA",
            "SURDU VASILE",
            "SUTEU LIGIA CLAUDIA",
            "SZILAGYI ANA",
            "SZOCS BOTOND CSABA",
            "ȚÂBIAN DANIEL",
            "TABIRCA MARIUS SABIN",
            "TACHE ILEANA",
            "TALPA NICOLAE",
            "TAMAS FLORIN-LUCIAN",
            "ŢĂRANU DAN MARIUS",
            "TARNOVEANU MIRELA ADRIANA",
            "TARULESCU RADU",
            "TARULESCU STELIAN",
            "TAȚA ANITHA",
            "TATU OANA",
            "TAUS DANIEL",
            "TAUS NICOLETA",
            "TECAU ALINA SIMONA",
            "TEODORESCU ANDREEA",
            "TEODORESCU DRAGHICESCU HORATIU",
            "TERESNEU CORNEL CRISTIAN",
            "TERIȘ ȘTEFAN",
            "TESCASIU BIANCA",
            "THIERHEIMER WALTER WILHELM",
            "TIEREAN MIRCEA HORIA",
            "TIMAR JANOS",
            "TIMAR MARIA CRISTINA",
            "TINT DIANA",
            "TISMĂNAR IOANA",
            "TIŢA NICOLESCU GABRIEL",
            "TOADER ADRIAN",
            "TOADER SERBAN-SIXTUS",
            "TODOR RALUCA DANIA",
            "TOFAN DANIEL",
            "TOGANEL GEORGE RADU",
            "TOHANEAN DRAGOS IOAN - EFS",
            "TOMA SEBASTIAN IONUȚ",
            "TOMELE SIMONA CONSTANȚA",
            "TOPALA IOANA ROXANA",
            "TRIFAN ADRIAN",
            "TRUSCA DANIEL DRAGOS",
            "TRUTA CAMELIA",
            "TUCHEL IONUȚ-VLAD",
            "TUDORAN GHEORGHE MARIAN",
            "TULBURE TRAIAN TIBERIU",
            "TURCANU CRISTINA",
            "TURCU IOAN",
            "TURCULET ALINA RALUCA",
            "TUTU DUMITRU CIPRIAN",
            "UDROIU RAZVAN",
            "UNCU IONUT",
            "UNGUREANU CAMELIA",
            "UNGUREANU ELENA",
            "UNGUREANU VALENTIN VASILE",
            "UNIANU ECATERINA MARIA",
            "UNTARU ELENA NICOLETA",
            "URETU NOEMI",
            "URSU PETRONELA ELENA",
            "VALCEA CRISTINA SILVIA",
            "VARCIU MIHAI STELIAN",
            "VARGA IOANA",
            "VARVARICHI LEONA",
            "VASIAN BIANCA IOANA",
            "VASILESCU ANCA",
            "VASILESCU MARIA MAGDALENA",
            "VELEA MARIAN NICOLAE",
            "VELICU RADU GABRIEL",
            "VIZITIU ANAMARIA",
            "VLĂDOIU NASTY MARIAN",
            "VODA DANIELA MARIANA",
            "VOICESCU CORNELIU GEORGE",
            "VOICU NICOLETA",
            "VOINEA MIHAELA",
            "VOLMER MARIUS",
            "VOROVENCII IOSIF",
            "ZAHARIA CORNELIU",
            "ZAHARIA SEBASTIAN MARIAN",
            "ZAMFIRACHE ALEXANDRA",
            "ZELENIUC OCTAVIA",
        };

        // Normalizeaza un nume pentru matching cu lista hardcodata
        private static string NormalizeName(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim().ToUpperInvariant();
            // Collapse multiple spaces
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            return s;
        }

        private static bool IsTitular(string? numeIntreg) =>
            TitulariHardcoded.Contains(NormalizeName(numeIntreg));

        // Query BD care ia toti titularii din view (Titular=1), deduplicat pe ID_Profesor
        // CORECT: foloseste prof.Titular=1, un singur rand per profesor, 3 coloane
        private List<(string Profesor, string Departament, string Facultate)> GetTitulariFromDB(string? facultate, string? departament)
        {
            var result = new List<(string, string, string)>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(@"
                WITH Unici AS (
                    SELECT
                        CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS         AS Profesor,
                        ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Departament,
                        ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Facultate,
                        ROW_NUMBER() OVER(PARTITION BY prof.ID_Profesor ORDER BY prof.ID_Profesor) AS Rn
                    FROM [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof
                    WHERE prof.Titular = 1
                      AND prof.NumeIntreg IS NOT NULL
                      AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                      AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                )
                SELECT Profesor, Departament, Facultate FROM Unici WHERE Rn = 1 ORDER BY Profesor", conn);
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add((r["Profesor"].ToString()!,
                            NormalizeDept(r["Departament"].ToString()),
                            r["Facultate"].ToString()!));
            return result;
        }

        [HttpGet("titulari")]
        public IActionResult GetTitulari(string? anUniv, string? facultate, string? departament)
        {
            var rows = GetTitulariFromDB(facultate, departament);
            return Ok(rows.Select(r => new { NumeSiPrenume = r.Profesor, r.Departament, r.Facultate }));
        }

        [HttpGet("export/titulari")]
        public IActionResult ExportTitulari(string? anUniv, string? facultate, string? departament)
        {
            var rows = GetTitulariFromDB(facultate, departament);
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"),
                new DataColumn("Departament"),
                new DataColumn("Facultate")
            });
            rows.ForEach(r => dt.Rows.Add(r.Profesor, r.Departament, r.Facultate));
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Titulari");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            HdrStyle(ws, 1, dt.Columns.Count);
            ws.Columns().AdjustToContents();
            return Xlsx(wb, "Cadre_Didactice_Titulare.xlsx");
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 7 – COLABORATORI
        // FIX: Profa a zis "foloseste tabelul DRU_Profesor"
        // Colaboratorii = profesori care au ore in statele de functii dar NU sunt titulari (Titular=0)
        // Deduplicat pe ID_Profesor - un singur rand per persoana
        // ═════════════════════════════════════════════════════════════════════

        private const string ColabQ = @"
            WITH ColabIds AS (
                SELECT DISTINCT sf.ID_Profesor
                FROM   [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                LEFT   JOIN [AGSIS].[dbo].[AnUniversitar] au ON sf.id_anuniv = au.ID_AnUniv
                WHERE  sf.xTipCuplaj <> 'CuplajeCareNuMaiExista'
                  AND  (@an = 'Toti' OR UPPER(LTRIM(RTRIM(CAST(au.Denumire AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
            ),
            Unici AS (
                SELECT
                    CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS         AS Profesor,
                    ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Departament,
                    ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Facultate,
                    ROW_NUMBER() OVER(PARTITION BY prof.ID_Profesor ORDER BY prof.ID_Profesor) AS Rn
                FROM   [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof
                INNER  JOIN ColabIds c ON prof.ID_Profesor = c.ID_Profesor
                WHERE  prof.Titular = 0
                  AND  prof.NumeIntreg IS NOT NULL
                  AND  (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                  AND  (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
            )
            SELECT Profesor, Departament, Facultate FROM Unici WHERE Rn = 1 ORDER BY Profesor";

        [HttpGet("colaboratori")]
        public IActionResult GetColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var res = new List<object>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(ColabQ, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                res.Add(new
                {
                    NumeSiPrenume = r["Profesor"].ToString(),
                    Departament = NormalizeDept(r["Departament"].ToString()),
                    Facultate = r["Facultate"].ToString()
                });
            return Ok(res);
        }

        [HttpGet("export/colaboratori")]
        public IActionResult ExportColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"),
                new DataColumn("Departament"),
                new DataColumn("Facultate")
            });
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(ColabQ, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
                dt.Rows.Add(r["Profesor"].ToString(),
                    NormalizeDept(r["Departament"].ToString()),
                    r["Facultate"].ToString());

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Colaboratori");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            HdrStyle(ws, 1, dt.Columns.Count);
            ws.Columns().AdjustToContents();
            return Xlsx(wb, "Cadre_Didactice_Colaboratori.xlsx");
        }

        #endregion

        // ═════════════════════════════════════════════════════════════════════
        #region RAPORT 8 – ANS
        // ═════════════════════════════════════════════════════════════════════

        [HttpGet("date-ans")]
        public IActionResult GetDateAns(int idAnUniv = 45)
        {
            var rows = ReadAnsRows(idAnUniv);
            var result = AggregateAns(rows).Select(p => new {
                p.NumeComplet,
                p.Facultate,
                p.Departament,
                p.GradFunctie,
                DomeniiMapate = p.Fractiuni
                    .Where(kv => AnsIdToCol.ContainsKey(kv.Key) && (AnsIdToCol[kv.Key] - 10) < DomeniiExcel.Length)
                    .ToDictionary(kv => DomeniiExcel[AnsIdToCol[kv.Key] - 10], kv => kv.Value)
            }).ToList();
            return Ok(result);
        }

        [HttpGet("export/raport-ans")]
        public IActionResult ExportAns(int idAnUniv = 45)
        {
            var profesori = AggregateAns(ReadAnsRows(idAnUniv))
                           .OrderBy(p => p.NumeComplet).ToList();
            var wb = BuildAnsWb(); var ws = wb.Worksheets.First();
            int sRow = 9;
            if (profesori.Count > 1) ws.Row(sRow).InsertRowsBelow(profesori.Count - 1);
            for (int i = 0; i < profesori.Count; i++)
            {
                var p = profesori[i]; int rr = sRow + i;
                ws.Cell(rr, 1).Value = i + 1; ws.Cell(rr, 2).Value = p.NumeComplet;
                ws.Cell(rr, 3).Value = ""; ws.Cell(rr, 4).Value = p.GradFunctie;
                ws.Cell(rr, 5).Value = 1; ws.Cell(rr, 6).Value = 0;
                ws.Cell(rr, 7).Value = ""; ws.Cell(rr, 8).Value = p.Facultate;
                ws.Cell(rr, 9).Value = NormalizeDept(p.Departament);
                foreach (var kv in p.Fractiuni) ws.Cell(rr, kv.Key).Value = kv.Value;
                ws.Cell(rr, 50).FormulaA1 = $"=SUM(J{rr}:AW{rr})";
                if (i % 2 != 0)
                    for (int c = 1; c <= 50; c++)
                        ws.Cell(rr, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5f5f5");
            }
            int totR = sRow + profesori.Count;
            ws.Cell(totR, 1).Value = "Total general:"; ws.Cell(totR, 1).Style.Font.Bold = true;
            for (int c = 10; c <= 49; c++) { string cl = ColLetter(c); ws.Cell(totR, c).FormulaA1 = $"=SUM({cl}{sRow}:{cl}{totR - 1})"; ws.Cell(totR, c).Style.Font.Bold = true; }
            ws.Cell(totR, 50).FormulaA1 = $"=SUM(J{totR}:AW{totR})"; ws.Cell(totR, 50).Style.Font.Bold = true;
            using var ms = new MemoryStream(); wb.SaveAs(ms); wb.Dispose();
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Date_ANS_{idAnUniv}.xlsx");
        }

        private List<AnsRow> ReadAnsRows(int idAnUniv)
        {
            var list = new List<AnsRow>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            // ANS foloseste ACEEASI lista de titulari ca raportul Titulari (hardcodata).
            // Se filtreaza in C# dupa IsTitular(). Nu mai putini, nu mai multi.
            using var cmd = new SqlCommand($@"
                WITH {CteDedup}
                SELECT
                    ISNULL(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Prof.ID '+CAST(sf.ID_Profesor AS VARCHAR)) AS Profesor,
                    ISNULL(prof.DenumireGradDidactic,'')          AS Grad,
                    ISNULL(prof.DenumireFacultate COLLATE Romanian_CI_AS,'Nespecificat') AS Facultate,
                    ISNULL(prof.DenumireCatedra COLLATE Romanian_CI_AS,  'Nespecificat') AS Departament,
                    ISNULL(sf.NrOreConventionale,0)               AS OreConv,
                    sf.id_metaspecializare                        AS IdMeta
                FROM SfDedup sf
                LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof ON sf.ID_Profesor=prof.ID_Profesor
                WHERE sf.RnPost=1 AND sf.id_anuniv=@id", conn);
            cmd.Parameters.AddWithValue("@id", idAnUniv);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var numeProf = r["Profesor"].ToString()!;
                // Filtru strict: doar profesorii din lista hardcodata de titulari
                if (!IsTitular(numeProf)) continue;
                int meta = r["IdMeta"] != DBNull.Value ? Convert.ToInt32(r["IdMeta"]) : 0;
                if (!MappingMetaspec.TryGetValue(meta, out int ansId)) continue;
                if (!AnsIdToCol.ContainsKey(ansId)) continue;
                list.Add(new AnsRow
                {
                    Profesor = numeProf,
                    Grad = r["Grad"].ToString()!,
                    Facultate = r["Facultate"].ToString()!,
                    Departament = r["Departament"].ToString()!,
                    OreConv = Convert.ToDecimal(r["OreConv"]),
                    IdAns = ansId
                });
            }
            return list;
        }

        private List<AnsProf> AggregateAns(List<AnsRow> rows) =>
            rows.GroupBy(x => x.Profesor).Select(g =>
            {
                var best = g.GroupBy(x => new { x.Departament, x.Facultate })
                    .Select(d => new { d.Key.Departament, d.Key.Facultate, Tot = d.Sum(x => x.OreConv), Grad = d.OrderByDescending(x => x.OreConv).First().Grad })
                    .OrderByDescending(d => d.Tot).First();
                var byCol = new Dictionary<int, decimal>();
                foreach (var row in g) { int col = AnsIdToCol[row.IdAns]; if (!byCol.ContainsKey(col)) byCol[col] = 0m; byCol[col] += row.OreConv; }
                decimal total = byCol.Values.Sum();
                var frac = new Dictionary<int, decimal>();
                if (total > 0)
                {
                    int mx = byCol.OrderByDescending(kv => kv.Value).First().Key; decimal s = 0;
                    foreach (var kv in byCol) { if (kv.Key == mx) continue; decimal f = Math.Round(kv.Value / total, 2); frac[kv.Key] = f; s += f; }
                    frac[mx] = Math.Round(1m - s, 2);
                }
                return new AnsProf { NumeComplet = g.Key, Departament = best.Departament, Facultate = best.Facultate, GradFunctie = MapGrad(best.Grad), Fractiuni = frac };
            }).ToList();

        private static string MapGrad(string g)
        {
            if (string.IsNullOrWhiteSpace(g)) return "Asistent";
            string s = g.ToLower();
            if (s.Contains("profesor")) return "Profesor";
            if (s.Contains("conferentiar")) return "Conferentiar";
            if (s.Contains("lector") || s.Contains("sef lucrari")) return "Lector/Sef de lucrari (SL)";
            if (s.Contains("asistent de cercetare")) return "Asistent de cercetare";
            if (s.Contains("asistent")) return "Asistent";
            if (s.Contains("preparator")) return "Preparator";
            if (s.Contains("cercetator")) return "Cercetator";
            return "Asistent";
        }

        private XLWorkbook BuildAnsWb()
        {
            var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("CD DRU");
            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare"; ws.Range(2, 1, 2, 50).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov"; ws.Range(3, 1, 3, 6).Merge();
            ws.Cell(5, 1).Value = "Nr.\nCrt."; ws.Cell(5, 2).Value = "Nume si prenume"; ws.Cell(5, 3).Value = "CNP";
            ws.Cell(5, 4).Value = "Functie"; ws.Cell(5, 5).Value = "Forma angajare"; ws.Cell(5, 6).Value = "Conducator doctorat";
            ws.Cell(5, 7).Value = "Varsta"; ws.Cell(5, 8).Value = "Facultate"; ws.Cell(5, 9).Value = "Departament";
            ws.Cell(5, 10).Value = "Matematica si stiinte ale naturii"; ws.Cell(5, 15).Value = "Stiinte ingineresti";
            ws.Cell(5, 22).Value = "Stiinte biologice si biomedicale"; ws.Cell(5, 28).Value = "Stiinte sociale";
            ws.Cell(5, 37).Value = "Stiinte umaniste si arte"; ws.Cell(5, 50).Value = "Total";
            foreach (var (r1, c1, r2, c2) in new[] { (5, 1, 7, 1), (5, 2, 7, 2), (5, 3, 7, 3), (5, 4, 7, 4), (5, 5, 7, 5), (5, 6, 7, 6), (5, 7, 7, 7), (5, 8, 7, 8), (5, 9, 7, 9), (5, 10, 5, 14), (5, 15, 5, 21), (5, 22, 5, 27), (5, 28, 5, 36), (5, 37, 5, 49), (5, 50, 7, 50) })
                ws.Range(r1, c1, r2, c2).Merge();
            string[] sub = { "Matematica", "Informatica", "Fizica", "Chimie si inginerie chimica", "Stiintele pamantului si atmosferei", "Inginerie civila", "Inginerie electrica, electronica si telecomunicatii", "Inginerie geologica, mine, petrol si gaze", "Ingineria transporturilor", "Ingineria resurselor vegetale si animale", "Ingineria sistemelor, calculatoare si tehnologia informatiei", "Inginerie mecanica, mecatronica, inginerie industriala si management", "Biologie", "Biochimie", "Medicina", "Medicina veterinara", "Medicina dentara", "Farmacie", "Stiinte juridice", "Stiinte administrative", "Stiinte ale comunicarii", "Sociologie", "Stiinte politice", "Stiinte militare, informatii si ordine publica", "Stiinte economice (doar Cibernetica, statistica si informatica economica)", "Stiinte economice (fara Cibernetica, statistica si informatica economica)", "Psihologie si stiinte comportamentale", "Filologie", "Filosofie", "Istorie", "Teologie", "Studii culturale", "Arhitectura si urbanism", "Arte vizuale (fara Istoria si teoria artei)", "Arte vizuale (doar Istoria si teoria artei)", "Teatru si artele spectacolului", "Cinematografie si media", "Muzica (doar Interpretare muzicala)", "Muzica (fara Interpretare muzicala)", "Stiintele Sportului si Educatiei Fizice" };
            for (int i = 0; i < sub.Length; i++) { ws.Cell(6, 10 + i).Value = sub[i]; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }
            for (int i = 0; i < 9; i++) ws.Cell(8, i + 1).Value = ((char)('A' + i)).ToString();
            for (int i = 0; i < 41; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, 50).Value = "40";
            var hdr = XLColor.FromHtml(BrandColor);
            for (int row = 5; row <= 8; row++) for (int c = 1; c <= 50; c++)
            {
                ws.Cell(row, c).Style.Font.Bold = true; ws.Cell(row, c).Style.Fill.BackgroundColor = XLColor.White;
                ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, c).Style.Alignment.WrapText = true;
                ws.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(row, c).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }
            ws.Range(5, 1, 5, 50).Style.Fill.BackgroundColor = hdr; ws.Range(5, 1, 5, 50).Style.Font.FontColor = XLColor.White;
            ws.Column(1).Width = 5; ws.Column(2).Width = 30; ws.Column(3).Width = 14; ws.Column(4).Width = 22;
            ws.Column(5).Width = 10; ws.Column(6).Width = 12; ws.Column(7).Width = 8; ws.Column(8).Width = 28; ws.Column(9).Width = 28;
            for (int c = 10; c <= 50; c++) ws.Column(c).Width = 12;
            return wb;
        }

        private class AnsRow { public string Profesor = ""; public string Grad = ""; public string Facultate = ""; public string Departament = ""; public decimal OreConv; public int IdAns; }
        private class AnsProf { public string NumeComplet = ""; public string Departament = ""; public string Facultate = ""; public string GradFunctie = ""; public Dictionary<int, decimal> Fractiuni = new(); }

        #endregion
    }
}