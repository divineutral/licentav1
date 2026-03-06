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
        //   ppm  = Post_Profesor_Materie - FILTRAT pe ID_Post_Profesor_Materie (nu pe ID_Profesor!)
        //          Join-ul e pe ID_Post_Profesor_Materie = un singur rand exact, fara multiplicare
        //          Folosit ca fallback pentru NumeIntreg cand prof e null
        private const string SqlJoin = @"
            FROM SfDedup sf
            LEFT JOIN [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof
                   ON sf.ID_Profesor = prof.ID_Profesor
            LEFT JOIN (
                SELECT ID_Post_Profesor_Materie, ID_Profesor,
                       CAST(NumeIntreg          AS NVARCHAR(500)) COLLATE Romanian_CI_AS AS NumeIntreg,
                       CAST(DenumireFacultate   AS NVARCHAR(500)) COLLATE Romanian_CI_AS AS DenumireFacultate,
                       CAST(DenumireCatedraProfesor AS NVARCHAR(500)) COLLATE Romanian_CI_AS AS DenumireCatedraProfesor
                FROM [agsis_dw].[dbo].[Post_Profesor_Materie]
            ) ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
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
              AND (@dept  = 'Toti' OR
                UPPER(LTRIM(RTRIM(CASE
                    WHEN UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) LIKE 'CATEDRA DE %'
                        THEN 'DEPARTAMENTUL ' + LTRIM(SUBSTRING(UPPER(LTRIM(RTRIM(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS))), 12, 500))
                    WHEN UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) LIKE 'DEPARTAMENTUL %'
                        THEN UPPER(LTRIM(RTRIM(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                    ELSE 'DEPARTAMENTUL ' + UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS,''))))
                END))) = UPPER(LTRIM(RTRIM(@dept))))
              AND (@prof  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(COALESCE(CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS, ppm.NumeIntreg),'')))) = UPPER(LTRIM(RTRIM(@prof))))
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

        // Normalizeaza orice forma de departament din BD -> "Departamentul X"
        // Surse BD: "Catedra de X", "CATEDRA DE X", "Departamentul X", "DEPARTAMENTUL X", "X" (fara prefix)
        private static string NormalizeDept(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "Nespecificat";
            s = s.Trim();
            // Elimina prefixe cunoscute (case-insensitive)
            string[] prefixes = {
                "Catedra de ", "Catedra De ", "CATEDRA DE ", "CATEDRA de ",
                "Departamentul ", "DEPARTAMENTUL ", "Departament ",
                "DEPARTAMENTUL DE ", "Departamentul de "
            };
            foreach (var p in prefixes)
                if (s.StartsWith(p, StringComparison.OrdinalIgnoreCase))
                    return "Departamentul " + s.Substring(p.Length).Trim();
            // Daca nu are prefix, adauga "Departamentul "
            if (!s.StartsWith("Departamentul", StringComparison.OrdinalIgnoreCase))
                return "Departamentul " + s;
            return s;
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
        public IActionResult GetDepartamente(string? anUniv, string? numeFacultate, string? ciclu, string? formaInv)
        {
            var list = new List<string>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand($@"
                WITH {CteDedup},
                DeptCiclu AS (
                    SELECT DISTINCT
                        UPPER(LTRIM(RTRIM(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) AS Dept,
                        MAX(sf.NrAnStudii) AS MaxAn
                    {SqlJoin}
                    WHERE  sf.RnPost = 1 AND prof.DenumireCatedra COLLATE Romanian_CI_AS IS NOT NULL
                      AND  (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(CAST(au.Denumire AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
                      AND  (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                      AND  (@formaId = 0   OR sf.ID_TipFormaInv = @formaId)
                    GROUP BY UPPER(LTRIM(RTRIM(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                )
                SELECT Dept FROM DeptCiclu
                WHERE (@ciclu = 'Toti'
                    OR (@ciclu = 'Licenta' AND MaxAn >= 3)
                    OR (@ciclu = 'Master'  AND MaxAn <= 2))
                ORDER BY Dept", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            cmd.Parameters.AddWithValue("@ciclu", ciclu ?? "Toti");
            cmd.Parameters.AddWithValue("@formaId", FormaToId(formaInv));
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var v = r[0]?.ToString(); if (!string.IsNullOrWhiteSpace(v)) list.Add(NormalizeDept(v)); }
            return Ok(list.Distinct().OrderBy(x => x).ToList());
        }

        // Helper SQL pentru normalizare departament in query (elimina prefix Catedra/Departamentul)
        private const string SqlNormDept = @"
            UPPER(LTRIM(RTRIM(
                CASE
                    WHEN UPPER(LTRIM(RTRIM(prof.DenumireCatedra COLLATE Romanian_CI_AS))) LIKE 'CATEDRA DE %'
                        THEN 'DEPARTAMENTUL ' + LTRIM(SUBSTRING(LTRIM(RTRIM(prof.DenumireCatedra COLLATE Romanian_CI_AS)), 12, 500))
                    WHEN UPPER(LTRIM(RTRIM(prof.DenumireCatedra COLLATE Romanian_CI_AS))) LIKE 'DEPARTAMENTUL %'
                        THEN UPPER(LTRIM(RTRIM(prof.DenumireCatedra COLLATE Romanian_CI_AS)))
                    ELSE 'DEPARTAMENTUL ' + UPPER(LTRIM(RTRIM(prof.DenumireCatedra COLLATE Romanian_CI_AS)))
                END
            )))";

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
                        UPPER(LTRIM(RTRIM(CAST(sf.DenumireSpecializare AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) AS Spec,
                        MAX(sf.NrAnStudii) AS MaxAn
                    {SqlJoin}
                    WHERE  sf.RnPost = 1 AND sf.DenumireSpecializare COLLATE Romanian_CI_AS IS NOT NULL
                      AND  (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(CAST(au.Denumire AS NVARCHAR(500)) COLLATE Romanian_CI_AS))) = UPPER(LTRIM(RTRIM(@an))))
                      AND  (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                      AND  (@dept = 'Toti' OR {SqlNormDept} = UPPER(LTRIM(RTRIM(@dept))))
                    GROUP BY UPPER(LTRIM(RTRIM(CAST(sf.DenumireSpecializare AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                )
                SELECT Spec FROM SpecCiclu
                WHERE (@ciclu = 'Toti'
                    OR (@ciclu = 'Licenta' AND MaxAn >= 3)
                    OR (@ciclu = 'Master'  AND MaxAn <= 2))
                ORDER BY Spec", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", numeDepartament?.ToUpper() ?? "Toti");
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
                  AND (@dept = 'Toti' OR
                    UPPER(LTRIM(RTRIM(CASE
                        WHEN UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) LIKE 'CATEDRA DE %'
                            THEN 'DEPARTAMENTUL ' + LTRIM(SUBSTRING(UPPER(LTRIM(RTRIM(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS))), 12, 500))
                        WHEN UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) LIKE 'DEPARTAMENTUL %'
                            THEN UPPER(LTRIM(RTRIM(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS)))
                        ELSE 'DEPARTAMENTUL ' + UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra AS NVARCHAR(500)) COLLATE Romanian_CI_AS,''))))
                    END))) = UPPER(LTRIM(RTRIM(CAST(@dept AS NVARCHAR(500)) COLLATE Romanian_CI_AS))))
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
                   CAST(SUM(OreConv) * 14 AS DECIMAL(10,2)) AS TotalSem
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
                    TotalAnualOreConv = Convert.ToDecimal(r["TotalSem"])
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
                    Convert.ToDecimal(r["TotalSem"]));

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
            "ARMASELU ANCA",
            "ARMĂSAR IOANA PAULA",
            "ARON IOAN",
            "ARVATESCU CRISTIAN",
            "ATUDOREI IOANA ANISA",
            "BABA MARIUS NICOLAE",
            "BABA MIRELA CAMELIA",
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
            "BRĂNESCU GERONIMO-RĂDUCU",
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
            "BÂRSAN MARIA IONELA",
            "BÂRSAN MARIA MAGDALENA",
            "BĂDĂRĂU CARMEN LILIANA",
            "CALIN MARIUS DANIEL",
            "CAMPEAN MIHAELA",
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
            "COMSIT MIHAI",
            "COMĂNESCU IOANA SONIA",
            "CONDREA EMILIA-GABRIELA",
            "CONSTANTIN BOGDAN",
            "CONSTANTIN CRISTINEL PETRISOR",
            "CONSTANTIN DAN ALEXANDRU",
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
            "CÂMPEAN STEFAN-IOAN",
            "CĂLIN (COMȘIȚ) ANDREEA-MIHAELA",
            "DAMŞESCU ADRIAN",
            "DANCIU GABRIEL MIHAIL",
            "DANILA ADRIAN",
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
            "DINU CRISTINA",
            "DINU CĂTĂLINA GEORGETA",
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
            "DĂNILĂ DANIEL MIHAI",
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
            "FÎNTÎNĂ IOANA MARIA",
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
            "LĂCĂTUȘ ANCA MARIA",
            "MACESANU GIGEL",
            "MACHEDON PISU MIHAI",
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
            "MĂDA STANCA",
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
            "POPA BOGDAN",
            "POPA DANIELA (EFS)",
            "POPA DANIELA (PSE)",
            "POPA GEORGE-BOGDAN",
            "POPA IULIAN",
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
            "PĂDUREANU VASILE",
            "RACASAN SERGIU",
            "RADU (MATEI) SIMONA CORINA",
            "RADU ALEXANDRU IONUT",
            "RADU CRISTINA IOANA",
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
            "RĂDOI-ENCEA RALUCA-STEFANIA",
            "SABOU FLORIN-LUCIAN-PETRICĂ",
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
            "SPIRCHEZ GHEORGHE COSMIN",
            "SPRIDON DELIA - ELENA",
            "SPÎRCHEZ GEORGETA BIANCA",
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
            "SĂFTOIU RĂZVAN GEORGIAN",
            "TABIRCA MARIUS SABIN",
            "TACHE ILEANA",
            "TALPA NICOLAE",
            "TAMAS FLORIN-LUCIAN",
            "TARNOVEANU MIRELA ADRIANA",
            "TARULESCU RADU",
            "TARULESCU STELIAN",
            "TATU OANA",
            "TAUS DANIEL",
            "TAUS NICOLETA",
            "TAȚA ANITHA",
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
            "TUCHEL IONU?-VLAD",
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
            "ŢĂRANU DAN MARIUS",
            "ȚÂBIAN DANIEL",
        };

        // Normalizeaza un nume pentru matching cu lista hardcodata
        private static string NormalizeName(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim().ToUpperInvariant();
            s = s.Replace("*", "");          // elimina asteriscuri (ex: MUNTEANU DANIEL*)
            while (s.Contains("  ")) s = s.Replace("  ", " "); // spatii duble
            return s.Trim();
        }

        private static bool IsTitular(string? numeIntreg) =>
            TitulariHardcoded.Contains(NormalizeName(numeIntreg));

        // Query BD cu filtru pe lista hardcodata de titulari (sursa de adevar furnizata de prof. Ionita)
        // Matching: NormalizeName(NumeIntreg) in TitulariHardcoded -> ~778 persoane
        private List<(string Profesor, string Departament, string Facultate)> GetTitulariFromDB(string? facultate, string? departament)
        {
            var result = new List<(string, string, string)>();
            using var conn = new SqlConnection(_connectionString); conn.Open();
            using var cmd = new SqlCommand(@"
                WITH Unici AS (
                    SELECT
                        CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS AS Profesor,
                        ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Departament,
                        ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Facultate,
                        ROW_NUMBER() OVER(PARTITION BY prof.ID_Profesor ORDER BY prof.ID_Profesor) AS Rn
                    FROM [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof
                    WHERE prof.NumeIntreg IS NOT NULL
                      AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
                      AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
                )
                SELECT Profesor, Departament, Facultate FROM Unici WHERE Rn = 1 ORDER BY Profesor", conn);
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var nume = r["Profesor"].ToString()!;
                if (IsTitular(nume))
                    result.Add((nume,
                                NormalizeDept(r["Departament"].ToString()),
                                r["Facultate"].ToString()!));
            }
            return result;
        }

        [HttpGet("titulari")]
        public IActionResult GetTitulari(string? anUniv, string? facultate, string? departament)
        {
            var rows = GetTitulariFromDB(facultate, departament);
            return Ok(rows.Select(r => new {
                profesor = r.Profesor,
                departament = r.Departament,
                facultate = r.Facultate,
                grad = ""
            }));
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

        // Colaboratori = profesori din View_ProfesoriActivi_CF care NU sunt titulari
        // View-ul contine o inregistrare per (profesor, departament) deci un prof
        // apare de mai multe ori daca e afiliat la mai multe departamente -> ~742 randuri
        private const string ColabQ = @"
            SELECT
                CAST(prof.NumeIntreg AS NVARCHAR(500)) COLLATE Romanian_CI_AS AS Profesor,
                ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Departament,
                ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS, 'Nespecificat') AS Facultate,
                ISNULL(prof.DenumireGradDidactic, '') AS Grad
            FROM [agsis_dw].[dbo].[View_ProfesoriActivi_CF] prof
            WHERE prof.Titular = 0
              AND prof.NumeIntreg IS NOT NULL
              AND (@fac  = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireFacultate AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@fac))))
              AND (@dept = 'Toti' OR UPPER(LTRIM(RTRIM(ISNULL(CAST(prof.DenumireCatedra   AS NVARCHAR(500)) COLLATE Romanian_CI_AS,'')))) = UPPER(LTRIM(RTRIM(@dept))))
            ORDER BY Profesor, Departament";

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
                    profesor = r["Profesor"].ToString(),
                    departament = NormalizeDept(r["Departament"].ToString()),
                    facultate = r["Facultate"].ToString(),
                    grad = r["Grad"].ToString()
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
                new DataColumn("Facultate"),
                new DataColumn("Grad Didactic")
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
                    r["Facultate"].ToString(),
                    r["Grad"].ToString());

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
            // Pasul 1: agreg din BD (pentru Facultate, Departament, OreConv)
            var fromDb = AggregateAns(ReadAnsRows(idAnUniv))
                        .ToDictionary(p => NormalizeName(p.NumeComplet), p => p);

            // Pasul 2: include TOTI profesorii din referinta profei (740),
            // nu doar cei cu ore in BD
            var profesori = new List<AnsProf>();
            foreach (var kv in AnsRef)
            {
                var nameKey = kv.Key;
                if (fromDb.TryGetValue(nameKey, out var dbEntry))
                {
                    // Profesorul are ore in BD -> luam Facultate/Departament din BD
                    // dar grade si fractiuni din AnsRef (deja setat in AggregateAns)
                    profesori.Add(dbEntry);
                }
                else
                {
                    // Profesorul nu are ore in BD pentru acest an -> adaugam cu date din referinta
                    var frac = kv.Value.Domenii.ToDictionary(d => d.Key + 9, d => d.Value);
                    profesori.Add(new AnsProf
                    {
                        NumeComplet = kv.Key,
                        Departament = "",
                        Facultate = "",
                        GradFunctie = kv.Value.Grad,
                        Fractiuni = frac
                    });
                }
            }
            profesori = profesori.OrderBy(p => p.NumeComplet).ToList();
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
                var nameKey = NormalizeName(g.Key);
                var best = g.GroupBy(x => new { x.Departament, x.Facultate })
                    .Select(d => new { d.Key.Departament, d.Key.Facultate, Tot = d.Sum(x => x.OreConv), Grad = d.OrderByDescending(x => x.OreConv).First().Grad })
                    .OrderByDescending(d => d.Tot).First();

                // Fractiunile si gradul: daca exista in referinta exacta a profei, folosim acelea
                Dictionary<int, decimal> frac;
                string gradFunctie;
                if (AnsRef.TryGetValue(nameKey, out var refData))
                {
                    // Folosim datele exacte din fisierul de referinta
                    // refData.Domenii are ansId (1-40); convertim la coloana Excel (ansId+9)
                    frac = refData.Domenii.ToDictionary(kv => kv.Key + 9, kv => kv.Value);
                    gradFunctie = refData.Grad; // ex: "Prof. dr.", "Conf. dr."
                }
                else
                {
                    // Fallback: calcul din ore conventionale
                    var byCol = new Dictionary<int, decimal>();
                    foreach (var row in g) { int col = AnsIdToCol[row.IdAns]; if (!byCol.ContainsKey(col)) byCol[col] = 0m; byCol[col] += row.OreConv; }
                    decimal total = byCol.Values.Sum();
                    frac = new Dictionary<int, decimal>();
                    if (total > 0)
                    {
                        int mx = byCol.OrderByDescending(kv => kv.Value).First().Key; decimal s = 0;
                        foreach (var kv in byCol) { if (kv.Key == mx) continue; decimal f = Math.Round(kv.Value / total, 2); frac[kv.Key] = f; s += f; }
                        frac[mx] = Math.Round(1m - s, 2);
                    }
                    gradFunctie = MapGrad(best.Grad);
                }
                return new AnsProf { NumeComplet = g.Key, Departament = best.Departament, Facultate = best.Facultate, GradFunctie = gradFunctie, Fractiuni = frac };
            }).ToList();

        // Mapeaza DenumireGradDidactic din BD -> formatul exact din fisa ANS
        // Formate din fisierul de referinta al profei: "Prof. dr.", "Conf. dr.",
        // "Lect. dr.", "Sef lucr. dr.", "Asist. dr.", "Asist.drd."
        private static string MapGrad(string g)
        {
            if (string.IsNullOrWhiteSpace(g)) return "Asist. dr.";
            string s = g.ToLower().Trim();
            if (s.Contains("profesor")) return "Prof. dr.";
            if (s.Contains("conferentiar") || s.Contains("conferen")) return "Conf. dr.";
            if (s.Contains("sef lucr") || s.Contains("sef de lucr")) return "Sef lucr. dr.";
            if (s.Contains("lector")) return "Lect. dr.";
            if (s.Contains("drd") || s.Contains("doctorand")) return "Asist.drd.";
            if (s.Contains("asistent")) return "Asist. dr.";
            return "Asist. dr.";
        }

        private XLWorkbook BuildAnsWb()
        {
            var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("CD DRU");
            // Row 2: titlu
            ws.Cell(2, 1).Value = "\nAnexa 1. Tabel instituţional privind normarea şi activitatea de cercetare a cadrelor didactice şi de cercetare din universitate (raportare IC2015)";
            ws.Range(2, 1, 2, 50).Merge();
            // Row 3: universitate
            ws.Cell(3, 1).Value = "Universitatea……………………………………... ";
            ws.Range(3, 1, 3, 50).Merge();
            // Row 4: nota (exacta din fisierul de referinta al profei)
            ws.Cell(4, 1).Value = "NOTĂ: \nSe includ în tabel toate cadrele didactice şi de cercetare titulare (inclusiv cadrele didactice angajate cu normă întreagă, cu un contract pe perioadă determinată conform art.294, din LEN 1/2011, valid în perioada de raportare). Pentru facilitarea verificărilor interne recomandăm gruparea pe facultăţi, respectiv departamente. \nFiecare cadru didactic sau de cercetare al universităţii se raportează pe un singur rând.\nCompletarea în câmpurile aferente col.D-F din tabel se realizează prin selectarea valorii corespunzatoare din lista predefinita in col.D, respectiv completarea cu numarul corespunzator valorii din listele predefinite in col.E si col.F.\nVă rugăm să completați numai spațiile marcate cu culoarea galben. Puteţi insera rânduri în document, doar înainte de rândul cu TOTAL, prin selectarea unui rând formatat (marcat cu culoarea galben) şi apoi Copy & Insert Copied Cells.";
            ws.Range(4, 1, 4, 9).Merge();
            ws.Cell(4, 10).Value = "NOTĂ: \nIMPORTANT! Vă rugăm să completați în prima fază, în sheet-ul \"Ramuri-Ştiinţă\", valoarea \"1\"în col.C, pentru ramurile de ştiinţă  în care există programe de studii la nivel de universitate. \nÎn cazul personalului didactic care predă la programe aparţinând mai multor ramuri de ştiinţă, se raportează fracţionat, în funcţie de ponderea activităţilor aferente programelor respective în postul de bază din statul de funcţii  (maximum două zecimale, exemplu: jumatate de norma = 0,50), suma fracţiilor pentru un cadru didactic având valoarea 1 (col.40).\nVă rugăm să completați numai spațiile marcate cu culoarea galben. Puteţi insera rânduri în document, doar înainte de rândul cu TOTAL, prin selectarea unui rând formatat (marcat cu culoarea galben) şi apoi Copy & Insert Copied Cells.";
            ws.Range(4, 10, 4, 49).Merge();
            // Row 5: header-uri principale (exact ca in referinta)
            ws.Cell(5, 1).Value = "Nr. \nCrt.";
            ws.Cell(5, 2).Value = "Nume si prenume cadru didactic";
            ws.Cell(5, 3).Value = "CNP";
            ws.Cell(5, 4).Value = "Funcţie cadru didactic sau cercetare";
            ws.Cell(5, 5).Value = "Forma de angajare";
            ws.Cell(5, 6).Value = "Calitate conducator doctorat";
            ws.Cell(5, 7).Value = "Varsta";
            ws.Cell(5, 8).Value = "Facultate";
            ws.Cell(5, 9).Value = "Departament";
            ws.Cell(5, 10).Value = "Matematică şi ştiinţe ale naturii";
            ws.Cell(5, 15).Value = "Ştiinţe inginereşti";
            ws.Cell(5, 22).Value = "Ştiinţe biologice şi biomedicale";
            ws.Cell(5, 28).Value = "Ştiinţe sociale";
            ws.Cell(5, 37).Value = "Ştiinţe umaniste şi arte";
            ws.Cell(5, 50).Value = "Total";
            // Merge header rows 5-7 for first 9 cols + groups
            foreach (var (r1, c1, r2, c2) in new[] {
                (5,1,7,1),(5,2,7,2),(5,3,7,3),(5,4,7,4),(5,5,7,5),
                (5,6,7,6),(5,7,7,7),(5,8,7,8),(5,9,7,9),
                (5,10,5,14),(5,15,5,21),(5,22,5,27),(5,28,5,36),(5,37,5,49),(5,50,7,50)
            }) ws.Range(r1, c1, r2, c2).Merge();
            // Row 6: sub-domenii
            string[] sub = {
                "Matematică","Informatică","Fizică","Chimie şi inginerie chimică",
                "Ştiinţele pământului şi atmosferei","Inginerie civilă",
                "Inginerie electrică, electronică şi telecomunicaţii",
                "Inginerie geologică, mine, petrol şi gaze","Ingineria transporturilor",
                "Ingineria resurselor vegetale şi animale",
                "Ingineria sistemelor, calculatoare şi tehnologia informaţiei",
                "Inginerie mecanică, mecatronică, inginerie industrială şi management",
                "Biologie","Biochimie","Medicină","Medicină veterinară","Medicină dentară","Farmacie",
                "Ştiinţe juridice","Ştiinţe administrative","Ştiinţe ale comunicării","Sociologie",
                "Ştiinţe politice","Ştiinţe militare, informaţii şi ordine publică",
                "Ştiinţe economice (doar Cibernetică, statistică şi informatică economică)",
                "Ştiinţe economice (fără  Cibernetică, statistică şi informatică economică)",
                "Psihologie şi ştiinţe comportamentale","Filologie","Filosofie","Istorie",
                "Teologie","Studii culturale","Arhitectură şi urbanism",
                "Arte vizuale (fără Istoria şi teoria artei)","Arte vizuale (doar Istoria şi teoria artei)",
                "Teatru şi artele spectacolului","Cinematografie şi media",
                "Muzică (doar Interpretare muzicală)","Muzică (fără Interpretare muzicală)",
                "Ştiinţele Sportului şi Educaţiei Fizice"
            };
            for (int i = 0; i < sub.Length; i++) { ws.Cell(6, 10 + i).Value = sub[i]; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }
            // Row 8: litere A-I si numere 1-40
            for (int i = 0; i < 9; i++) ws.Cell(8, i + 1).Value = ((char)('A' + i)).ToString();
            ws.Cell(8, 50).Value = "40";
            // Stilizare header rows
            var hdr = XLColor.FromHtml(BrandColor);
            for (int row = 5; row <= 8; row++) for (int c = 1; c <= 50; c++)
            {
                ws.Cell(row, c).Style.Font.Bold = true;
                ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, c).Style.Alignment.WrapText = true;
                ws.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(row, c).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }
            ws.Range(5, 1, 5, 50).Style.Fill.BackgroundColor = hdr;
            ws.Range(5, 1, 5, 50).Style.Font.FontColor = XLColor.White;
            // Coloane widths
            ws.Column(1).Width = 5; ws.Column(2).Width = 30; ws.Column(3).Width = 14; ws.Column(4).Width = 22;
            ws.Column(5).Width = 10; ws.Column(6).Width = 12; ws.Column(7).Width = 8; ws.Column(8).Width = 28; ws.Column(9).Width = 28;
            for (int c = 10; c <= 50; c++) ws.Column(c).Width = 12;
            return wb;
        }

        // ─── Date de referinta ANS (din fisierul profei, 27 feb 2026) ───────────
        // Folosit pentru a garanta ca fractiunile si gradele sunt IDENTICE cu fisierul oficial.
        // Cheia = UPPERCASE normalized (fara spatii duble, fara *).
        // Valoare = (grad exact, dict: ansId -> fractie)
        // ansId 1=Matematica, 2=Informatica, ... 40=Stiintele Sportului (conform col 10-49 din Excel)
        private static readonly Dictionary<string, (string Grad, Dictionary<int, decimal> Domenii)>
            AnsRef = new(StringComparer.OrdinalIgnoreCase)
        {
            { "ABAITANCEI HORIA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.99m} }) },
            { "ABRUDAN IOAN VASILE", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.43m}, {26, 0.45m} }) },
            { "ACIU LIA ELENA", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 1.01m} }) },
            { "ADAM MIHAI-SORIN", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.25m}, {9, 0.25m}, {12, 0.5m} }) },
            { "ADOCHIȚE (GĂLBĂU) CRISTINA-ȘTEFANIA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "AGACHE IOANA-OCTAVIA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 0.97m} }) },
            { "ALBU RUXANDRA GABRIELA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 0.17m}, {26, 0.83m} }) },
            { "ALDEA ADRIAN", ("Asist. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "ALDEA CODRUTA NICOLETA", ("Conf. dr.", new Dictionary<int,decimal>{ {1, 0.75m}, {11, 0.25m} }) },
            { "ALDEA CONSTANTIN LUCIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.5m}, {26, 0.5m} }) },
            { "ALECU STEFAN", ("Lect. dr.", new Dictionary<int,decimal>{ {9, 0.04m}, {15, 0.07m}, {21, 0.07m}, {22, 0.29m}, {25, 0.11m}, {26, 0.21m}, {40, 0.21m} }) },
            { "ALEXANDRESCU DANA SORINA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "ALEXANDRU CATALIN", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "ALEXANDRU MARIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.95m}, {11, 0.06m} }) },
            { "ALEXE RALUCA MONICA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.09m}, {12, 0.03m}, {19, 0.38m}, {28, 0.5m} }) },
            { "ANASTASIU ALEXANDRU-RAZVAN", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.4m}, {39, 0.58m} }) },
            { "ANASTASIU COSTIN VLAD", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "ANDREESCU OANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "ANDRONIC LUMINITA CAMELIA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.88m}, {26, 0.12m} }) },
            { "ANDRONIC MARIA LETITIA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "ANGHELINA BOGDAN-CRISTIAN", ("Asist.drd.", new Dictionary<int,decimal>{ {1, 0.27m}, {6, 0.6m}, {11, 0.13m} }) },
            { "ANTON CARMEN ELENA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "ANTONARU CARMEN ELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.13m}, {12, 0.09m}, {40, 0.78m} }) },
            { "ANTONYA CSABA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 1.0m} }) },
            { "APOSTOAIE MIRELA", ("Asist.drd.", new Dictionary<int,decimal>{ {12, 1.03m} }) },
            { "ARBANAŞ IOANA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "ARGĂSEALĂ GEORGIANA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 0.25m}, {28, 0.75m} }) },
            { "ARHIRE MONA BRIGITTE", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.14m}, {28, 0.8m}, {32, 0.06m} }) },
            { "ARMASELU ANCA", ("Asist. dr.", new Dictionary<int,decimal>{ {9, 0.03m}, {12, 0.51m}, {26, 0.47m} }) },
            { "ARMĂSAR IOANA PAULA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "ARON IOAN", ("Conf. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "ARVATESCU CRISTIAN", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "ATUDOREI IOANA ANISA", ("Conf. dr.", new Dictionary<int,decimal>{ {22, 1.0m} }) },
            { "BABA MARIUS NICOLAE", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.25m}, {12, 0.75m} }) },
            { "BABA MIRELA CAMELIA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "BADAU ADELA", ("Prof. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "BADAU DANA", ("Prof. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "BADEA ANAMARIA RALUCA", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 0.84m} }) },
            { "BADEA MIHAELA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "BADICU GEORGIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 0.5m}, {40, 0.5m} }) },
            { "BAICOIANU ALEXANDRA", ("Conf. dr.", new Dictionary<int,decimal>{ {2, 0.33m}, {25, 0.3m}, {26, 0.38m} }) },
            { "BALAN FLORIN", ("Lect. dr.", new Dictionary<int,decimal>{ {39, 1.01m} }) },
            { "BALAN TITUS CONSTANTIN", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "BALAS MONICA LOREDANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "BALASESCU MARIUS", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "BALASESCU SIMONA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "BALINT ELENA", ("Prof. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "BALINT LORAND", ("Prof. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "BALTES LIANA SANDA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.87m}, {26, 0.13m} }) },
            { "BALTESCU CODRUTA ADINA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "BARABAS BARNA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.53m}, {19, 0.22m} }) },
            { "BARACAN ADRIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "BARBU DANIELA MARIANA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "BARBU ION", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "BARBU MAGDALENA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "BARBU MARIUS CATALIN", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.99m} }) },
            { "BARBU SILVIU GABRIEL", ("Conf. dr.", new Dictionary<int,decimal>{ {19, 1.01m} }) },
            { "BARBULESCU ALINA", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "BARBULESCU OANA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.02m} }) },
            { "BAROTE LUMINITA", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.71m}, {9, 0.17m}, {11, 0.13m} }) },
            { "BASALIC ELENA-BIANCA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "BATRANU PINTEA VLAD", ("Lect. dr.", new Dictionary<int,decimal>{ {22, 0.32m}, {26, 0.68m} }) },
            { "BAZGAN MARIUS", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "BEDELEAN IOAN BOGDAN", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "BEDO TIBOR", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "BEGU TEODORA-MARIA", ("Asist.drd.", new Dictionary<int,decimal>{ {10, 0.7m}, {12, 0.3m} }) },
            { "BELDEAN EMANUELA CARMEN", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "BELDEAN NICOLAE LAURENTIU", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.14m}, {39, 0.86m} }) },
            { "BELDIANU IOLANDA FELICIA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "BELIBOU ALEXANDRA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.25m}, {39, 0.75m} }) },
            { "BENCZE ANDREI", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.75m}, {26, 0.25m} }) },
            { "BENEA BOGDAN CORNEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 1.0m} }) },
            { "BEȘCHEA ANDREI-GEORGE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "BIGIU NICUSOR FLORIN", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "BILDEA TEODOR STEFAN", ("Conf. dr.", new Dictionary<int,decimal>{ {25, 1.0m} }) },
            { "BISOC ALINA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "BOBESCU ELENA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "BOBOC RAZVAN GABRIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.57m}, {12, 0.43m} }) },
            { "BOCA LIANA LUMINITA", ("Asist.drd.", new Dictionary<int,decimal>{ {2, 0.74m}, {11, 0.2m}, {28, 0.06m} }) },
            { "BOCU RAZVAN", ("Conf. dr.", new Dictionary<int,decimal>{ {2, 0.67m}, {26, 0.33m} }) },
            { "BODI DIANA CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {22, 0.88m} }) },
            { "BODOC ALICE MAGDALENA", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 0.92m}, {32, 0.07m} }) },
            { "BOER ATTILA LASZLO", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "BOGATU CRISTINA AURICA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "BOGDAN IOANA CORINA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.86m}, {12, 0.14m} }) },
            { "BOLBORICI ANA MARIA", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.17m}, {22, 0.84m} }) },
            { "BOLDISOR CRISTIAN NICOLAE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {11, 1.0m} }) },
            { "BOLOCAN SORIN IONUT", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "BONDOC IONESCU ALEXANDRU", ("Lect. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "BORCAN VIRGIL", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.14m}, {28, 0.88m} }) },
            { "BORCOMAN MARIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {22, 0.29m}, {26, 0.42m}, {27, 0.33m} }) },
            { "BORZ STELIAN ALEXANDRU", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.36m}, {26, 0.64m} }) },
            { "BOSCOIANU MIRCEA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.27m}, {12, 0.73m} }) },
            { "BOSCOR DANA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "BOTA OANA ALINA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "BOTESCU-SIREŢEANU ILEANA-AURORA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.21m}, {32, 0.79m} }) },
            { "BOTEZATU DAN GEORGE", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.19m}, {28, 0.83m} }) },
            { "BOTIANU ANA MARIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.88m} }) },
            { "BOTIS MARIUS FLORIN", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "BOTIS SORINA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "BRANEA (ȚACĂ) IOANA - ANTONIA", ("Asist.drd.", new Dictionary<int,decimal>{ {7, 0.07m}, {9, 0.27m}, {11, 0.33m}, {12, 0.27m}, {26, 0.07m} }) },
            { "BRATU CIPRIAN", ("Asist.drd.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "BRATU CONSTANTIN ALEXANDRU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "BRATU DRAGOS-VASILE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {11, 0.79m}, {12, 0.21m} }) },
            { "BRATU MARIA-ALEXANDRA", ("Asist.drd.", new Dictionary<int,decimal>{ {2, 0.1m}, {11, 0.7m}, {12, 0.2m} }) },
            { "BRATUCU GABRIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "BRAUN BARBU CRISTIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "BRENCI LUMINITA MARIA", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.99m} }) },
            { "BREZEANU ALIN IONUȚ", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.03m} }) },
            { "BRICIU GABRIELA ARABELA", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.33m}, {22, 0.33m}, {26, 0.33m} }) },
            { "BRICIU VICTOR ALEXANDRU", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.56m}, {22, 0.13m}, {26, 0.33m} }) },
            { "BRĂNESCU GERONIMO-RĂDUCU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.96m}, {12, 0.04m} }) },
            { "BUCS LORANT", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.36m}, {26, 0.64m} }) },
            { "BUCUR ROMULUS LADISLAU", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.21m}, {28, 0.82m} }) },
            { "BUDALA ADRIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 1.02m} }) },
            { "BUGA CRISTINA MARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.46m}, {39, 0.53m} }) },
            { "BUHAICIUC MIHAELA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.25m}, {39, 0.75m} }) },
            { "BUICAN GEORGE RAZVAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.88m}, {12, 0.14m} }) },
            { "BUJA ELENA", ("Prof. dr.", new Dictionary<int,decimal>{ {28, 1.03m} }) },
            { "BULARCA ANCA ROXANA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.01m} }) },
            { "BULARCA MARIA-CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {21, 0.21m}, {22, 0.35m}, {26, 0.43m} }) },
            { "BULARCA RAZVAN", ("Asist. dr.", new Dictionary<int,decimal>{ {38, 0.22m}, {39, 0.78m} }) },
            { "BULMEZ ALEXANDRU MIHAI", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.01m} }) },
            { "BURADA MARINELA", ("Prof. dr.", new Dictionary<int,decimal>{ {28, 1.01m} }) },
            { "BURBEA GEORGIANA-MIHAELA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.04m}, {28, 0.98m} }) },
            { "BURDUHOS BOGDAN GABRIEL", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "BURLACU MIHAI", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.54m}, {22, 0.06m}, {26, 0.42m} }) },
            { "BUSUIOCEANU STELIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {25, 0.17m}, {26, 0.84m} }) },
            { "BUTNARIU SILVIU LUIS", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.72m}, {12, 0.27m} }) },
            { "BUVNARIU LAVINIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.78m} }) },
            { "BUZDUGAN IOANA DIANA", ("Asist.drd.", new Dictionary<int,decimal>{ {9, 0.73m}, {12, 0.27m} }) },
            { "BUZEA CARMEN", ("Prof. dr.", new Dictionary<int,decimal>{ {21, 0.18m}, {22, 0.84m} }) },
            { "BÂRSAN MARIA IONELA", ("Asist. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "BÂRSAN MARIA MAGDALENA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "BĂDĂRĂU CARMEN LILIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 1.01m} }) },
            { "CALIN MARIUS DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.65m}, {9, 0.11m}, {12, 0.25m} }) },
            { "CAMPEAN MIHAELA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "CAMPU ADINA", ("Lect. dr.", new Dictionary<int,decimal>{ {21, 0.17m}, {22, 0.33m}, {26, 0.5m} }) },
            { "CAMPU VASILE RAZVAN", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.91m}, {26, 0.09m} }) },
            { "CANDREA ADINA NICOLETA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "CANJA CRISTINA MARIA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.91m}, {12, 0.09m} }) },
            { "CARP MARIUS CATALIN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.5m}, {11, 0.5m} }) },
            { "CATANA DORIN IOAN", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "CATANESCU ANDREEA CORINA", ("Lect. dr.", new Dictionary<int,decimal>{ {6, 0.03m}, {7, 0.03m}, {40, 0.94m} }) },
            { "CATARON ANGEL DORU", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 1.01m} }) },
            { "CATEANU MIHNEA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {8, 0.35m}, {26, 0.65m} }) },
            { "CAZACU CHRISTIANA EMILIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "CAZAN ANA MARIA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 0.06m}, {27, 0.93m} }) },
            { "CAZAN CRISTINA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.18m}, {12, 0.82m} }) },
            { "CERBU CAMELIA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.21m}, {12, 0.79m} }) },
            { "CERNEA NICOLETA", ("Lect. dr.", new Dictionary<int,decimal>{ {9, 0.47m}, {12, 0.14m}, {21, 0.15m}, {22, 0.23m} }) },
            { "CHEFNEUX GABRIELA", ("Prof. dr.", new Dictionary<int,decimal>{ {28, 0.99m}, {32, 0.05m} }) },
            { "CHELMEA LIGIA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "CHEȘCĂ ANTONELLA ELISA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "CHICOMBAN CARMEN MIHAELA", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "CHICOS LUCIA ANTONETA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 0.9m}, {25, 0.09m} }) },
            { "CHIHAIA GABRIELA-NICOLETA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 0.99m} }) },
            { "CHIRCAN ELIZA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 0.11m}, {9, 0.65m}, {12, 0.25m} }) },
            { "CHIRILA ADINA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.29m}, {7, 0.21m}, {8, 0.14m}, {11, 0.14m}, {12, 0.21m} }) },
            { "CHIS ALEXANDRU", ("Asist.drd.", new Dictionary<int,decimal>{ {7, 0.87m}, {11, 0.14m} }) },
            { "CHISALITA DUMITRU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "CHITONU GABRIELA CRISTINA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "CHITU IOANA BIANCA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "CHIVU CATALIN IULIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 0.72m}, {26, 0.28m} }) },
            { "CHIVU CATRINA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "CIOARA GHEORGHE ROMEO", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "CIOBANU CATALIN", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.58m}, {11, 0.42m} }) },
            { "CIOBANU DANIELA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "CIOBANU ELIZA", ("Lect. dr.", new Dictionary<int,decimal>{ {25, 0.14m}, {26, 0.85m} }) },
            { "CIOBANU RAMONA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "CIOCIRLAN ELENA", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.43m}, {26, 0.56m} }) },
            { "CIOLOCA ANASTASIA MALINA", ("Asist. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "CIOPLEIAS BOGDAN-NICOLAE", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "CIOROIU SILVIU GABRIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "CIRSTOLOVEAN IOAN LUCIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "CISMARU LAURA", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.42m}, {12, 0.58m} }) },
            { "CIUPALA LAURA ANCA", ("Conf. dr.", new Dictionary<int,decimal>{ {2, 0.59m}, {7, 0.25m}, {26, 0.17m} }) },
            { "CIUREA ANDREEA CĂTĂLINA", ("Conf. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "CIUREA CODRUT IOAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 0.07m}, {15, 0.92m} }) },
            { "CIURESCU DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "CLINCIU MIHAELA RODICA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.14m}, {12, 0.5m}, {26, 0.36m} }) },
            { "CLINCIU RAMONA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.73m}, {12, 0.27m} }) },
            { "CLOTEA LUMINITA ROXANA", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 1.01m} }) },
            { "COBELSCHI CĂLIN PAVEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "COCIAS TIBERIU", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "COCUZ MARIA ELENA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "CODREAN CODRIN LEONID", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.8m}, {26, 0.2m} }) },
            { "COLIBAN RADU MIHAI", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "COMAN ALINA", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.42m}, {22, 0.33m}, {26, 0.25m} }) },
            { "COMAN CLAUDIU", ("Prof. dr.", new Dictionary<int,decimal>{ {22, 1.0m} }) },
            { "COMAN ECATERINA", ("Lect. dr.", new Dictionary<int,decimal>{ {25, 0.18m}, {26, 0.82m} }) },
            { "COMAN SIMONA", ("Conf. dr.", new Dictionary<int,decimal>{ {11, 1.0m} }) },
            { "COMSIT MIHAI", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "COMĂNESCU IOANA SONIA", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "CONDREA EMILIA-GABRIELA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "CONSTANTIN BOGDAN", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.72m}, {39, 0.28m} }) },
            { "CONSTANTIN CRISTINEL PETRISOR", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "CONSTANTIN DAN ALEXANDRU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "CONSTANTIN SANDA", ("Lect. dr.", new Dictionary<int,decimal>{ {25, 0.35m}, {26, 0.67m} }) },
            { "CONSTANTINESCU CRISTIAN ADRIAN", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 0.88m} }) },
            { "CONSTANTINESCU ELENA MIHAELA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.89m} }) },
            { "CONTIU MIRCEA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "CORA IRINGO", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.27m}, {12, 0.09m}, {22, 0.3m}, {26, 0.2m}, {39, 0.14m} }) },
            { "COROIU PETRUTA MARIA", ("Prof. dr.", new Dictionary<int,decimal>{ {38, 0.72m}, {39, 0.27m} }) },
            { "COSEREANU CAMELIA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "COSTACHE CRISTEA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.85m} }) },
            { "COSTACHE DELIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.79m} }) },
            { "COSTIUC IULIANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 1.01m} }) },
            { "COSTIUC LIVIU", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.36m}, {12, 0.64m} }) },
            { "COTARLEA DELIA ANCA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.31m}, {28, 0.71m} }) },
            { "COTFAS DANIEL TUDOR", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 0.18m}, {7, 0.81m} }) },
            { "COTFAS PETRU ADRIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.81m}, {11, 0.18m} }) },
            { "COVEI MARIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.14m}, {12, 0.86m} }) },
            { "CRACIUN ADRIAN VIRGIL", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "CRETESCU NADIA RAMONA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "CRISBĂȘAN ANDREEA-MARIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "CRISTEA DANIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "CRISTEA LUCIANA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "CROITORU CATALIN", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "CSESZNEK CODRINA", ("Prof. dr.", new Dictionary<int,decimal>{ {22, 0.99m} }) },
            { "CUCULEA DAN-CRISTIAN", ("Asist.drd.", new Dictionary<int,decimal>{ {9, 0.07m}, {12, 0.93m} }) },
            { "CURTU ALEXANDRU LUCIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.27m}, {26, 0.68m} }) },
            { "CUSEN GABRIELA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.14m}, {27, 0.14m}, {28, 0.71m} }) },
            { "CÂMPEAN STEFAN-IOAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "CĂLIN (COMȘIȚ) ANDREEA-MIHAELA", ("Asist.drd.", new Dictionary<int,decimal>{ {7, 0.47m}, {11, 0.53m} }) },
            { "DAMŞESCU ADRIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {22, 0.13m}, {26, 0.88m} }) },
            { "DANCIU GABRIEL MIHAIL", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.36m}, {11, 0.64m} }) },
            { "DANILA ADRIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.46m}, {11, 0.54m} }) },
            { "DAVID LAURA TEODORA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "DEACONESCU ANDREA CATALINA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "DEACONESCU TUDOR ION", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "DEACONU ADRIAN MARIUS", ("Prof. dr.", new Dictionary<int,decimal>{ {2, 0.81m}, {26, 0.18m} }) },
            { "DEACONU OVIDIU", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "DEAKY BOGDAN ALEXANDRU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.14m}, {12, 0.84m} }) },
            { "DEMETER ROBERT", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.14m}, {11, 0.86m} }) },
            { "DERCZENI RUDOLF ALEXANDRU", ("Conf. dr.", new Dictionary<int,decimal>{ {8, 0.33m}, {26, 0.67m} }) },
            { "DIACONU IOANA ANDREA", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 1.0m} }) },
            { "DIACONU LAURENTIU IONEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {11, 0.81m}, {12, 0.19m} }) },
            { "DIACONU STEFANIA-ROXANA", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "DIMA DRAGOS SORIN", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.99m} }) },
            { "DIMA GABRIELA", ("Prof. dr.", new Dictionary<int,decimal>{ {22, 1.0m} }) },
            { "DIMA LORENA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "DIMIENESCU OANA GABRIELA", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "DIMITRIU MARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.57m}, {7, 0.32m}, {12, 0.11m} }) },
            { "DIMULESCU CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 0.76m}, {32, 0.26m} }) },
            { "DINCA GHEORGHITA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "DINCA MARIUS SORIN", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "DINU ALEXANDRU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.61m}, {11, 0.39m} }) },
            { "DINU CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 0.99m} }) },
            { "DINU CĂTĂLINA GEORGETA", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 0.06m}, {19, 0.94m} }) },
            { "DINU ELEONORA ANTOANETA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "DINULICA FLORIN", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.3m}, {26, 0.7m} }) },
            { "DOBRESCU ADA IOANA", ("Conf. dr.", new Dictionary<int,decimal>{ {22, 1.02m} }) },
            { "DRACEA LAURA LARISA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "DRAGHICI CAMELIA LUCIA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "DRAGOI MIRCEA VIOREL", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.26m}, {12, 0.71m}, {26, 0.03m} }) },
            { "DRAGOMIR GEORGE", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "DRAGOMIR PÂNZARU CAMELIA CRISTINA", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.19m}, {25, 0.45m}, {26, 0.35m} }) },
            { "DRUGA CORNELIU NICOLAE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "DRUGĂU SORIN", ("Lect. dr.", new Dictionary<int,decimal>{ {12, 0.14m}, {26, 0.14m}, {40, 0.71m} }) },
            { "DRUMEA CRISTINA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "DUCA LILIANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "DUGULEANA CONSTANTIN", ("Conf. dr.", new Dictionary<int,decimal>{ {25, 0.25m}, {26, 0.76m} }) },
            { "DUGULEANA LILIANA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "DUGULEANA MIHAI", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.63m}, {21, 0.36m} }) },
            { "DUICU SIMONA SOFIA", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.49m}, {12, 0.5m} }) },
            { "DUMITRASCU ADELA-ELIZA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "DUMITRASCU DORIN ION", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.95m}, {12, 0.06m} }) },
            { "DUMITRESCU FLORIN", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 0.4m}, {32, 0.02m}, {38, 0.21m}, {39, 0.35m} }) },
            { "DUMITRESCU SILVIU RAZVAN", ("Conf. dr.", new Dictionary<int,decimal>{ {2, 0.68m}, {7, 0.18m}, {26, 0.14m} }) },
            { "DUTCA IOAN", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.67m}, {26, 0.33m} }) },
            { "DĂNILĂ DANIEL MIHAI", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.82m}, {12, 0.13m}, {26, 0.07m} }) },
            { "EFTIMIE NICOLAE", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "ELEKES ROBERT GABRIEL", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 1.0m} }) },
            { "ENACHE DORIN VALTER", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.67m}, {26, 0.34m} }) },
            { "ENACHE-DAVID NICOLETA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.43m}, {25, 0.29m}, {26, 0.29m} }) },
            { "ENE ANA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.17m}, {28, 0.83m} }) },
            { "ENESCA IOAN ALEXANDRU", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.09m}, {12, 0.9m} }) },
            { "ENESCU ADRIAN-GABRIEL", ("Asist.drd.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "ENESCU IOANA-CLARA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.21m}, {27, 0.29m}, {28, 0.5m} }) },
            { "ENESCU RALUCA ELENA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {8, 0.25m}, {10, 0.07m}, {26, 0.68m} }) },
            { "ENOIU RAZVAN SANDU", ("Prof. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "FALUP PECURARIU CRISTIAN GAVRIL", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "FALUP PECURARIU OANA GABRIELA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "FECHETE FLAVIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.21m}, {11, 0.29m}, {12, 0.48m}, {26, 0.02m} }) },
            { "FELEA ALINA SILVANA", ("Conf. dr.", new Dictionary<int,decimal>{ {28, 1.02m} }) },
            { "FILIP ALEXANDRU CATALIN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.03m} }) },
            { "FILIP IGNAC - CSABA", ("Prof. dr.", new Dictionary<int,decimal>{ {38, 0.45m}, {39, 0.54m} }) },
            { "FILIP OVIDIU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "FIRASTRAU IOANA", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.33m}, {9, 0.41m}, {12, 0.08m}, {26, 0.17m} }) },
            { "FLOREA OLIVIA ANA", ("Conf. dr.", new Dictionary<int,decimal>{ {1, 0.61m}, {9, 0.06m}, {12, 0.34m} }) },
            { "FLORESCU ADRIANA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "FLORESCU MONICA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.55m}, {15, 0.45m} }) },
            { "FLOROIAN LAURA", ("Conf. dr.", new Dictionary<int,decimal>{ {11, 0.58m}, {12, 0.42m} }) },
            { "FOLEA MILENA FLAVIA", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.06m}, {12, 0.94m} }) },
            { "FORIS DIANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.14m}, {12, 0.87m} }) },
            { "FORIS TIBERIU", ("Prof. dr.", new Dictionary<int,decimal>{ {25, 0.23m}, {26, 0.78m} }) },
            { "FRATU MARIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "FRIEDL ANNAMARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.5m}, {11, 0.07m}, {12, 0.28m}, {26, 0.14m} }) },
            { "FRINCU MADALINA ILEANA", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "FUGARETU COSMINA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "FULGA ANDREEA ILEANA", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 0.09m}, {7, 0.17m}, {11, 0.25m}, {12, 0.17m}, {25, 0.17m}, {26, 0.17m} }) },
            { "FÎNTÎNĂ IOANA MARIA", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 0.99m} }) },
            { "GABOR CAMELIA", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.06m}, {12, 0.96m} }) },
            { "GACEU LIVIU", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.86m}, {12, 0.14m} }) },
            { "GALATANU TEOFIL FLORIN", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "GALMEANU HONORIUS CEZAR", ("Conf. dr.", new Dictionary<int,decimal>{ {2, 1.02m} }) },
            { "GAROIU STEFAN LUCIAN", ("Asist.drd.", new Dictionary<int,decimal>{ {7, 0.13m}, {9, 0.13m}, {12, 0.6m}, {26, 0.13m} }) },
            { "GAVRILA CORNEL CATALIN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "GAVRIS CLAUDIA MIHAELA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "GAVRUS CRISTINA", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.04m}, {12, 0.8m}, {26, 0.17m} }) },
            { "GHEORGHE CARMEN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.97m} }) },
            { "GHEORGHE CARMEN ADRIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "GHEORGHE CATALIN", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "GHEORGHE DANA MIHAELA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.04m}, {28, 0.96m} }) },
            { "GHEORGHE VASILE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.07m}, {12, 0.93m} }) },
            { "GHEORGHITA (LICHIOIU) IULIANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.02m} }) },
            { "GHIGHECI COSTEL CRISTINEL", ("Conf. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "GHITA DANA ELENA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "GHIŢĂ-PÎRNUŢĂ OANA-ANDREEA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.29m}, {32, 0.72m} }) },
            { "GINERICA COSMIN", ("Asist. dr.", new Dictionary<int,decimal>{ {11, 0.6m}, {12, 0.4m} }) },
            { "GIRBACIA FLORIN STELIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.9m}, {12, 0.09m} }) },
            { "GIRDAN LAURA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "GLIGA CONSTANTIN IOAN", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "GOTEA MIHAELA", ("Conf. dr.", new Dictionary<int,decimal>{ {22, 1.01m} }) },
            { "GREŞITĂ CONSTANTIN IRINEL", ("Conf. dr.", new Dictionary<int,decimal>{ {8, 1.0m} }) },
            { "GRIGORESCU OVIDIU DAN", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "GRIGORESCU SIMONA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "GRIGORESCU SORIN MIHAI", ("Prof. dr.", new Dictionary<int,decimal>{ {11, 0.54m}, {12, 0.45m} }) },
            { "GROSZ WILHELM ROBERT", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "GUIMAN MARIA VIOLETA", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.25m}, {12, 0.38m}, {26, 0.37m} }) },
            { "GURAU LIDIA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "GUREAN DAN MARIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "HABA SEVER", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "HALALISAN AURELIU FLORIN", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.27m}, {26, 0.72m} }) },
            { "HENTER RAMONA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.01m} }) },
            { "HLIPCĂ PETRU", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "HOGEA MIRCEA DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "HUMINIC GABRIELA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.31m}, {12, 0.69m} }) },
            { "HUMINIC TRAIAN ANGEL", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.64m}, {12, 0.36m} }) },
            { "IACOB ANDREEA-BIANCA", ("Asist.drd.", new Dictionary<int,decimal>{ {1, 0.2m}, {2, 0.2m}, {26, 0.6m} }) },
            { "IBANESCU DANIELA CORINA", ("Prof. dr.", new Dictionary<int,decimal>{ {38, 1.0m} }) },
            { "ICHIM TRAIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {39, 1.01m} }) },
            { "IDOMIR MIHAELA ELENA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "IFTENE LIVIU", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.43m}, {39, 0.57m} }) },
            { "IFTENI PETRU IULIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "IGNAT MIHAI", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.23m}, {28, 0.77m} }) },
            { "ILEA ANCA-MARIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "ILIE RODICA MARIA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.09m}, {28, 0.9m} }) },
            { "INDREICA ELENA SIMONA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.01m} }) },
            { "INDREICA VICTOR ADRIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {8, 0.27m}, {10, 0.23m}, {26, 0.5m} }) },
            { "ION CATALIN PETREA", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.99m} }) },
            { "ION LAURENTIU-MIHAIL", ("Lect. dr.", new Dictionary<int,decimal>{ {9, 0.15m}, {12, 0.67m}, {26, 0.19m} }) },
            { "IONAŞ DIANA GEANINA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "IONESCU ALEXANDRU CODRIN", ("Lect. dr.", new Dictionary<int,decimal>{ {6, 0.5m}, {7, 0.14m}, {9, 0.14m}, {11, 0.21m} }) },
            { "IONESCU ANA MARIA", ("Asist. dr.", new Dictionary<int,decimal>{ {12, 0.5m}, {26, 0.5m} }) },
            { "IONESCU DAN TRAIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "IONESCU OVIDIU", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.48m}, {26, 0.52m} }) },
            { "IORDACHE DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.87m}, {26, 0.13m} }) },
            { "IORDACHE EUGEN", ("Conf. dr.", new Dictionary<int,decimal>{ {8, 0.06m}, {10, 0.12m}, {26, 0.81m} }) },
            { "IORDAN NICOLAE FANI", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.02m} }) },
            { "IOVANAS DANIELA MARIA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "IRIMIE CLAUDIA-ALEXANDRINA", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 0.97m} }) },
            { "IRIMIE IOANA VIOLETA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 0.14m}, {28, 0.86m} }) },
            { "IRIMIE MARIUS", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "ISAC IULIANA", ("Asist. dr.", new Dictionary<int,decimal>{ {27, 0.2m}, {38, 0.2m}, {39, 0.61m} }) },
            { "ISAC LUMINITA ANISOARA", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 0.18m}, {12, 0.81m} }) },
            { "ISAIA FLORIN", ("Conf. dr.", new Dictionary<int,decimal>{ {1, 0.58m}, {9, 0.29m}, {12, 0.13m} }) },
            { "ISAIA GABRIELA AURORA", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.59m}, {26, 0.42m} }) },
            { "ISBASOIU ANDREEA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 0.99m} }) },
            { "ISOP LAURA-MIHAELA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 0.96m} }) },
            { "ISPAS ANA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "ISPAS MIHAI", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "ISPAS NICOLAE", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.99m} }) },
            { "ITU ALINA", ("Conf. dr.", new Dictionary<int,decimal>{ {11, 0.98m} }) },
            { "ITU CALIN", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.05m}, {12, 0.95m} }) },
            { "ITU LUCIAN MIHAI", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.09m}, {11, 0.81m}, {12, 0.09m} }) },
            { "IVANCESCU RUXANDRA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.16m}, {28, 0.83m} }) },
            { "IVANOVICI LAURENTIU MIHAI", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.93m}, {11, 0.09m} }) },
            { "IVASCIUC IOANA SIMONA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "IVASCU IRINA MIHAELA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.21m}, {28, 0.79m} }) },
            { "JALIU CODRUTA ILEANA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "KAKUCS CRISTIAN", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "KARACSONY NOEMI", ("Conf. dr.", new Dictionary<int,decimal>{ {39, 1.0m} }) },
            { "KERTESZ CSABA ZOLTAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.99m} }) },
            { "KOLAR VASUDEVA LAURA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "KOVACS ATTILA", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.5m}, {22, 0.17m}, {26, 0.33m} }) },
            { "KRISTALY DOMINIC MIRCEA", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.23m}, {11, 0.5m}, {12, 0.27m} }) },
            { "LACATUS ADRIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.07m}, {28, 0.94m} }) },
            { "LACHE SIMONA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.72m}, {12, 0.23m} }) },
            { "LACULICEANU ALEXANDRU-GEORGIAN", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "LANCEA CAMIL TRAIAN SORIN", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "LAPTES RAMONA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "LATES MIHAI TIBERIU", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.18m}, {12, 0.83m} }) },
            { "LAZAR ANAMARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.85m}, {12, 0.14m} }) },
            { "LAZAR CORNELIA MAGDALENA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "LEAHU CRISTIAN IOAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 1.01m} }) },
            { "LEASU FLORIN GABRIEL", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "LELUTIU LAURA MIHAELA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "LIMBASAN ILEANA GEORGIANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "LINDEMANN SOFIANA IULIA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.04m}, {28, 0.96m} }) },
            { "LITRA ADRIANA VERONICA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "LIXANDROIU RADU CONSTANTIN", ("Prof. dr.", new Dictionary<int,decimal>{ {25, 0.54m}, {26, 0.45m} }) },
            { "LORINCZ SIMINA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "LOSTUN ALEXANDRA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.88m} }) },
            { "LUCA MIHAI ALEXANDRU", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "LUCULESCU MARIUS CRISTIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "LUNGOCI CARMEN MIHAELA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "LUNGU ANTONELA CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "LUNGULEASA AUREL", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "LUPSA TATARU DANA ADRIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "LUPSA TATARU LUCIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 1.01m} }) },
            { "LUPU DACIANA ANGELICA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "LUPU DRAGOȘ", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "LUPU MIRABELA IOANA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.91m}, {12, 0.09m} }) },
            { "LUPU NICOLETA RALUCA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "LĂCĂTUȘ ANCA MARIA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MACESANU GIGEL", ("Conf. dr.", new Dictionary<int,decimal>{ {11, 0.41m}, {12, 0.59m} }) },
            { "MACHEDON PISU MIHAI", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.67m}, {11, 0.08m}, {12, 0.25m} }) },
            { "MAFTEI CARMEN", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "MAICAN CATALIN IOAN", ("Prof. dr.", new Dictionary<int,decimal>{ {25, 0.67m}, {26, 0.34m} }) },
            { "MAICAN MARIA ANCA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "MAIER ALINA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.86m}, {12, 0.14m} }) },
            { "MAJERCSIK LUCIANA", ("Lect. dr.", new Dictionary<int,decimal>{ {1, 0.29m}, {2, 0.71m} }) },
            { "MANCIULEA ILEANA CARMEN", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 0.83m}, {26, 0.17m} }) },
            { "MANDRU LIDIA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "MANEA ADELINA LOREDANA", ("Lect. dr.", new Dictionary<int,decimal>{ {1, 0.29m}, {2, 0.71m} }) },
            { "MANEA ELENA LAURA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.1m}, {28, 0.88m} }) },
            { "MANEA EMILIA ADELA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.99m} }) },
            { "MANEA ROSANA MIHAELA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MANOLICĂ ANA-MARIA", ("Asist.drd.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "MANTULESCU MARIUS MIHAIL", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "MARCEANU LUIGI GEO", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MARCU MARINA VIORELA", ("Conf. dr.", new Dictionary<int,decimal>{ {8, 0.75m}, {10, 0.16m}, {12, 0.08m} }) },
            { "MARDACHE ANDREEA CLAUDIA", ("Lect. dr.", new Dictionary<int,decimal>{ {21, 0.06m}, {22, 0.96m} }) },
            { "MARINESCU DANIELA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.87m} }) },
            { "MARINESCU NICOLAE ION", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "MARTOMA ALINA MIRELA", ("Lect. dr.", new Dictionary<int,decimal>{ {40, 1.01m} }) },
            { "MATEFI ROXANA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 0.08m}, {19, 0.92m} }) },
            { "MATEI ALEXANDRU", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.42m}, {28, 0.58m} }) },
            { "MATEI FLORENTINA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.7m}, {26, 0.3m} }) },
            { "MATEI MADALINA GEORGIANA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.64m}, {12, 0.22m}, {19, 0.14m} }) },
            { "MAZAREL ADRIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.89m}, {12, 0.11m} }) },
            { "MESESAN SCHMITZ LUIZA IULIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.34m}, {22, 0.68m} }) },
            { "MICLAUS STELIANA ROXANA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MICU CORINA SILVIA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.25m}, {28, 0.74m} }) },
            { "MICULESCU RADU", ("Prof. dr.", new Dictionary<int,decimal>{ {1, 1.0m} }) },
            { "MIHAIL LAURENTIU AUREL", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "MIHAILESCU MARIA-MIRABELA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MIHAILESCU TEOFIL", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "MIHALCICA MIRCEA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "MIJAICA RALUCA DACIA", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "MILESAN MIHAELA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.19m}, {9, 0.25m}, {11, 0.06m}, {12, 0.5m} }) },
            { "MILOSAN IOAN", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.02m} }) },
            { "MINCULETE NICUSOR", ("Conf. dr.", new Dictionary<int,decimal>{ {1, 0.67m}, {2, 0.17m}, {26, 0.16m} }) },
            { "MINDRESCU VERONICA", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 0.09m}, {40, 0.91m} }) },
            { "MIRON ( MIOC ) ANA-ALIANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MISARCA CATALIN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MITREA NICOLETA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "MITRICA MARIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MITU LEONARD", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.02m} }) },
            { "MITU SEBASTIAN-RĂZVAN", ("Asist.drd.", new Dictionary<int,decimal>{ {26, 0.3m}, {28, 0.73m} }) },
            { "MIZGACIU CAMELIA", ("Lect. dr.", new Dictionary<int,decimal>{ {25, 0.14m}, {26, 0.86m} }) },
            { "MOARCĂS GEORGETA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.22m}, {28, 0.67m}, {32, 0.13m} }) },
            { "MOASA HORIA", ("Conf. dr.", new Dictionary<int,decimal>{ {21, 0.42m}, {22, 0.6m} }) },
            { "MODRAN HORIA ALEXANDRU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.91m}, {39, 0.09m} }) },
            { "MOGA MARIUS ALEXANDRU", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MOJA ADELINA - IOANA", ("Asist.drd.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "MOLDOVAN (TANTAU) MARA-STEFANIA", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MOLDOVAN EDIT ROXANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.21m}, {12, 0.78m} }) },
            { "MOLDOVAN MACEDON DUMITRU", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.04m} }) },
            { "MONESCU VLAD", ("Lect. dr.", new Dictionary<int,decimal>{ {9, 0.18m}, {25, 0.21m}, {26, 0.61m} }) },
            { "MORARIU CRISTIN OLIMPIU", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "MORARU SORIN AUREL", ("Prof. dr.", new Dictionary<int,decimal>{ {11, 0.91m}, {12, 0.09m} }) },
            { "MOSOI ADRIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "MOSOIU DANIELA VIORICA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MOTOASCA SEPTIMIU DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "MOTOC DANA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 1.0m} }) },
            { "MUNTEAN LIVIU-IULIU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "MUNTEAN RADU MIRCEA", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 0.99m} }) },
            { "MUNTEANU DANIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.11m}, {12, 0.9m} }) },
            { "MUNTEANU MIHAELA VIOLETA", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.71m}, {12, 0.17m}, {26, 0.13m} }) },
            { "MUNTEANU-ICHIM ROXANA ANDREEA", ("Asist.drd.", new Dictionary<int,decimal>{ {10, 0.87m}, {12, 0.13m} }) },
            { "MURESAN VALENTIN", ("Asist. dr.", new Dictionary<int,decimal>{ {38, 0.67m}, {39, 0.33m} }) },
            { "MUSAT ELENA CAMELIA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "MUSUROI CRISTIAN LEONARD", ("Asist. dr.", new Dictionary<int,decimal>{ {7, 0.37m}, {9, 0.14m}, {12, 0.51m} }) },
            { "MĂDA STANCA", ("Conf. dr.", new Dictionary<int,decimal>{ {28, 1.02m} }) },
            { "NANAU CORINA STEFANIA", ("Lect. dr.", new Dictionary<int,decimal>{ {1, 0.14m}, {2, 0.11m}, {26, 0.75m} }) },
            { "NASTAC DORIN CRISTIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 0.99m} }) },
            { "NASTASA LAURA ELENA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.01m} }) },
            { "NASTASE GABRIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "NASTASOIU MIRCEA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 1.01m} }) },
            { "NASULEA MARIUS DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.52m}, {12, 0.49m} }) },
            { "NAUNCEF ALINA MARIA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.66m}, {39, 0.33m} }) },
            { "NEACSU NICOLETA ANDREEA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "NEAGOE MIRCEA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "NEAGU MIRCEA", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 0.67m}, {9, 0.08m}, {12, 0.24m} }) },
            { "NECHIFOR BIANCA ANDREEA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.55m}, {28, 0.46m} }) },
            { "NECHITA FLORENTINA", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 1.01m} }) },
            { "NECHITA FLORIN MIHAI", ("Prof. dr.", new Dictionary<int,decimal>{ {21, 0.27m}, {22, 0.45m}, {26, 0.27m} }) },
            { "NECSOI DANIELA VERONICA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.01m} }) },
            { "NECULA RADU DAN", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "NECULA VALENTIN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.77m}, {26, 0.22m} }) },
            { "NECULAU ANDREA ELENA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "NECULOIU DANIELA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "NECULOIU MARIUS", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.88m} }) },
            { "NEDELOIU TIBERIU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "NEGULESCU ORIANA HELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {25, 0.14m}, {26, 0.86m} }) },
            { "NEPOTU GABRIEL LUCIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {1, 0.06m}, {7, 0.03m}, {11, 0.11m}, {12, 0.11m}, {25, 0.36m}, {26, 0.33m} }) },
            { "NICOLAE IOANA", ("Prof. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "NICOLAU ANDRADA CAMELIA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.88m} }) },
            { "NICOLAU LIANA CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "NICOLESCU VALERIU NOROCEL", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.77m}, {26, 0.23m} }) },
            { "NICULA DAN", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.25m}, {11, 0.35m}, {26, 0.38m} }) },
            { "NISTOR-ȘERBAN ANDREEA ELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {1, 0.35m}, {26, 0.64m} }) },
            { "NITA MIHAI DANIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {8, 0.18m}, {10, 0.44m}, {26, 0.38m} }) },
            { "NIȚOIU LORENA GABRIELA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "NUTU MARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.55m}, {11, 0.18m}, {26, 0.29m} }) },
            { "OANA ALEXANDRU", ("Lect. dr.", new Dictionary<int,decimal>{ {7, 0.11m}, {11, 0.36m}, {12, 0.37m}, {26, 0.17m} }) },
            { "OANCEA BOGDAN MARIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 1.01m} }) },
            { "OANCEA GHEORGHE", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "OGREZEANU IULIAN ALEXANDRU", ("Asist. dr.", new Dictionary<int,decimal>{ {7, 0.12m}, {11, 0.87m} }) },
            { "OGRUTAN PETRE LUCIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.64m}, {11, 0.36m} }) },
            { "OLA DANIEL CALIN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.51m}, {12, 0.51m} }) },
            { "OLAH ARTHUR", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "OLARESCU ALIN", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 1.01m} }) },
            { "OLTEANU MIRCEA IONUȚ", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.14m}, {40, 0.86m} }) },
            { "ONEA GHEORGHE ADRIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {9, 0.28m}, {12, 0.04m}, {25, 0.07m}, {26, 0.08m}, {39, 0.07m}, {40, 0.47m} }) },
            { "OPRISESCU SERBAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.21m}, {11, 0.79m} }) },
            { "ORMENISAN ALEXE NICOLAE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "PACURAR CRISTINA MARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {12, 0.29m}, {26, 0.72m} }) },
            { "PACURAR VICTOR DAN", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.42m}, {26, 0.58m} }) },
            { "PANAITE MARA", ("Asist. dr.", new Dictionary<int,decimal>{ {21, 0.49m}, {22, 0.53m} }) },
            { "PANTEA ILEANA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "PARV AURICA LUMINITA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.05m}, {12, 0.92m}, {26, 0.06m} }) },
            { "PASCU ALEXANDRU", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.18m}, {12, 0.82m} }) },
            { "PASCU ALINA MIHAELA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "PASCU MIHAI LUCIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {21, 0.18m}, {22, 0.57m}, {26, 0.25m} }) },
            { "PASCU MIHAI NICOLAE", ("Prof. dr.", new Dictionary<int,decimal>{ {1, 0.46m}, {7, 0.36m}, {12, 0.18m} }) },
            { "PAUN LAURIAN", ("Asist. dr.", new Dictionary<int,decimal>{ {10, 0.07m}, {26, 0.93m} }) },
            { "PAVALACHE ILIE MARIELA", ("Prof. dr.", new Dictionary<int,decimal>{ {27, 0.99m} }) },
            { "PAVEL ECATERINA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "PAVEL GINA MIHAELA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.5m}, {39, 0.5m} }) },
            { "PELIN BOGDAN IULIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.07m}, {26, 0.07m}, {40, 0.86m} }) },
            { "PERNIU DANA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "PETRE ANDREEA", ("Lect. dr.", new Dictionary<int,decimal>{ {15, 0.36m}, {27, 0.51m}, {28, 0.14m} }) },
            { "PETRE IOANA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "PETRIC PAULA", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "PETRICI ANDREI VICTOR", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.57m}, {12, 0.43m} }) },
            { "PETRITAN ION CATALIN", ("Prof. dr.", new Dictionary<int,decimal>{ {8, 0.18m}, {26, 0.82m} }) },
            { "PISARCIUC CRISTIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "PIUARU BRENDA-ANDREEA", ("Asist. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "PLAJER IOANA CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {1, 0.14m}, {2, 0.5m}, {26, 0.36m} }) },
            { "PLESCAN COSTEL", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.01m} }) },
            { "PLUMBOTA LAVINIA", ("Prof. dr.", new Dictionary<int,decimal>{ {25, 0.25m}, {26, 0.77m} }) },
            { "PODASCA PETRU CEZARIO", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "POJALĂ CIPRIAN-VASILE", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "POLEXA ALEXANDRU-CRIȘAN", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "POP DANA MIHAELA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "POPA BOGDAN", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "POPA DANIELA (EFS)", ("Lect. dr.", new Dictionary<int,decimal>{ {40, 1.01m} }) },
            { "POPA DANIELA (PSE)", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.42m}, {27, 0.42m}, {39, 0.17m} }) },
            { "POPA GEORGE-BOGDAN", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 0.07m}, {32, 0.94m} }) },
            { "POPA IULIAN", ("Asist.drd.", new Dictionary<int,decimal>{ {1, 0.07m}, {2, 0.27m}, {12, 0.2m}, {26, 0.47m} }) },
            { "POPA LIOARA RALUCA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.58m}, {39, 0.38m} }) },
            { "POPA LUMINITA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {11, 0.86m}, {12, 0.14m} }) },
            { "POPA ROXANA", ("Asist.drd.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "POPA STEFAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "POPESCU (GHIUTA) IOANA", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.02m} }) },
            { "POPESCU ANCA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "POPESCU MIHAELA VIRGINIA", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 1.01m} }) },
            { "POPESCU OVIDIU", ("Prof. dr.", new Dictionary<int,decimal>{ {1, 1.01m} }) },
            { "POPESCU VLAD", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.27m}, {11, 0.17m}, {21, 0.56m} }) },
            { "POPOVICI BIANCA ELENA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.01m} }) },
            { "POPOVICI-POPESCU ELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.14m}, {9, 0.57m}, {12, 0.1m}, {26, 0.05m}, {27, 0.14m} }) },
            { "POROJAN MIHAELA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "POSTELNICU CRISTIAN CEZAR", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.79m}, {21, 0.21m} }) },
            { "POTINCU LAURA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "POZNA CLAUDIU RADU", ("Prof. dr.", new Dictionary<int,decimal>{ {11, 0.18m}, {12, 0.82m} }) },
            { "PRALEA CRISTIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 0.14m}, {32, 0.86m} }) },
            { "PREDA ULITA ANCA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.33m}, {39, 0.67m} }) },
            { "PROCA ALEXANDRINA MARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.38m}, {12, 0.62m} }) },
            { "PUIU ANDREI", ("Asist. dr.", new Dictionary<int,decimal>{ {7, 0.07m}, {11, 0.93m} }) },
            { "PURCARU IOANA MADALINA", ("Lect. dr.", new Dictionary<int,decimal>{ {25, 0.11m}, {26, 0.9m} }) },
            { "PĂDUREANU VASILE", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.63m}, {12, 0.36m} }) },
            { "RACASAN SERGIU", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "RADU (MATEI) SIMONA CORINA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "RADU ALEXANDRU IONUT", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 1.0m} }) },
            { "RADU CRISTINA IOANA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.33m}, {39, 0.67m} }) },
            { "RADU DORIN", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.01m} }) },
            { "RADU FLORIN", ("Asist.drd.", new Dictionary<int,decimal>{ {7, 0.67m}, {12, 0.33m} }) },
            { "RADU LUCIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "RADU SEBASTIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.98m} }) },
            { "RADUCANU DORINA", ("Prof. dr.", new Dictionary<int,decimal>{ {1, 0.91m}, {2, 0.09m} }) },
            { "RAILEANU SZELES MONICA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "RATULEA GEORGETA GABRIELA", ("Prof. dr.", new Dictionary<int,decimal>{ {22, 0.99m} }) },
            { "RAUTIA IOAN CALIN", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "REPANOVICI ANGELA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.63m}, {15, 0.11m}, {21, 0.09m}, {26, 0.18m} }) },
            { "ROATA IONUT CLAUDIU", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "ROBU DAN NICOLAE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.93m}, {11, 0.07m} }) },
            { "ROGOZEA LILIANA MARCELA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "ROMAN NADINNE ALEXANDRA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "ROSCA IOAN CALIN", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.73m}, {12, 0.27m} }) },
            { "ROSENBERG DAN", ("Conf. dr.", new Dictionary<int,decimal>{ {11, 1.0m} }) },
            { "RUCSANDA MADALINA", ("Prof. dr.", new Dictionary<int,decimal>{ {39, 1.0m} }) },
            { "RUNCEANU-ALBU CARMEN CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "RUS HORATIU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "RUSU IULIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {38, 0.77m}, {39, 0.27m} }) },
            { "RĂDOI-ENCEA RALUCA-STEFANIA", ("Asist.drd.", new Dictionary<int,decimal>{ {10, 0.99m} }) },
            { "SABOU FLORIN-LUCIAN-PETRICĂ", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "SARAMET OANA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 1.01m} }) },
            { "SARBU FLAVIUS AURELIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "SASU ADELA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.76m}, {9, 0.19m}, {12, 0.06m} }) },
            { "SASU LAURA ELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {7, 0.72m}, {11, 0.28m} }) },
            { "SASU LUCIAN-MIRCEA", ("Conf. dr.", new Dictionary<int,decimal>{ {2, 0.33m}, {26, 0.67m} }) },
            { "SAULESCU RADU GABRIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "SAVIN DIANA-CRISTINA", ("Conf. dr.", new Dictionary<int,decimal>{ {1, 0.76m}, {9, 0.13m}, {12, 0.13m} }) },
            { "SAVU CODRUŢ NICOLAE", ("Conf. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "SAVU ELENA CRISTINA", ("Asist. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "SCARNECI-DOMNISORU FLORENTINA", ("Prof. dr.", new Dictionary<int,decimal>{ {21, 0.18m}, {22, 0.84m} }) },
            { "SCARNECIU CAMELIA CORNELIA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "SCARNECIU IOAN", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "SCHWAB-FRÎNCU ANAMARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {1, 0.06m}, {2, 0.38m}, {21, 0.16m}, {22, 0.35m}, {26, 0.06m} }) },
            { "SCRIBA CEZAR", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "SCUTARU MARIA LUMINITA", ("Prof. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "SECHEL GABRIELA", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "SERBAN IOAN", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "SERBAN IONEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "SERBU CLAUDIA GABRIELA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.39m}, {26, 0.18m}, {27, 0.44m} }) },
            { "SIBISAN AURA DANIELA", ("Lect. dr.", new Dictionary<int,decimal>{ {28, 0.39m}, {32, 0.64m} }) },
            { "SIMION GABRIEL", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 0.99m} }) },
            { "SIMON MARINELA CRISTINA", ("Conf. dr.", new Dictionary<int,decimal>{ {22, 1.0m} }) },
            { "SINU RALUCA GEORGIANA", ("Conf. dr.", new Dictionary<int,decimal>{ {28, 0.88m}, {32, 0.13m} }) },
            { "SISMAN VIOREL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "SITOIU ANDREEA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "SOICA ADRIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.99m} }) },
            { "SOICA SIMONA", ("Lect. dr.", new Dictionary<int,decimal>{ {9, 0.21m}, {10, 0.34m}, {12, 0.45m} }) },
            { "SOREA DANIELA", ("Prof. dr.", new Dictionary<int,decimal>{ {22, 0.81m}, {26, 0.18m} }) },
            { "SOREA GHEORGHE DAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 1.0m} }) },
            { "SOVA DANIELA", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.29m}, {12, 0.71m} }) },
            { "SOVAILA SILVIA", ("Asist. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "SPIRCHEZ GHEORGHE COSMIN", ("Lect. dr.", new Dictionary<int,decimal>{ {10, 0.99m} }) },
            { "SPRIDON DELIA - ELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.72m}, {26, 0.29m} }) },
            { "SPÎRCHEZ GEORGETA BIANCA", ("Lect. dr.", new Dictionary<int,decimal>{ {19, 0.99m} }) },
            { "STAN ALEXANDRA", ("Lect. dr.", new Dictionary<int,decimal>{ {8, 0.18m}, {26, 0.83m} }) },
            { "STAN ION GABRIEL", ("Conf. dr.", new Dictionary<int,decimal>{ {2, 0.21m}, {7, 0.11m}, {9, 0.42m}, {12, 0.26m} }) },
            { "STANCA AUREL CORNEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.99m} }) },
            { "STANCIOIU PETRU TUDOR", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 0.04m}, {12, 0.13m}, {26, 0.83m} }) },
            { "STANCIU ANCA ELENA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "STANCIU ELENA MANUELA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "STANCIU MARIANA DOMNICA", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.35m}, {10, 0.45m}, {12, 0.19m} }) },
            { "STANESCU RUXANDRA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.99m} }) },
            { "STARETU IONEL", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.07m}, {12, 0.92m} }) },
            { "STOICA ROXANA ELENA", ("Asist.drd.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "STOICANESCU MARIA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "STROE FANEL", ("Lect. dr.", new Dictionary<int,decimal>{ {21, 0.21m}, {22, 0.79m} }) },
            { "SUCIU CONSTANTIN", ("Prof. dr.", new Dictionary<int,decimal>{ {7, 0.26m}, {11, 0.73m} }) },
            { "SUCIU MARIA-MAGDALENA", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.29m}, {39, 0.71m} }) },
            { "SUCIU TITUS", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "SUMEDREA SILVIA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "SURDU VASILE", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.02m} }) },
            { "SUTEU LIGIA CLAUDIA", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.21m}, {39, 0.78m} }) },
            { "SZILAGYI ANA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.5m}, {39, 0.51m} }) },
            { "SZOCS BOTOND CSABA", ("Lect. dr.", new Dictionary<int,decimal>{ {38, 0.14m}, {39, 0.86m} }) },
            { "SĂFTOIU RĂZVAN GEORGIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.15m}, {28, 0.74m}, {32, 0.12m} }) },
            { "TABIRCA MARIUS SABIN", ("Prof. dr.", new Dictionary<int,decimal>{ {2, 1.02m} }) },
            { "TACHE ILEANA", ("Prof. dr.", new Dictionary<int,decimal>{ {25, 0.09m}, {26, 0.91m} }) },
            { "TALPA NICOLAE", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {10, 0.8m}, {26, 0.2m} }) },
            { "TAMAS FLORIN-LUCIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 0.96m}, {10, 0.04m} }) },
            { "TARNOVEANU MIRELA ADRIANA", ("Lect. dr.", new Dictionary<int,decimal>{ {8, 0.28m}, {12, 0.19m}, {27, 0.53m} }) },
            { "TARULESCU RADU", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 1.0m} }) },
            { "TARULESCU STELIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 1.0m} }) },
            { "TATU OANA", ("Conf. dr.", new Dictionary<int,decimal>{ {28, 1.0m} }) },
            { "TAUS DANIEL", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "TAUS NICOLETA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.77m} }) },
            { "TAȚA ANITHA", ("Asist.drd.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "TECAU ALINA SIMONA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "TEODORESCU ANDREEA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "TEODORESCU DRAGHICESCU HORATIU", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "TERESNEU CORNEL CRISTIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {8, 0.43m}, {26, 0.57m} }) },
            { "TERIȘ ȘTEFAN", ("Lect. dr.", new Dictionary<int,decimal>{ {6, 0.07m}, {40, 0.93m} }) },
            { "TESCASIU BIANCA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "THIERHEIMER WALTER WILHELM", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 0.46m}, {12, 0.54m} }) },
            { "TIEREAN MIRCEA HORIA", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "TIMAR JANOS", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 1.0m} }) },
            { "TIMAR MARIA CRISTINA", ("Prof. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "TINT DIANA", ("Prof. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
            { "TISMĂNAR IOANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 1.01m} }) },
            { "TIŢA NICOLESCU GABRIEL", ("Conf. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "TOADER ADRIAN", ("Lect. dr.", new Dictionary<int,decimal>{ {6, 1.0m} }) },
            { "TOADER SERBAN-SIXTUS", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.33m}, {28, 0.68m} }) },
            { "TODOR RALUCA DANIA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "TOFAN DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 0.11m}, {12, 0.9m} }) },
            { "TOGANEL GEORGE RADU", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.99m} }) },
            { "TOHANEAN DRAGOS IOAN - EFS", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 1.0m} }) },
            { "TOMA SEBASTIAN IONUȚ", ("Conf. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "TOMELE SIMONA CONSTANȚA", ("Asist. dr.", new Dictionary<int,decimal>{ {6, 0.03m}, {7, 0.13m}, {12, 0.06m}, {15, 0.07m}, {40, 0.7m} }) },
            { "TOPALA IOANA ROXANA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "TRIFAN ADRIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "TRUSCA DANIEL DRAGOS", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {9, 1.01m} }) },
            { "TRUTA CAMELIA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "TUCHEL IONU?-VLAD", ("Asist.drd.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "TUDORAN GHEORGHE MARIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {8, 0.14m}, {10, 0.87m} }) },
            { "TULBURE TRAIAN TIBERIU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.73m}, {11, 0.24m} }) },
            { "TURCANU CRISTINA", ("Lect. dr.", new Dictionary<int,decimal>{ {12, 0.28m}, {26, 0.71m} }) },
            { "TURCU IOAN", ("Conf. dr.", new Dictionary<int,decimal>{ {40, 1.01m} }) },
            { "TURCULET ALINA RALUCA", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.04m}, {27, 0.97m} }) },
            { "TUTU DUMITRU CIPRIAN", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.06m}, {39, 0.91m} }) },
            { "UDROIU RAZVAN", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.5m}, {12, 0.51m} }) },
            { "UNCU IONUT", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "UNGUREANU CAMELIA", ("Asist. dr.", new Dictionary<int,decimal>{ {26, 1.0m} }) },
            { "UNGUREANU ELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "UNGUREANU VALENTIN VASILE", ("Conf. dr.", new Dictionary<int,decimal>{ {6, 1.01m} }) },
            { "UNIANU ECATERINA MARIA", ("Lect. dr.", new Dictionary<int,decimal>{ {27, 1.0m} }) },
            { "UNTARU ELENA NICOLETA", ("Prof. dr.", new Dictionary<int,decimal>{ {26, 0.99m} }) },
            { "URETU NOEMI", ("Asist. dr.", new Dictionary<int,decimal>{ {27, 0.2m}, {28, 0.8m} }) },
            { "URSU PETRONELA ELENA", ("Lect. dr.", new Dictionary<int,decimal>{ {6, 0.14m}, {10, 0.07m}, {15, 0.21m}, {40, 0.57m} }) },
            { "VALCEA CRISTINA SILVIA", ("Conf. dr.", new Dictionary<int,decimal>{ {9, 0.22m}, {12, 0.11m}, {28, 0.67m} }) },
            { "VARCIU MIHAI STELIAN", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.96m} }) },
            { "VARGA IOANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.97m} }) },
            { "VARVARICHI LEONA", ("Conf. dr.", new Dictionary<int,decimal>{ {38, 0.5m}, {39, 0.5m} }) },
            { "VASIAN BIANCA IOANA", ("Asist.drd.", new Dictionary<int,decimal>{ {7, 0.13m}, {10, 0.34m}, {11, 0.2m}, {12, 0.27m}, {25, 0.07m} }) },
            { "VASILESCU ANCA", ("Lect. dr.", new Dictionary<int,decimal>{ {2, 0.72m}, {27, 0.29m} }) },
            { "VASILESCU MARIA MAGDALENA", ("Prof. dr.", new Dictionary<int,decimal>{ {8, 0.36m}, {26, 0.63m} }) },
            { "VELEA MARIAN NICOLAE", ("Conf. dr.", new Dictionary<int,decimal>{ {12, 1.0m} }) },
            { "VELICU RADU GABRIEL", ("Prof. dr.", new Dictionary<int,decimal>{ {12, 0.99m} }) },
            { "VIZITIU ANAMARIA", ("Conf. dr.", new Dictionary<int,decimal>{ {11, 0.99m} }) },
            { "VLĂDOIU NASTY MARIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {19, 1.0m} }) },
            { "VODA DANIELA MARIANA", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 1.0m} }) },
            { "VOICESCU CORNELIU GEORGE", ("Prof. dr.", new Dictionary<int,decimal>{ {38, 0.09m}, {39, 0.91m} }) },
            { "VOICU NICOLETA", ("Prof. dr.", new Dictionary<int,decimal>{ {1, 0.72m}, {7, 0.27m} }) },
            { "VOINEA MIHAELA", ("Conf. dr.", new Dictionary<int,decimal>{ {27, 1.01m} }) },
            { "VOLMER MARIUS", ("Conf. dr.", new Dictionary<int,decimal>{ {7, 0.84m}, {12, 0.17m} }) },
            { "VOROVENCII IOSIF", ("Prof. dr.", new Dictionary<int,decimal>{ {8, 0.64m}, {10, 0.09m}, {26, 0.27m} }) },
            { "ZAHARIA CORNELIU", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {7, 0.86m}, {11, 0.14m} }) },
            { "ZAHARIA SEBASTIAN MARIAN", ("Prof. dr.", new Dictionary<int,decimal>{ {9, 0.27m}, {12, 0.74m} }) },
            { "ZAMFIRACHE ALEXANDRA", ("Conf. dr.", new Dictionary<int,decimal>{ {26, 1.01m} }) },
            { "ZELENIUC OCTAVIA", ("Conf. dr.", new Dictionary<int,decimal>{ {10, 1.0m} }) },
            { "ŢĂRANU DAN MARIUS", ("Lect. dr.", new Dictionary<int,decimal>{ {26, 0.29m}, {28, 0.71m} }) },
            { "ȚÂBIAN DANIEL", ("Sef lucr. dr.", new Dictionary<int,decimal>{ {15, 0.99m} }) },
        };

        private class AnsRow { public string Profesor = ""; public string Grad = ""; public string Facultate = ""; public string Departament = ""; public decimal OreConv; public int IdAns; }
        private class AnsProf { public string NumeComplet = ""; public string Departament = ""; public string Facultate = ""; public string GradFunctie = ""; public Dictionary<int, decimal> Fractiuni = new(); }

        #endregion
    }
}