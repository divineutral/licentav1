using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using System.IO.Compression;
using ClosedXML.Excel;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RapoarteController : ControllerBase
    {
        private readonly string _cs;
        private readonly IMemoryCache _cache;
        private const string Green = "#56723e";

        public RapoarteController(IConfiguration cfg, IMemoryCache cache)
        {
            _cs = cfg.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        private int GetAnCurent() => _cache.GetOrCreate("AnCurent_v5", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(4);
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT TOP 1 au.ID_AnUniv
                FROM [AGSIS].[dbo].[AnUniversitar] au
                INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                    ON ppm.ID_AnUniv = au.ID_AnUniv
                GROUP BY au.ID_AnUniv, au.Ordine
                ORDER BY au.Ordine DESC", conn);
            var r = cmd.ExecuteScalar();
            return r != null ? Convert.ToInt32(r) : 45;
        });

        private static decimal NormaFallback(int? id) => id switch
        {
            1 => 11m,
            2 => 12m,
            3 => 14m,
            4 => 15m,
            5 => 18m,
            6 => 11m,
            7 => 18m,
            9 => 15m,
            10 => 14m,
            11 => 14m,
            18 => 2m,
            _ => 18m
        };

        private static string GradANS(int? id, string? den) => id switch
        {
            1 => "Prof. dr.",
            2 => "Conf. dr.",
            3 => "Șef lucr. dr.",
            4 => "Asist. dr.",
            5 => "Preparator",
            6 => "Prof. dr.",
            7 => "Asist. dr.",
            9 => "Asist. dr.",
            10 => "Șef lucr. dr.",
            11 => "Șef lucr. dr.",
            18 => "Asist. dr.",
            19 => "CS I",
            20 => "CS II",
            21 => "CS III",
            _ => den ?? "Asist. dr."
        };

        // =====================================================================
        // DEDUP CTE - elimina duplicatele JSONGrupe din Post_Profesor_Materie
        // Include DenumireCatedra, ID_PlanMaterie_Prestator si DenumireScurtaSpecializare
        // pentru departamentul orei si mentiunile de cuplaje
        // =====================================================================
        private const string DedupCte = @"
        PpmDedup AS (
            SELECT
                ID_Profesor,
                DenumireSpecializare,
                DenumireScurtaSpecializare,
                CASE WHEN CHARINDEX(N'+', DenumireSpecializare) > 0
                     THEN LEFT(DenumireSpecializare, CHARINDEX(N'+', DenumireSpecializare) - 1)
                     ELSE DenumireSpecializare END                   AS SpecCurata,
                Denumire                                             AS DenumireMaterie,
                NrSemestruDinAn,
                TitularSauSuplinitor,
                LimbaDePredare,
                DenumireFormaInv,
                DenumireFacultate                                    AS FacPPM,
                DenumireCatedra                                      AS DeptOra,
                ApartineDeCuplaj,
                MAX(ID_PlanMaterie_Prestator)                        AS ID_PlanMaterie_Prestator,
                MAX(NrOreConventionale)                              AS OreConv,
                MAX(Nr_Ore_Curs)                                     AS OreCurs,
                MAX(Nr_Ore_Seminar)                                  AS OreSem,
                MAX(Nr_Ore_Laborator)                                AS OreLab,
                MAX(Nr_Ore_Proiect)                                  AS OreProj
            FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie]
            WHERE ID_AnUniv = @idAn
            GROUP BY
                ID_Profesor, DenumireSpecializare, DenumireScurtaSpecializare, Denumire,
                NrSemestruDinAn, TitularSauSuplinitor,
                LimbaDePredare, DenumireFormaInv,
                DenumireFacultate, DenumireCatedra, ApartineDeCuplaj
        )";

        // =====================================================================
        // IDENTITATE CTE - sursa suprema pentru identitatea profesorului
        // NOTA: CNP exista in view dar nu il expunem in CTE pentru a evita
        // confuzia; il folosim direct in query-urile care au nevoie de el
        // =====================================================================
        private const string IdenCte = @"
        Identitate AS (
            SELECT
                p.ID_Profesor,
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                p.NumeIntreg,
                p.ID_Facultate,
                p.DenumireFacultate,
                p.ID_Catedra,
                p.DenumireCatedra,
                p.TitularAnUniv,
                p.ID_TipGradDidacticAnUniv                          AS ID_TipGrad,
                p.DenumireGradDidactic,
                ISNULL(n.NrOreConventionaleTitular, 0)              AS NormaDB,
                ISNULL(n.NrOreConventionaleSuplinitor, 18)          AS NormaSupl
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            LEFT JOIN [AGSIS].[pi].[NormaOreConventionale] n
                ON n.ID_TipGradDidactic = p.ID_TipGradDidacticAnUniv
                AND n.ID_AnUniv = @idAn
            WHERE p.ID_AnUnivCatedra = @idAn
        )";

        // =====================================================================
        // REGULA TITULARI OFICIALI (754 conform query verificat)
        // Conditii cumulative:
        //   TitularAnUniv = 1       - angajat titular in anul respectiv
        //   ID_Facultate != 41      - exclude DPPD (4 profesori)
        //   ID_TipGrad IN (1..4,10,11) - grade didactice cu norma definita
        //   p.CNP IS NOT NULL       - identitate verificabila (exclude pensionari fara CNP)
        //   SUM(ppm.NrOreConventionale) > 0 - activitate reala in an
        // =====================================================================
        private const string FiltruTitulariOficiali = @"
              AND i.TitularAnUniv = 1
              AND i.ID_Facultate != 41
              AND i.ID_TipGrad IN (1,2,3,4,10,11)
              AND EXISTS (
                  SELECT 1
                  FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppmx
                  WHERE ppmx.ID_Profesor = i.ID_Profesor
                    AND ppmx.ID_AnUniv = @idAn
                  HAVING SUM(ppmx.NrOreConventionale) > 0
              )";

        // Filtru pe identitate profesor (facultate angajator, departament angajator)
        private const string FiltruIdent = @"
              AND (@idFac = 0 OR i.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR i.ID_Catedra = @idCatedra)";

        private void AddCore(SqlCommand cmd, int idAn, int idFac, int idCat)
        {
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFac);
            cmd.Parameters.AddWithValue("@idCatedra", idCat);
        }

        // =====================================================================
        // Helper: adauga rand de filtre in Excel (randul 2 al fiecarui export)
        // =====================================================================
        private void AddExcelFilters(IXLWorksheet ws, int colCount,
            int idAn, string? fac, string? dept, string? prof,
            string? tipPost, string? formaInv, string? spec, int sem)
        {
            var parts = new List<string>();
            parts.Add($"An: {idAn}");
            if (!string.IsNullOrWhiteSpace(fac) && fac != "0" && fac != "Toate")
                parts.Add($"Facultate: {fac}");
            if (!string.IsNullOrWhiteSpace(dept) && dept != "0" && dept != "Toate")
                parts.Add($"Dept: {dept}");
            if (!string.IsNullOrWhiteSpace(prof) && prof != "Toti")
                parts.Add($"Profesor: {prof}");
            if (!string.IsNullOrWhiteSpace(tipPost) && tipPost != "Toti")
                parts.Add($"Tip: {tipPost}");
            if (!string.IsNullOrWhiteSpace(formaInv) && formaInv != "Toti")
                parts.Add($"Forma: {formaInv}");
            if (!string.IsNullOrWhiteSpace(spec) && spec != "Toti")
                parts.Add($"Spec: {spec}");
            if (sem > 0) parts.Add($"Sem.: {sem}");

            ws.Cell(2, 1).Value = "Filtre: " + string.Join("  |  ", parts);
            ws.Cell(2, 1).Style.Font.Italic = true;
            ws.Cell(2, 1).Style.Font.FontColor = XLColor.FromHtml("#555555");
            ws.Range(2, 1, 2, colCount).Merge();
        }

        private void StyleHdr(IXLRange r)
        {
            r.Style.Fill.BackgroundColor = XLColor.FromHtml(Green);
            r.Style.Font.FontColor = XLColor.White;
            r.Style.Font.Bold = true;
        }

        private string FileName(string raport, string? profesor)
        {
            if (string.IsNullOrWhiteSpace(profesor) || profesor == "Toti")
                return $"{raport}.xlsx";
            var safe = string.Concat(profesor.Take(35)
                .Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
            return $"{safe}_{raport}.xlsx";
        }

        // =====================================================================
        // LISTE DROPDOWNURI
        // =====================================================================
        [HttpGet("liste/ani-universitari")]
        public IActionResult GetAni() => Ok(_cache.GetOrCreate("AniUniv_v5", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
            var lst = new List<object>();
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT au.ID_AnUniv, LTRIM(RTRIM(au.Denumire)) AS Denumire
                FROM [AGSIS].[dbo].[AnUniversitar] au
                INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                    ON ppm.ID_AnUniv = au.ID_AnUniv
                GROUP BY au.ID_AnUniv, au.Denumire, au.Ordine
                ORDER BY au.Ordine DESC", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                lst.Add(new { id = Convert.ToInt32(r[0]), nume = r[1]?.ToString() ?? "" });
            return lst;
        }));

        [HttpGet("liste/facultati")]
        public IActionResult GetFacultati() => Ok(_cache.GetOrCreate("Fac_v6", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
            var lst = new List<object> { new { id = 0, nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT p.ID_Facultate, p.DenumireFacultate
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                WHERE p.ID_AnUnivCatedra = (
                    SELECT TOP 1 au.ID_AnUniv
                    FROM [AGSIS].[dbo].[AnUniversitar] au
                    INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm2
                        ON ppm2.ID_AnUniv = au.ID_AnUniv
                    GROUP BY au.ID_AnUniv, au.Ordine ORDER BY au.Ordine DESC)
                  AND p.DenumireFacultate IS NOT NULL
                ORDER BY p.DenumireFacultate", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                lst.Add(new
                {
                    id = Convert.ToInt32(r["ID_Facultate"]),
                    nume = r["DenumireFacultate"]?.ToString() ?? ""
                });
            return lst;
        }));

        [HttpGet("liste/departamente")]
        public IActionResult GetDepartamente([FromQuery] int? idAnUniv, [FromQuery] int? idFacultate)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = 0, nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT p.ID_Catedra, p.DenumireCatedra
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                WHERE p.ID_AnUnivCatedra = @idAn
                  AND p.DenumireCatedra IS NOT NULL
                  AND LTRIM(RTRIM(p.DenumireCatedra)) != ''
                  AND (@idFac = 0 OR p.ID_Facultate = @idFac)
                ORDER BY p.DenumireCatedra", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                lst.Add(new
                {
                    id = Convert.ToInt32(r["ID_Catedra"]),
                    nume = r["DenumireCatedra"]?.ToString() ?? ""
                });
            return Ok(lst);
        }

        [HttpGet("liste/specializari")]
        public IActionResult GetSpecializari([FromQuery] int? idAnUniv,
            [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = "Toti", nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT
                    CASE WHEN CHARINDEX(N'+', ppm.DenumireSpecializare) > 0
                         THEN LEFT(ppm.DenumireSpecializare, CHARINDEX(N'+', ppm.DenumireSpecializare)-1)
                         ELSE ppm.DenumireSpecializare END AS Spec
                FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                INNER JOIN [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                    ON p.ID_Profesor = ppm.ID_Profesor AND p.ID_AnUnivCatedra = ppm.ID_AnUniv
                WHERE ppm.ID_AnUniv = @idAn
                  AND ppm.DenumireSpecializare IS NOT NULL
                  AND (@idFac = 0 OR p.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR p.ID_Catedra = @idCatedra)
                ORDER BY Spec", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var s = r["Spec"]?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(s)) lst.Add(new { id = s, nume = s });
            }
            return Ok(lst);
        }

        // Filtru ciclu studii din View_FDS
        [HttpGet("liste/cicluri-studii")]
        public IActionResult GetCicluri([FromQuery] int? idAnUniv) => Ok(
            _cache.GetOrCreate($"Cicluri_{idAnUniv ?? GetAnCurent()}", e =>
            {
                e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                int idAn = idAnUniv ?? GetAnCurent();
                var lst = new List<object> { new { id = "Toti", nume = "Toate" } };
                using var conn = new SqlConnection(_cs); conn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT DenumireCicluInv
                    FROM [AGSIS].[dbo].[View_FDS]
                    WHERE ID_AnUniv = @idAn AND DenumireCicluInv IS NOT NULL
                    ORDER BY DenumireCicluInv", conn);
                cmd.Parameters.AddWithValue("@idAn", idAn);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    var v = r["DenumireCicluInv"]?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(v)) lst.Add(new { id = v, nume = v });
                }
                return lst;
            }));

        // Profesori — se incarca fara filtru cand filtrele sunt "Toate"
        [HttpGet("liste/profesori")]
        public IActionResult GetProfesori([FromQuery] int? idAnUniv,
            [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? specializare, [FromQuery] string? tipPost)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = "Toti", nume = "— Toți profesorii —" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            // Se incarca intotdeauna, fara restrictie pe alte filtre
            // Restrictia e doar pe facultate si departament (biroul profesorului)
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT p.NumeIntreg
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                    ON ppm.ID_Profesor = p.ID_Profesor AND ppm.ID_AnUniv = p.ID_AnUnivCatedra
                WHERE p.ID_AnUnivCatedra = @idAn
                  AND p.NumeIntreg IS NOT NULL AND LTRIM(RTRIM(p.NumeIntreg)) != ''
                  AND (@idFac = 0 OR p.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR p.ID_Catedra = @idCatedra)
                  AND (@spec = N'Toti' OR
                       CASE WHEN CHARINDEX(N'+', ppm.DenumireSpecializare) > 0
                            THEN LEFT(ppm.DenumireSpecializare, CHARINDEX(N'+', ppm.DenumireSpecializare)-1)
                            ELSE ppm.DenumireSpecializare END
                       COLLATE Romanian_CI_AS = @spec COLLATE Romanian_CI_AS)
                  AND (@tipPost = N'Toti' OR
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                ORDER BY p.NumeIntreg", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.CommandTimeout = 30;
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var n = r["NumeIntreg"]?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(n)) lst.Add(new { id = n, nume = n });
            }
            return Ok(lst);
        }

        // =====================================================================
        // RAPORT 1: NORMA PROFESORI
        // Departament = ppm.DeptOra (departamentul OREI din PPM, nu al angajatorului)
        // Mentiuni = subquery in View_DetaliereCuplaje:
        //   - Gasim ID_Cuplaj via ID_PlanMaterie_Prestator
        //   - Concatenam DenumireScurtaSpecializare ale partenerilor (altele decat cea curenta)
        // =====================================================================
        private string SqlNorma() => $@"
            WITH
            {IdenCte},
            {DedupCte},
            Baza AS (
                SELECT
                    i.IdentificatorUnic,
                    i.NumeIntreg,
                    i.DenumireFacultate                                     AS FacAngajator,
                    ppm.DeptOra                                             AS DepartamentOra,
                    i.DenumireGradDidactic,
                    CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END AS TipPost,
                    ppm.SpecCurata                                          AS Specializare,
                    ppm.DenumireMaterie,
                    ppm.NrSemestruDinAn                                     AS Semestru,
                    ppm.DenumireFormaInv,
                    ppm.OreConv,
                    ppm.OreCurs,
                    ppm.OreSem + ppm.OreLab + ppm.OreProj                  AS OreAplic,
                    CASE
                        WHEN ppm.ApartineDeCuplaj IS NOT NULL
                        THEN N'Cuplat cu: ' + ISNULL(STUFF((
                            SELECT DISTINCT N', ' + vdc.DenumireScurtaSpecializare
                            FROM [AGSIS].[pi].[View_DetaliereCuplaje] vdc
                            WHERE vdc.ID_AnUniv = @idAn
                              AND vdc.ID_Cuplaj = (
                                  SELECT TOP 1 x.ID_Cuplaj
                                  FROM [AGSIS].[pi].[View_DetaliereCuplaje] x
                                  WHERE x.ID_PlanMaterie_Prestator = ppm.ID_PlanMaterie_Prestator
                                    AND x.ID_AnUniv = @idAn
                              )
                              AND vdc.DenumireScurtaSpecializare != ppm.DenumireScurtaSpecializare
                            FOR XML PATH(N''), TYPE
                        ).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N''), N'încărcare comună')
                        ELSE N''
                    END                                                     AS Mentiuni
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE 1=1
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                  AND (@spec = N'Toti' OR ppm.SpecCurata COLLATE Romanian_CI_AS = @spec COLLATE Romanian_CI_AS)
                  AND (@tipPost = N'Toti' OR
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                  AND (@sem = 0 OR ppm.NrSemestruDinAn = @sem)
                  AND (@formaInv = N'Toti' OR ppm.DenumireFormaInv COLLATE Romanian_CI_AS
                       = @formaInv COLLATE Romanian_CI_AS)
            )
            SELECT
                NumeIntreg, FacAngajator, DepartamentOra,
                DenumireGradDidactic, TipPost, Specializare,
                DenumireMaterie, Semestru, DenumireFormaInv,
                CAST(OreCurs  AS DECIMAL(10,2)) AS OreCurs,
                CAST(OreAplic AS DECIMAL(10,2)) AS OreAplic,
                CAST(OreConv  AS DECIMAL(10,2)) AS OreConv,
                Mentiuni
            FROM Baza
            ORDER BY NumeIntreg, TipPost DESC, Specializare, Semestru";

        private void AddNormaParams(SqlCommand cmd, int idAn, int idFac, int idCat,
            string prof, string spec, string tipPost, int sem, string formaInv)
        {
            AddCore(cmd, idAn, idFac, idCat);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(prof) ? "Toti" : prof.Trim());
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(spec) ? "Toti" : spec.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@sem", sem);
            cmd.Parameters.AddWithValue("@formaInv", string.IsNullOrWhiteSpace(formaInv) ? "Toti" : formaInv.Trim());
        }

        [HttpGet("norma")]
        public async Task<IActionResult> GetNorma(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? specializare,
            [FromQuery] string? tipPost, [FromQuery] int? semestru, [FromQuery] string? formaInvatamant)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma(), conn);
            cmd.CommandTimeout = 120;
            AddNormaParams(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0,
                profesor ?? "", specializare ?? "", tipPost ?? "Toti",
                semestru ?? 0, formaInvatamant ?? "Toti");
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Facultate = r["FacAngajator"]?.ToString() ?? "",
                    Departament = r["DepartamentOra"]?.ToString() ?? "",
                    Grad = r["DenumireGradDidactic"]?.ToString() ?? "",
                    TipPost = r["TipPost"]?.ToString() ?? "",
                    Specializare = r["Specializare"]?.ToString() ?? "",
                    Materie = r["DenumireMaterie"]?.ToString() ?? "",
                    Semestru = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    FormaInv = r["DenumireFormaInv"]?.ToString() ?? "",
                    OreCurs = r["OreCurs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreCurs"]),
                    OreAplicatii = r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
                    OreConv = r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]),
                    Mentiuni = r["Mentiuni"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/norma")]
        public async Task<IActionResult> ExportNorma(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? specializare,
            [FromQuery] string? tipPost, [FromQuery] int? semestru, [FromQuery] string? formaInvatamant,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Dept. Ora"), new DataColumn("Specializare"),
                new DataColumn("Materie"), new DataColumn("Tip Post"),
                new DataColumn("Sem.", typeof(int)), new DataColumn("Forma Inv."),
                new DataColumn("Ore Curs", typeof(decimal)),
                new DataColumn("Ore Aplic.", typeof(decimal)),
                new DataColumn("Ore Conv.", typeof(decimal)),
                new DataColumn("Mențiuni")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma(), conn);
            cmd.CommandTimeout = 120;
            AddNormaParams(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0,
                profesor ?? "", specializare ?? "", tipPost ?? "Toti",
                semestru ?? 0, formaInvatamant ?? "Toti");
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["NumeIntreg"]?.ToString(), r["DepartamentOra"]?.ToString(),
                    r["Specializare"]?.ToString(), r["DenumireMaterie"]?.ToString(),
                    r["TipPost"]?.ToString(),
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["DenumireFormaInv"]?.ToString(),
                    r["OreCurs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreCurs"]),
                    r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
                    r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]),
                    r["Mentiuni"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Norme");
            ws.Cell(1, 1).Value = $"Detaliere Norme Profesori";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn,
                numeFacultate, numeDepartament, profesor, tipPost, formaInvatamant, specializare, semestru ?? 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var c in new[] { "Ore Curs", "Ore Aplic.", "Ore Conv." })
                tbl.Field(c).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count));
            ws.Columns(1, 11).AdjustToContents();
            ws.Column(12).Width = 60; ws.Column(12).Style.Alignment.WrapText = true;
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Norme", profesor));
        }

        // =====================================================================
        // RAPORT 2: TOTALURI NORME (IF/ID/IFR)
        // =====================================================================
        private string SqlTotaluri() => $@"
            WITH
            {IdenCte},
            {DedupCte},
            Baza AS (
                SELECT
                    i.IdentificatorUnic, i.NumeIntreg,
                    i.DenumireFacultate, i.DenumireCatedra,
                    CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END AS TipPost,
                    CASE
                        WHEN ppm.DenumireFormaInv = N'Cu frecvență'           THEN 'IF'
                        WHEN ppm.DenumireFormaInv = N'Frecvență redusă'       THEN 'IFR'
                        WHEN ppm.DenumireFormaInv = N'Învățământ la distanță' THEN 'ID'
                        ELSE 'IF'
                    END AS FormaInv,
                    ppm.DenumireMaterie, ppm.NrSemestruDinAn, ppm.OreConv
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE 1=1
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                  AND (@tipPost = N'Toti' OR
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
            ),
            Dedup2 AS (
                SELECT IdentificatorUnic, NumeIntreg, DenumireFacultate, DenumireCatedra,
                       TipPost, FormaInv, DenumireMaterie, NrSemestruDinAn, MAX(OreConv) AS OreD
                FROM Baza
                GROUP BY IdentificatorUnic, NumeIntreg, DenumireFacultate, DenumireCatedra,
                         TipPost, FormaInv, DenumireMaterie, NrSemestruDinAn
            ),
            Agreg AS (
                SELECT IdentificatorUnic, NumeIntreg,
                    MAX(DenumireFacultate) AS Facultate,
                    MAX(DenumireCatedra)   AS Departament, TipPost,
                    CAST(SUM(CASE WHEN FormaInv='IF'  THEN OreD ELSE 0 END) AS DECIMAL(10,2)) AS OreIF,
                    CAST(SUM(CASE WHEN FormaInv='ID'  THEN OreD ELSE 0 END) AS DECIMAL(10,2)) AS OreID,
                    CAST(SUM(CASE WHEN FormaInv='IFR' THEN OreD ELSE 0 END) AS DECIMAL(10,2)) AS OreIFR,
                    CAST(SUM(OreD) AS DECIMAL(10,2))                                          AS TotalOre
                FROM Dedup2
                GROUP BY IdentificatorUnic, NumeIntreg, TipPost
            )
            SELECT NumeIntreg, Facultate, Departament, TipPost,
                   OreIF, OreID, OreIFR, TotalOre,
                   CAST(TotalOre * 14 AS DECIMAL(10,2)) AS TotalAnual
            FROM Agreg ORDER BY NumeIntreg, TipPost DESC";

        [HttpGet("norma-totaluri")]
        public async Task<IActionResult> GetNormaTotaluri(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri(), conn);
            cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    TipPost = r["TipPost"]?.ToString() ?? "",
                    OreIF = r["OreIF"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIF"]),
                    OreID = r["OreID"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreID"]),
                    OreIFR = r["OreIFR"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIFR"]),
                    TotalOreConv = r["TotalOre"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOre"]),
                    TotalAnual = r["TotalAnual"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalAnual"])
                });
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public async Task<IActionResult> ExportNormaTotaluri(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"),
                new DataColumn("Tip Post"),
                new DataColumn("Ore IF", typeof(decimal)), new DataColumn("Ore ID", typeof(decimal)),
                new DataColumn("Ore IFR", typeof(decimal)), new DataColumn("Total Ore Conv.", typeof(decimal)),
                new DataColumn("Total Anual (×14)", typeof(decimal))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri(), conn);
            cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["NumeIntreg"]?.ToString(), r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(), r["TipPost"]?.ToString(),
                    r["OreIF"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIF"]),
                    r["OreID"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreID"]),
                    r["OreIFR"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIFR"]),
                    r["TotalOre"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOre"]),
                    r["TotalAnual"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalAnual"]));
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Totaluri Norme");
            ws.Cell(1, 1).Value = "Totaluri Norme (IF / ID / IFR)";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, tipPost, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var c in new[] { "Ore IF", "Ore ID", "Ore IFR", "Total Ore Conv.", "Total Anual (×14)" })
                tbl.Field(c).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL GENERAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Totaluri_Norme", profesor));
        }

        // =====================================================================
        // RAPORT 3: DISTRIBUTIE ORE PE PROGRAME
        // =====================================================================
        private string SqlDistrib() => $@"
            WITH
            {IdenCte},
            {DedupCte},
            OrePerProg AS (
                SELECT
                    i.IdentificatorUnic, i.NumeIntreg,
                    i.DenumireFacultate, i.DenumireCatedra,
                    ppm.SpecCurata AS Program,
                    SUM(ppm.OreConv) AS OreProgram
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE 1=1
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                GROUP BY i.IdentificatorUnic, i.NumeIntreg, i.DenumireFacultate,
                         i.DenumireCatedra, ppm.SpecCurata
            ),
            TotUniv AS (
                SELECT IdentificatorUnic, SUM(OreProgram) AS TotalOreUniv
                FROM OrePerProg GROUP BY IdentificatorUnic
            )
            SELECT
                o.NumeIntreg, o.DenumireFacultate, o.DenumireCatedra, o.Program,
                CAST(o.OreProgram   AS DECIMAL(10,2)) AS OreProgram,
                CAST(t.TotalOreUniv AS DECIMAL(10,2)) AS TotalOreUniv,
                CAST(CASE WHEN t.TotalOreUniv > 0
                     THEN ROUND(o.OreProgram / t.TotalOreUniv * 100, 2)
                     ELSE 0 END AS DECIMAL(10,2))     AS Procent
            FROM OrePerProg o
            INNER JOIN TotUniv t ON t.IdentificatorUnic = o.IdentificatorUnic
            ORDER BY o.NumeIntreg, Procent DESC";

        [HttpGet("distributie-ore")]
        public async Task<IActionResult> GetDistrib(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDistrib(), conn);
            cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Facultate = r["DenumireFacultate"]?.ToString() ?? "",
                    Departament = r["DenumireCatedra"]?.ToString() ?? "",
                    Program = r["Program"]?.ToString() ?? "",
                    OreProgram = r["OreProgram"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreProgram"]),
                    TotalOreUniv = r["TotalOreUniv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreUniv"]),
                    Procent = r["Procent"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Procent"])
                });
            return Ok(result);
        }

        [HttpGet("export/distributie-ore")]
        public async Task<IActionResult> ExportDistrib(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"),
                new DataColumn("Program Studiu"),
                new DataColumn("Ore Program", typeof(decimal)),
                new DataColumn("Total Ore Univ.", typeof(decimal)),
                new DataColumn("Procent %", typeof(decimal))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDistrib(), conn);
            cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["NumeIntreg"]?.ToString(), r["DenumireFacultate"]?.ToString(),
                    r["DenumireCatedra"]?.ToString(), r["Program"]?.ToString(),
                    r["OreProgram"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreProgram"]),
                    r["TotalOreUniv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreUniv"]),
                    r["Procent"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Procent"]));
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Distributie Ore");
            ws.Cell(1, 1).Value = "Distribuție Ore pe Programe";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, null, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Program").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Distributie_Ore", profesor));
        }

        // =====================================================================
        // RAPORT 4: LIMBI STRAINE
        // Structura noua: Nr, Profesor, Semestru, Total Ore, Limbi Predate,
        //                 Program Studiu, Ciclu Studii, Specializare
        // Ciclu studii din View_FDS.DenumireCicluInv via JOIN dublu
        // =====================================================================
        private string SqlLimbi() => $@"
            WITH
            {IdenCte},
            {DedupCte},
            Filtrat AS (
                SELECT
                    i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra,
                    CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END AS TipPost,
                    ppm.OreConv, ppm.NrSemestruDinAn, ppm.DenumireMaterie,
                    ppm.LimbaDePredare,
                    ppm.SpecCurata                                       AS ProgramStudiu,
                    ISNULL(fds.DenumireCicluInv, N'Nespecificat')        AS CicluStudii,
                    ISNULL(fds.DenumireSpecializare, ppm.SpecCurata)     AS SpecFDS
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                LEFT JOIN [AGSIS].[dbo].[View_FDS] fds
                    ON ppm.SpecCurata COLLATE Romanian_CI_AS
                       = fds.DenumireSpecializare COLLATE Romanian_CI_AS
                    AND ppm.FacPPM COLLATE Romanian_CI_AS
                       = fds.DenumireFacultate COLLATE Romanian_CI_AS
                    AND fds.ID_AnUniv = @idAn
                    AND fds.id_metaspecializare > 0
                WHERE 1=1
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                  AND (@tipPost = N'Toti' OR
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                  AND (@ciclu = N'Toti' OR ISNULL(fds.DenumireCicluInv, N'Nespecificat')
                       COLLATE Romanian_CI_AS = @ciclu COLLATE Romanian_CI_AS)
                  AND ppm.LimbaDePredare IS NOT NULL
                  AND ppm.LimbaDePredare NOT IN (N'Romana', N'Română', N'Româna',
                                                 N'Romana/Engleza', N'Româna/Engleza')
                  AND ppm.OreConv > 0
            ),
            Dedup2 AS (
                SELECT NumeIntreg, DenumireMaterie, NrSemestruDinAn, TipPost,
                       LimbaDePredare, ProgramStudiu, CicluStudii, SpecFDS,
                       MAX(OreConv) AS OreD
                FROM Filtrat
                GROUP BY NumeIntreg, DenumireMaterie, NrSemestruDinAn, TipPost,
                         LimbaDePredare, ProgramStudiu, CicluStudii, SpecFDS
            )
            SELECT
                NumeIntreg,
                NrSemestruDinAn                                      AS Semestru,
                CAST(SUM(OreD) AS DECIMAL(10,2))                     AS TotalOre,
                STUFF((
                    SELECT DISTINCT N', ' + d2.LimbaDePredare
                    FROM Dedup2 d2 WHERE d2.NumeIntreg = Dedup2.NumeIntreg
                    FOR XML PATH(N''), TYPE
                ).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N'')          AS LimbiPredate,
                ProgramStudiu,
                CicluStudii,
                SpecFDS                                              AS Specializare
            FROM Dedup2
            GROUP BY NumeIntreg, NrSemestruDinAn, ProgramStudiu, CicluStudii, SpecFDS
            HAVING SUM(OreD) > 0
            ORDER BY NumeIntreg, NrSemestruDinAn";

        [HttpGet("limbi-straine")]
        public async Task<IActionResult> GetLimbi(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost, [FromQuery] string? cicluStudii)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi(), conn);
            cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@ciclu", string.IsNullOrWhiteSpace(cicluStudii) ? "Toti" : cicluStudii.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Semestru = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    TotalOre = r["TotalOre"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOre"]),
                    LimbiPredate = r["LimbiPredate"]?.ToString() ?? "",
                    ProgramStudiu = r["ProgramStudiu"]?.ToString() ?? "",
                    CicluStudii = r["CicluStudii"]?.ToString() ?? "",
                    Specializare = r["Specializare"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/limbi-straine")]
        public async Task<IActionResult> ExportLimbi(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost, [FromQuery] string? cicluStudii,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Semestru", typeof(int)),
                new DataColumn("Total Ore (h)", typeof(decimal)),
                new DataColumn("Limbi Predate"),
                new DataColumn("Program Studiu"),
                new DataColumn("Ciclu Studii"),
                new DataColumn("Specializare")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi(), conn);
            cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@ciclu", string.IsNullOrWhiteSpace(cicluStudii) ? "Toti" : cicluStudii.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["NumeIntreg"]?.ToString(),
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["TotalOre"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOre"]),
                    r["LimbiPredate"]?.ToString(), r["ProgramStudiu"]?.ToString(),
                    r["CicluStudii"]?.ToString(), r["Specializare"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Limbi Straine");
            ws.Cell(1, 1).Value = "Raport Limbi Străine";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, tipPost, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Total Ore (h)").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Limbi_Straine", profesor));
        }
        // =====================================================================
        // RAPORT 5: DISCIPLINE PER PROFESOR
        // =====================================================================
        private string SqlDisc(string? formaInvFilter = null) => $@"
            WITH
            {IdenCte},
            {DedupCte},
            Distinct_Mat AS (
                SELECT DISTINCT
                    i.IdentificatorUnic, i.NumeIntreg,
                    i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic,
                    ppm.DenumireMaterie, ppm.DenumireFormaInv
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE 1=1
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                  AND ppm.DenumireMaterie IS NOT NULL
                  AND LTRIM(RTRIM(ppm.DenumireMaterie)) != N''
                  {(string.IsNullOrWhiteSpace(formaInvFilter)
                      ? ""
                      : "AND ppm.DenumireFormaInv COLLATE Romanian_CI_AS = @formaInvFilter COLLATE Romanian_CI_AS")}
            )
            SELECT
                dm.NumeIntreg,
                MAX(dm.DenumireFacultate)     AS Facultate,
                MAX(dm.DenumireCatedra)       AS Departament,
                MAX(dm.DenumireGradDidactic)  AS Grad,
                STUFF((
                    SELECT DISTINCT N' | ' + dm2.DenumireMaterie
                    FROM Distinct_Mat dm2
                    WHERE dm2.IdentificatorUnic = dm.IdentificatorUnic
                      AND dm2.DenumireMaterie IS NOT NULL
                    FOR XML PATH(N''), TYPE
                ).value(N'.', N'NVARCHAR(MAX)'), 1, 3, N'') AS Discipline,
                COUNT(DISTINCT dm.DenumireMaterie) AS NrDisc
            FROM Distinct_Mat dm
            GROUP BY dm.IdentificatorUnic, dm.NumeIntreg
            ORDER BY dm.NumeIntreg";

        [HttpGet("discipline")]
        public async Task<IActionResult> GetDisc(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDisc(), conn);
            cmd.CommandTimeout = 180;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    Grad = r["Grad"]?.ToString() ?? "",
                    Discipline = r["Discipline"]?.ToString() ?? "",
                    NrDiscipline = r["NrDisc"] == DBNull.Value ? 0 : Convert.ToInt32(r["NrDisc"])
                });
            return Ok(result);
        }

        [HttpGet("export/discipline-zip")]
        public async Task<IActionResult> ExportDisciplineZip(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var forme = new[]
            {
                ("Cu frecvență",           "IF"),
                ("Frecvență redusă",       "IFR"),
                ("Învățământ la distanță", "ID")
            };
            using var memZip = new MemoryStream();
            using (var archive = new ZipArchive(memZip, ZipArchiveMode.Create, true))
            {
                foreach (var (formaVal, formaLabel) in forme)
                {
                    var dt = new DataTable();
                    dt.Columns.AddRange(new[]{
                        new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                        new DataColumn("Facultate"), new DataColumn("Departament"),
                        new DataColumn("Grad"), new DataColumn("Discipline"),
                        new DataColumn("Nr. Disc.", typeof(int))
                    });
                    using var conn = new SqlConnection(_cs); await conn.OpenAsync();
                    using var cmd = new SqlCommand(SqlDisc(formaVal), conn);
                    cmd.CommandTimeout = 180;
                    AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
                    cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor!.Trim());
                    cmd.Parameters.AddWithValue("@formaInvFilter", formaVal);
                    using var r = await cmd.ExecuteReaderAsync();
                    int nr = 1;
                    while (await r.ReadAsync())
                        dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["Facultate"]?.ToString(),
                            r["Departament"]?.ToString(), r["Grad"]?.ToString(),
                            r["Discipline"]?.ToString(),
                            r["NrDisc"] == DBNull.Value ? 0 : Convert.ToInt32(r["NrDisc"]));

                    var entry = archive.CreateEntry($"Discipline_{formaLabel}.xlsx");
                    using var es = entry.Open();
                    using var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add($"Disc {formaLabel}");
                    ws.Cell(1, 1).Value = $"Discipline predate — {formaLabel} | An: {idAn}";
                    ws.Cell(1, 1).Style.Font.Bold = true;
                    ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
                    ws.Range(1, 1, 1, dt.Columns.Count).Merge();
                    var tbl = ws.Cell(3, 1).InsertTable(dt);
                    tbl.Theme = XLTableTheme.None;
                    StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
                    ws.Columns(1, 5).AdjustToContents();
                    ws.Column(6).Width = 80; ws.Column(6).Style.Alignment.WrapText = true;
                    using var wbs = new MemoryStream();
                    wb.SaveAs(wbs); wbs.Position = 0; wbs.CopyTo(es);
                }
            }
            memZip.Position = 0;
            return File(memZip.ToArray(), "application/zip", "Discipline_Predate.zip");
        }

        // =====================================================================
        // RAPORT 6: TITULARI OFICIALI
        // Aplica FiltruTitulariOficiali (TitularAnUniv=1 + ID_Facultate!=41
        //   + ID_TipGrad IN(1,2,3,4,10,11) + CNP IS NOT NULL + ore > 0)
        // =====================================================================
        private string SqlTitulari() => $@"
            WITH {IdenCte}
            SELECT i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic
            FROM Identitate i
            WHERE 1=1
              {FiltruTitulariOficiali}
              {FiltruIdent}
              AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
            ORDER BY i.NumeIntreg";

        [HttpGet("titulari")]
        public async Task<IActionResult> GetTitulari(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulari(), conn);
            cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Facultate = r["DenumireFacultate"]?.ToString() ?? "",
                    Departament = r["DenumireCatedra"]?.ToString() ?? "",
                    Grad = r["DenumireGradDidactic"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/titulari")]
        public async Task<IActionResult> ExportTitulari(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"),
                new DataColumn("Grad Didactic")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulari(), conn);
            cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["DenumireFacultate"]?.ToString(),
                    r["DenumireCatedra"]?.ToString(), r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Titulari");
            ws.Cell(1, 1).Value = "Cadre Didactice Titulare (Oficiali)";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, null, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowFunction = XLTotalsRowFunction.Count;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Titulari", profesor));
        }

        // =====================================================================
        // RAPORT 7: COLABORATORI
        // =====================================================================
        private string SqlColab() => $@"
            WITH {IdenCte}
            SELECT DISTINCT i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic
            FROM Identitate i
            WHERE (i.TitularAnUniv = 0 OR i.TitularAnUniv IS NULL)
              AND i.NumeIntreg IS NOT NULL AND LTRIM(RTRIM(i.NumeIntreg)) != N''
              {FiltruIdent}
            ORDER BY i.NumeIntreg";

        [HttpGet("colaboratori")]
        public async Task<IActionResult> GetColab(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColab(), conn);
            cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Facultate = r["DenumireFacultate"]?.ToString() ?? "",
                    Departament = r["DenumireCatedra"]?.ToString() ?? "",
                    Grad = r["DenumireGradDidactic"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/colaboratori")]
        public async Task<IActionResult> ExportColab(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"),
                new DataColumn("Grad Didactic")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColab(), conn);
            cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["DenumireFacultate"]?.ToString(),
                    r["DenumireCatedra"]?.ToString(), r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Colaboratori");
            ws.Cell(1, 1).Value = "Asociați / Colaboratori";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, null, null, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowFunction = XLTotalsRowFunction.Count;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Colaboratori.xlsx");
        }

        // =====================================================================
        // RAPORT 8: ANS
        // FIX eroare "Invalid column name 'CNP'":
        //   - FiltruTitulariOficiali referea i.CNP (nu exista in SELECT al CTE)
        //   - Solutie: CNP IS NOT NULL se aplica direct in WHERE cu subquery pe VIEW
        //   - Norma: din View direct (nu prin CTE care nu o expune)
        // =====================================================================
        private const string SqlTitulariANS = @"
            SELECT DISTINCT
                p.ID_Profesor,
                p.NumeIntreg,
                p.DenumireFacultate,
                p.DenumireCatedra,
                p.DenumireGradDidactic,
                p.ID_TipGradDidacticAnUniv                       AS ID_TipGrad,
                ISNULL(n.NrOreConventionaleTitular, 0)           AS NormaDB
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            LEFT JOIN [AGSIS].[pi].[NormaOreConventionale] n
                ON n.ID_TipGradDidactic = p.ID_TipGradDidacticAnUniv
                AND n.ID_AnUniv = @idAn
            WHERE p.ID_AnUnivCatedra = @idAn
              AND p.TitularAnUniv = 1
              AND p.ID_Facultate != 41
              AND p.ID_TipGradDidacticAnUniv IN (1,2,3,4,10,11)
              AND p.CNP IS NOT NULL
              AND (@idFac = 0 OR p.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR p.ID_Catedra = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppmx
                  WHERE ppmx.ID_Profesor = p.ID_Profesor AND ppmx.ID_AnUniv = @idAn
                  HAVING SUM(ppmx.NrOreConventionale) > 0
              )
            ORDER BY p.NumeIntreg";

        private const string SqlOreANS = @"
            SELECT
                vc.ID_Profesor,
                Map.ID_ANS,
                -- Formula verificata: SUM(ore_conventionale) / 14 saptamani / norma_legala
                -- Returnam Ore brute (SUM), impartirea la 14 si la norma se face in C# per profesor
                CAST(SUM(vc.NrOreConventionale) AS DECIMAL(18,4)) AS OreConv
            FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vc
            CROSS APPLY (
                SELECT TOP 1 fds.ID_N_Domeniu_Studiu_ANS AS ID_ANS
                FROM [AGSIS].[dbo].[View_FDS] fds
                WHERE fds.DenumireSpecializare COLLATE Romanian_CI_AS
                      = vc.DenumireSpecializare COLLATE Romanian_CI_AS
                  AND fds.DenumireFacultate COLLATE Romanian_CI_AS
                      = vc.DenumireFacultate COLLATE Romanian_CI_AS
                  AND fds.ID_AnUniv = @idAn
                  AND fds.id_metaspecializare > 0
                  AND fds.ID_N_Domeniu_Studiu_ANS IS NOT NULL
            ) Map
            WHERE vc.ID_AnUniv = @idAn
            GROUP BY vc.ID_Profesor, Map.ID_ANS
            HAVING SUM(vc.NrOreConventionale) > 0";

        private class TitAns
        {
            public int IdProf { get; set; }
            public string Nume { get; set; } = "";
            public string Fac { get; set; } = "";
            public string Dept { get; set; } = "";
            public string GradD { get; set; } = "";
            public int? IdTip { get; set; }
            public decimal Norma { get; set; }
        }

        private async Task<List<(int Id, string Den)>> LoadDomeniiAsync(SqlConnection conn)
        {
            var lst = new List<(int, string)>();
            using var cmd = new SqlCommand(
                "SELECT ID_Element, Denumire FROM [AGSIS].[dbo].[N_RAMURA_STIINTA_ANS] ORDER BY ID_Element", conn);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                lst.Add((Convert.ToInt32(r["ID_Element"]), r["Denumire"]?.ToString() ?? ""));
            return lst;
        }

        private async Task<List<TitAns>> LoadTitulariANSAsync(SqlConnection conn, int idAn, int idFac, int idCat)
        {
            var lst = new List<TitAns>();
            using var cmd = new SqlCommand(SqlTitulariANS, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFac);
            cmd.Parameters.AddWithValue("@idCatedra", idCat);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                int? idTip = r["ID_TipGrad"] == DBNull.Value ? null : (int?)Convert.ToInt32(r["ID_TipGrad"]);
                decimal nDb = r["NormaDB"] == DBNull.Value ? 0m : Convert.ToDecimal(r["NormaDB"]);
                lst.Add(new TitAns
                {
                    IdProf = Convert.ToInt32(r["ID_Profesor"]),
                    Nume = r["NumeIntreg"]?.ToString() ?? "",
                    Fac = r["DenumireFacultate"]?.ToString() ?? "",
                    Dept = r["DenumireCatedra"]?.ToString() ?? "",
                    GradD = r["DenumireGradDidactic"]?.ToString() ?? "",
                    IdTip = idTip,
                    Norma = nDb > 0 ? nDb : NormaFallback(idTip)
                });
            }
            return lst;
        }

        private async Task<Dictionary<int, Dictionary<int, decimal>>> LoadOreANSAsync(SqlConnection conn, int idAn)
        {
            var dict = new Dictionary<int, Dictionary<int, decimal>>();
            using var cmd = new SqlCommand(SqlOreANS, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                int idP = Convert.ToInt32(r["ID_Profesor"]);
                int idA = Convert.ToInt32(r["ID_ANS"]);
                // OreConv din view este suma orelor conventionale pe tot semestrul
                // Impartim la 14 (saptamani) pentru a obtine ore saptamanale - analog cu formula verificata
                decimal oreConv = r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]);
                decimal oreSapt = oreConv / 14m;
                if (!dict.ContainsKey(idP)) dict[idP] = new();
                dict[idP].TryAdd(idA, 0m);
                dict[idP][idA] += oreSapt;
            }
            return dict;
        }

        [HttpGet("raport-ans")]
        public async Task<IActionResult> GetAns(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            var domenii = await LoadDomeniiAsync(conn);
            var titulari = await LoadTitulariANSAsync(conn, idAn, idFacultate ?? 0, idCatedra ?? 0);
            var oreAns = await LoadOreANSAsync(conn, idAn);

            var profesori = new List<object>();
            int nrCrt = 1;
            foreach (var t in titulari)
            {
                var frac = new Dictionary<string, decimal>();
                if (oreAns.TryGetValue(t.IdProf, out var oreP))
                {
                    // oreP contine deja ore saptamanale (NrOreConv / 14)
                    // Formula ANS: fractiune = ore_sapt_domeniu / norma_legala
                    if (t.Norma > 0)
                        foreach (var kv in oreP)
                        {
                            var dom = domenii.FirstOrDefault(d => d.Id == kv.Key);
                            if (dom.Den == null) continue;
                            decimal f = Math.Round(kv.Value / t.Norma, 2);
                            if (f > 0) frac[dom.Den] = f;
                        }
                }
                profesori.Add(new
                {
                    NrCrt = nrCrt++,
                    NumeComplet = t.Nume,
                    Facultate = t.Fac,
                    Departament = t.Dept,
                    GradFunctie = GradANS(t.IdTip, t.GradD),
                    DomeniiMapate = frac
                });
            }
            return Ok(new { Domenii = domenii.Select(d => d.Den).ToList(), Profesori = profesori });
        }

        [HttpGet("export/raport-ans")]
        public async Task<IActionResult> ExportAns(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            var domenii = await LoadDomeniiAsync(conn);
            var titulari = await LoadTitulariANSAsync(conn, idAn, idFacultate ?? 0, idCatedra ?? 0);
            var oreAns = await LoadOreANSAsync(conn, idAn);

            int nrD = domenii.Count, colTot = 9 + nrD + 1;
            var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("CD DRU");
            var gC = XLColor.FromHtml(Green);
            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, colTot).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov";
            ws.Range(3, 1, 3, 6).Merge();
            string[] antet = {"Nr.\nCrt.","Nume si prenume","CNP","Functie","Forma\nangajare",
                              "Cond.\ndoctorat","Varsta","Facultate","Departament"};
            for (int c = 1; c <= 9; c++) { ws.Cell(5, c).Value = antet[c - 1]; ws.Range(5, c, 7, c).Merge(); }
            ws.Cell(5, colTot).Value = "Total"; ws.Range(5, colTot, 7, colTot).Merge();
            for (int i = 0; i < nrD; i++) { ws.Cell(6, 10 + i).Value = domenii[i].Den; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }
            for (int c = 1; c <= 9; c++) ws.Cell(8, c).Value = ((char)('A' + c - 1)).ToString();
            for (int i = 0; i < nrD; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, colTot).Value = nrD;
            for (int row = 5; row <= 8; row++) for (int col = 1; col <= colTot; col++)
            {
                var cell = ws.Cell(row, col);
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.WrapText = true;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            ws.Range(5, 1, 5, colTot).Style.Fill.BackgroundColor = gC;
            ws.Range(5, 1, 5, colTot).Style.Font.FontColor = XLColor.White;
            ws.Column(1).Width = 5; ws.Column(2).Width = 30; ws.Column(3).Width = 14;
            ws.Column(4).Width = 22; ws.Column(5).Width = 10; ws.Column(6).Width = 12;
            ws.Column(7).Width = 8; ws.Column(8).Width = 28; ws.Column(9).Width = 28;
            for (int c = 10; c <= colTot; c++) ws.Column(c).Width = 12;

            int row2 = 9, nrCrt = 1;
            foreach (var t in titulari)
            {
                ws.Cell(row2, 1).Value = nrCrt++;
                ws.Cell(row2, 2).Value = t.Nume;
                ws.Cell(row2, 3).Value = "";
                ws.Cell(row2, 4).Value = GradANS(t.IdTip, t.GradD);
                ws.Cell(row2, 5).Value = ""; ws.Cell(row2, 6).Value = "";
                ws.Cell(row2, 7).Value = ""; ws.Cell(row2, 8).Value = t.Fac;
                ws.Cell(row2, 9).Value = t.Dept;
                decimal totFrac = 0m;
                if (oreAns.TryGetValue(t.IdProf, out var oreP))
                {
                    // oreP contine deja ore saptamanale (NrOreConv / 14)
                    // Formula ANS: fractiune = ore_sapt_domeniu / norma_legala
                    if (t.Norma > 0)
                        for (int i = 0; i < nrD; i++)
                            if (oreP.TryGetValue(domenii[i].Id, out decimal oreD))
                            {
                                decimal f = Math.Round(oreD / t.Norma, 2);
                                if (f > 0)
                                {
                                    ws.Cell(row2, 10 + i).Value = (double)f;
                                    ws.Cell(row2, 10 + i).Style.NumberFormat.Format = "0.00";
                                    totFrac += f;
                                }
                            }
                }
                ws.Cell(row2, colTot).Value = (double)Math.Round(totFrac, 2);
                ws.Cell(row2, colTot).Style.NumberFormat.Format = "0.00";
                for (int c = 1; c <= colTot; c++)
                    ws.Cell(row2, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                row2++;
            }
            using var stream = new MemoryStream(); wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Raport_ANS.xlsx");
        }
    }
}