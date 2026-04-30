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

        private int GetAnCurent() => _cache.GetOrCreate("AnCurent_v6", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(4);
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT TOP 1 au.ID_AnUniv
                FROM [AGSIS].[dbo].[AnUniversitar] au
                INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm ON ppm.ID_AnUniv = au.ID_AnUniv
                GROUP BY au.ID_AnUniv, au.Ordine ORDER BY au.Ordine DESC", conn);
            var r = cmd.ExecuteScalar();
            return r != null ? Convert.ToInt32(r) : 45;
        });

        // =====================================================================
        // DEDUP CTE — FIX MAJOR
        //
        // PROBLEMA ANTERIOARA: GROUP BY includea DenumireSpecializare si
        //   DenumireScurtaSpecializare → un curs cuplat la 3 specializari
        //   genera 3 randuri × NrOreConventionale → ore × 3 (GRESIT).
        //
        // FIX: GROUP BY pe ID_PlanMaterie_Prestator (identificator unic al
        //   activitatii in planul de invatamant). Specializarile sunt colectate
        //   ca metadate de afisare prin STRING_AGG echivalent (FOR XML).
        //   Orele fizice se calculeaza O SINGURA DATA per activitate.
        //
        // FIX WHITESPACE: LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(Denumire,
        //   ' ','><'),'<>',''),'>< ',' '))) elimina spatiile multiple interioare.
        //   Aceasta este metoda standard T-SQL pentru normalizare whitespace.
        // =====================================================================
        private const string DedupCte = @"
        PpmDedup AS (
            SELECT
                ID_Profesor,
                -- Specializarea primara (prima alfabetic) — doar pentru afisare
                MIN(DenumireSpecializare)      AS DenumireSpecializare,
                MIN(DenumireScurtaSpecializare)AS DenumireScurtaSpecializare,
                CASE WHEN CHARINDEX(N'+', MIN(DenumireSpecializare)) > 0
                     THEN LEFT(MIN(DenumireSpecializare), CHARINDEX(N'+', MIN(DenumireSpecializare)) - 1)
                     ELSE MIN(DenumireSpecializare) END                   AS SpecCurata,
                -- Normalizare whitespace pe denumirea materiei
                LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(Denumire, N'  ', N'><'), N'<>', N''), N'><', N' ')))
                                               AS DenumireMaterie,
                NrSemestruDinAn,
                TitularSauSuplinitor,
                MIN(LimbaDePredare)            AS LimbaDePredare,
                MIN(DenumireFormaInv)          AS DenumireFormaInv,
                MIN(DenumireFacultate)         AS FacPPM,
                MIN(DenumireCatedra)           AS DeptOra,
                MAX(ApartineDeCuplaj)          AS ApartineDeCuplaj,
                ID_PlanMaterie_Prestator,
                -- Ore: O singura valoare per activitate fizica
                MAX(NrOreConventionale)        AS OreConv,
                MAX(Nr_Ore_Curs)               AS OreCurs,
                MAX(Nr_Ore_Seminar)            AS OreSem,
                MAX(Nr_Ore_Laborator)          AS OreLab,
                MAX(Nr_Ore_Proiect)            AS OreProj
            FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie]
            WHERE ID_AnUniv = @idAn
            GROUP BY
                ID_Profesor, Denumire, NrSemestruDinAn,
                TitularSauSuplinitor, ID_PlanMaterie_Prestator
        )";

        // =====================================================================
        // IDENTITATE CTE
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
                ON n.ID_TipGradDidactic = p.ID_TipGradDidacticAnUniv AND n.ID_AnUniv = @idAn
            WHERE p.ID_AnUnivCatedra = @idAn
        )";

        // =====================================================================
        // FILTRU TITULARI: IN (1,2,3,10,11)
        // Fara gradul 4 (Asistent) → ~711 titulari
        // Fara EXISTS ore → include si profesorii fara ore alocate
        // =====================================================================
        private const string FiltruTitulariOficiali = @"
              AND i.TitularAnUniv = 1
              AND i.ID_Facultate != 41
              AND i.ID_TipGrad IN (1,2,3,10,11)";

        private const string FiltruIdent = @"
              AND (@idFac = 0 OR i.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR i.ID_Catedra = @idCatedra)";

        private void AddCore(SqlCommand cmd, int idAn, int idFac, int idCat)
        {
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFac);
            cmd.Parameters.AddWithValue("@idCatedra", idCat);
        }

        private void AddExcelFilters(IXLWorksheet ws, int colCount,
            int idAn, string? fac, string? dept, string? prof,
            string? tipPost, string? formaInv, string? spec, int sem)
        {
            var parts = new List<string> { $"An: {idAn}" };
            if (!string.IsNullOrWhiteSpace(fac) && fac != "0" && fac != "Toate") parts.Add($"Facultate: {fac}");
            if (!string.IsNullOrWhiteSpace(dept) && dept != "0" && dept != "Toate") parts.Add($"Dept: {dept}");
            if (!string.IsNullOrWhiteSpace(prof) && prof != "Toti") parts.Add($"Profesor: {prof}");
            if (!string.IsNullOrWhiteSpace(tipPost) && tipPost != "Toti") parts.Add($"Tip: {tipPost}");
            if (!string.IsNullOrWhiteSpace(formaInv) && formaInv != "Toti") parts.Add($"Forma: {formaInv}");
            if (!string.IsNullOrWhiteSpace(spec) && spec != "Toti") parts.Add($"Spec: {spec}");
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
            if (string.IsNullOrWhiteSpace(profesor) || profesor == "Toti") return $"{raport}.xlsx";
            var safe = string.Concat(profesor.Take(35).Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
            return $"{safe}_{raport}.xlsx";
        }

        // =====================================================================
        // LISTE DROPDOWN-URI
        // =====================================================================
        [HttpGet("liste/ani-universitari")]
        public IActionResult GetAni() => Ok(_cache.GetOrCreate("AniUniv_v6", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
            var lst = new List<object>();
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT au.ID_AnUniv, LTRIM(RTRIM(au.Denumire)) AS Denumire
                FROM [AGSIS].[dbo].[AnUniversitar] au
                INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm ON ppm.ID_AnUniv = au.ID_AnUniv
                GROUP BY au.ID_AnUniv, au.Denumire, au.Ordine ORDER BY au.Ordine DESC", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) lst.Add(new { id = Convert.ToInt32(r[0]), nume = r[1]?.ToString() ?? "" });
            return lst;
        }));

        [HttpGet("liste/facultati")]
        public IActionResult GetFacultati() => Ok(_cache.GetOrCreate("Fac_v7", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
            var lst = new List<object> { new { id = 0, nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT p.ID_Facultate, p.DenumireFacultate
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                WHERE p.ID_AnUnivCatedra = (
                    SELECT TOP 1 au.ID_AnUniv FROM [AGSIS].[dbo].[AnUniversitar] au
                    INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm2 ON ppm2.ID_AnUniv = au.ID_AnUniv
                    GROUP BY au.ID_AnUniv, au.Ordine ORDER BY au.Ordine DESC)
                  AND p.DenumireFacultate IS NOT NULL
                ORDER BY p.DenumireFacultate", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                lst.Add(new { id = Convert.ToInt32(r["ID_Facultate"]), nume = r["DenumireFacultate"]?.ToString() ?? "" });
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
                WHERE p.ID_AnUnivCatedra = @idAn AND p.DenumireCatedra IS NOT NULL
                  AND LTRIM(RTRIM(p.DenumireCatedra)) != '' AND (@idFac = 0 OR p.ID_Facultate = @idFac)
                ORDER BY p.DenumireCatedra", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                lst.Add(new { id = Convert.ToInt32(r["ID_Catedra"]), nume = r["DenumireCatedra"]?.ToString() ?? "" });
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
                WHERE ppm.ID_AnUniv = @idAn AND ppm.DenumireSpecializare IS NOT NULL
                  AND (@idFac = 0 OR p.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR p.ID_Catedra = @idCatedra)
                ORDER BY Spec", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var s = r["Spec"]?.ToString() ?? ""; if (!string.IsNullOrWhiteSpace(s)) lst.Add(new { id = s, nume = s }); }
            return Ok(lst);
        }

        [HttpGet("liste/cicluri-studii")]
        public IActionResult GetCicluri([FromQuery] int? idAnUniv)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            return Ok(_cache.GetOrCreate($"Cicluri_{idAn}", e =>
            {
                e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lst = new List<object> { new { id = "Toti", nume = "Toate" } };
                using var conn = new SqlConnection(_cs); conn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT DenumireCicluInv FROM [AGSIS].[dbo].[View_FDS]
                    WHERE ID_AnUniv = @idAn AND DenumireCicluInv IS NOT NULL ORDER BY DenumireCicluInv", conn);
                cmd.Parameters.AddWithValue("@idAn", idAn);
                using var r = cmd.ExecuteReader();
                while (r.Read()) { var v = r[0]?.ToString() ?? ""; if (!string.IsNullOrWhiteSpace(v)) lst.Add(new { id = v, nume = v }); }
                return lst;
            }));
        }

        [HttpGet("liste/profesori")]
        public IActionResult GetProfesori([FromQuery] int? idAnUniv,
            [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? specializare, [FromQuery] string? tipPost)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = "Toti", nume = "— Toți profesorii —" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT p.NumeIntreg
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                    ON ppm.ID_Profesor = p.ID_Profesor AND ppm.ID_AnUniv = p.ID_AnUnivCatedra
                WHERE p.ID_AnUnivCatedra = @idAn AND p.NumeIntreg IS NOT NULL
                  AND LTRIM(RTRIM(p.NumeIntreg)) != ''
                  AND (@idFac = 0 OR p.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR p.ID_Catedra = @idCatedra)
                  AND (@spec = N'Toti' OR
                       CASE WHEN CHARINDEX(N'+', ppm.DenumireSpecializare) > 0
                            THEN LEFT(ppm.DenumireSpecializare, CHARINDEX(N'+', ppm.DenumireSpecializare)-1)
                            ELSE ppm.DenumireSpecializare END
                       COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
                  AND (@tipPost = N'Toti' OR
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                ORDER BY p.NumeIntreg COLLATE Romanian_CI_AS", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.CommandTimeout = 30;
            using var r = cmd.ExecuteReader();
            while (r.Read()) { var n = r[0]?.ToString() ?? ""; if (!string.IsNullOrWhiteSpace(n)) lst.Add(new { id = n, nume = n }); }
            return Ok(lst);
        }

        // =====================================================================
        // RAPORT 1: NORMA PROFESORI
        // FIX: CicluInv via LEFT JOIN FdsCiclu (DISTINCT, fara multiplicare).
        //      Filtrul @ciclu este operational.
        // FIX: DedupCte nou eliminat multiplicarea cuplajelor.
        // =====================================================================
        private string SqlNorma() => $@"
            WITH {IdenCte}, {DedupCte},
            FdsCiclu AS (
                SELECT DISTINCT DenumireSpecializare, DenumireCicluInv, ID_AnUniv
                FROM [AGSIS].[dbo].[View_FDS] WHERE ID_AnUniv = @idAn AND id_metaspecializare > 0
            ),
            Baza AS (
                SELECT i.IdentificatorUnic, i.NumeIntreg,
                       i.DenumireFacultate AS FacAngajator, ppm.DeptOra AS DepartamentOra,
                       i.DenumireGradDidactic,
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END AS TipPost,
                       ppm.SpecCurata AS Specializare, ppm.DenumireMaterie,
                       ppm.NrSemestruDinAn AS Semestru, ppm.DenumireFormaInv,
                       ppm.OreConv, ppm.OreCurs, ppm.OreSem + ppm.OreLab + ppm.OreProj AS OreAplic,
                       ISNULL(fds.DenumireCicluInv, N'Nespecificat') AS CicluInv,
                       CASE WHEN ppm.ApartineDeCuplaj IS NOT NULL
                            THEN N'Cuplat cu: ' + ISNULL(STUFF((
                                SELECT DISTINCT N', ' + vdc.DenumireScurtaSpecializare
                                FROM [AGSIS].[pi].[View_DetaliereCuplaje] vdc
                                WHERE vdc.ID_AnUniv = @idAn
                                  AND vdc.ID_Cuplaj = (SELECT TOP 1 x.ID_Cuplaj
                                      FROM [AGSIS].[pi].[View_DetaliereCuplaje] x
                                      WHERE x.ID_PlanMaterie_Prestator = ppm.ID_PlanMaterie_Prestator
                                        AND x.ID_AnUniv = @idAn)
                                  AND vdc.DenumireScurtaSpecializare COLLATE DATABASE_DEFAULT
                                      != ppm.DenumireScurtaSpecializare COLLATE DATABASE_DEFAULT
                                FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N''),
                                N'incarcare comuna')
                            ELSE N'' END AS Mentiuni
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                LEFT JOIN FdsCiclu fds ON ppm.SpecCurata COLLATE DATABASE_DEFAULT
                                       = fds.DenumireSpecializare COLLATE DATABASE_DEFAULT
                WHERE 1=1 {FiltruIdent}
                  AND (@prof     = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@spec     = N'Toti' OR ppm.SpecCurata COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
                  AND (@tipPost  = N'Toti' OR CASE ppm.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
                  AND (@sem      = 0 OR ppm.NrSemestruDinAn = @sem)
                  AND (@formaInv = N'Toti' OR ppm.DenumireFormaInv COLLATE DATABASE_DEFAULT = @formaInv COLLATE DATABASE_DEFAULT)
                  AND (@ciclu    = N'Toti' OR ISNULL(fds.DenumireCicluInv, N'Nespecificat') COLLATE DATABASE_DEFAULT = @ciclu COLLATE DATABASE_DEFAULT)
            )
            SELECT NumeIntreg, FacAngajator, DepartamentOra, DenumireGradDidactic,
                   TipPost, Specializare, CicluInv, DenumireMaterie, Semestru, DenumireFormaInv,
                   CAST(OreCurs AS DECIMAL(10,2)) AS OreCurs,
                   CAST(OreAplic AS DECIMAL(10,2)) AS OreAplic,
                   CAST(OreConv AS DECIMAL(10,2)) AS OreConv, Mentiuni
            FROM Baza
            ORDER BY NumeIntreg COLLATE Romanian_CI_AS, TipPost DESC, Specializare, Semestru";

        private void AddNormaParams(SqlCommand cmd, int idAn, int idFac, int idCat,
            string prof, string spec, string tipPost, int sem, string formaInv, string ciclu)
        {
            AddCore(cmd, idAn, idFac, idCat);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(prof) ? "Toti" : prof.Trim());
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(spec) ? "Toti" : spec.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@sem", sem);
            cmd.Parameters.AddWithValue("@formaInv", string.IsNullOrWhiteSpace(formaInv) ? "Toti" : formaInv.Trim());
            cmd.Parameters.AddWithValue("@ciclu", string.IsNullOrWhiteSpace(ciclu) ? "Toti" : ciclu.Trim());
        }

        [HttpGet("norma")]
        public async Task<IActionResult> GetNorma(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? specializare, [FromQuery] string? tipPost,
            [FromQuery] int? semestru, [FromQuery] string? formaInvatamant, [FromQuery] string? cicluStudii)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma(), conn); cmd.CommandTimeout = 120;
            AddNormaParams(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0,
                profesor ?? "", specializare ?? "", tipPost ?? "Toti",
                semestru ?? 0, formaInvatamant ?? "Toti", cicluStudii ?? "Toti");
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
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
                    CicluInv = r["CicluInv"]?.ToString() ?? "",
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
            [FromQuery] string? profesor, [FromQuery] string? specializare, [FromQuery] string? tipPost,
            [FromQuery] int? semestru, [FromQuery] string? formaInvatamant, [FromQuery] string? cicluStudii,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"), new DataColumn("Dept. Ora"),
                new DataColumn("Specializare"), new DataColumn("Ciclu"), new DataColumn("Materie"),
                new DataColumn("Tip Post"), new DataColumn("Sem.", typeof(int)), new DataColumn("Forma Inv."),
                new DataColumn("Ore Curs", typeof(decimal)), new DataColumn("Ore Aplic.", typeof(decimal)),
                new DataColumn("Ore Conv.", typeof(decimal)), new DataColumn("Mentiuni")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma(), conn); cmd.CommandTimeout = 120;
            AddNormaParams(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0,
                profesor ?? "", specializare ?? "", tipPost ?? "Toti",
                semestru ?? 0, formaInvatamant ?? "Toti", cicluStudii ?? "Toti");
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["DepartamentOra"]?.ToString(),
                    r["Specializare"]?.ToString(), r["CicluInv"]?.ToString(), r["DenumireMaterie"]?.ToString(),
                    r["TipPost"]?.ToString(), r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["DenumireFormaInv"]?.ToString(),
                    r["OreCurs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreCurs"]),
                    r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
                    r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]),
                    r["Mentiuni"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Norme");
            ws.Cell(1, 1).Value = "Detaliere Norme Profesori"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, tipPost, formaInvatamant, specializare, semestru ?? 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var c in new[] { "Ore Curs", "Ore Aplic.", "Ore Conv." }) tbl.Field(c).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL"; StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count));
            ws.Columns(1, 12).AdjustToContents(); ws.Column(13).Width = 60; ws.Column(13).Style.Alignment.WrapText = true;
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Norme", profesor));
        }

        // =====================================================================
        // RAPORT 2: TOTALURI NORME
        // FIX: PIVOT in SQL → returneaza OreIF, OreID, OreIFR pe un singur rand.
        //      Nu mai e nevoie de ZIP separat cu un fisier per forma.
        //      Filtrul @formaInvVal operational.
        // =====================================================================
        private string SqlTotaluri() => $@"
            WITH {IdenCte}, {DedupCte},
            Baza AS (
                SELECT i.IdentificatorUnic, i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra,
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END AS TipPost,
                       ppm.DenumireFormaInv, ppm.DenumireMaterie, ppm.NrSemestruDinAn,
                       ppm.OreConv, ppm.ID_PlanMaterie_Prestator
                FROM PpmDedup ppm INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE 1=1 {FiltruIdent}
                  AND (@prof    = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@tipPost = N'Toti' OR CASE ppm.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
            )
            SELECT
                NumeIntreg, MAX(DenumireFacultate) AS Facultate, MAX(DenumireCatedra) AS Departament, TipPost,
                -- PIVOT inline: o coloana per forma de invatamant
                CAST(SUM(CASE WHEN DenumireFormaInv = N'Cu frecvență'              THEN OreConv ELSE 0 END) AS DECIMAL(10,2)) AS OreIF,
                CAST(SUM(CASE WHEN DenumireFormaInv = N'Frecvență redusă'          THEN OreConv ELSE 0 END) AS DECIMAL(10,2)) AS OreIFR,
                CAST(SUM(CASE WHEN DenumireFormaInv = N'Învățământ la distanță'    THEN OreConv ELSE 0 END) AS DECIMAL(10,2)) AS OreID,
                CAST(SUM(OreConv) AS DECIMAL(10,2)) AS TotalOreConv,
                CAST(SUM(OreConv) * 14 AS DECIMAL(10,2)) AS TotalAnual
            FROM Baza
            GROUP BY IdentificatorUnic, NumeIntreg, TipPost
            HAVING SUM(OreConv) > 0 OR @prof != N'Toti'
            ORDER BY NumeIntreg COLLATE Romanian_CI_AS, TipPost DESC";

        [HttpGet("norma-totaluri")]
        public async Task<IActionResult> GetNormaTotaluri(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri(), conn); cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    TipPost = r["TipPost"]?.ToString() ?? "",
                    OreIF = r["OreIF"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIF"]),
                    OreIFR = r["OreIFR"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIFR"]),
                    OreID = r["OreID"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreID"]),
                    TotalOreConv = r["TotalOreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreConv"]),
                    TotalAnual = r["TotalAnual"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalAnual"])
                });
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public async Task<IActionResult> ExportNormaTotaluriZip(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost,
            [FromQuery] string? formaInvatamant,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Tip Post"),
                new DataColumn("Ore IF",  typeof(decimal)), new DataColumn("Ore IFR", typeof(decimal)),
                new DataColumn("Ore ID",  typeof(decimal)),
                new DataColumn("Total Ore Conv.",  typeof(decimal)),
                new DataColumn("Total Anual (x14)",typeof(decimal))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri(), conn); cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor!.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost!.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
            {
                decimal oi = r["OreIF"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIF"]);
                decimal oifr = r["OreIFR"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIFR"]);
                decimal oid = r["OreID"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreID"]);
                // Daca e filtru pe forma, skip randuri cu 0 la acea forma
                bool include = string.IsNullOrWhiteSpace(formaInvatamant) || formaInvatamant == "Toti"
                    || (formaInvatamant.Contains("frecvență") && oi > 0)
                    || (formaInvatamant.Contains("redusă") && oifr > 0)
                    || (formaInvatamant.Contains("distanță") && oid > 0);
                if (!include) continue;
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(), r["TipPost"]?.ToString(), oi, oifr, oid,
                    r["TotalOreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreConv"]),
                    r["TotalAnual"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalAnual"]));
            }
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Totaluri");
            ws.Cell(1, 1).Value = $"Totaluri Norme | An: {idAn}"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, tipPost, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var c in new[] { "Ore IF", "Ore IFR", "Ore ID", "Total Ore Conv.", "Total Anual (x14)" })
                tbl.Field(c).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL"; StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Totaluri_Norme.xlsx");
        }

        // =====================================================================
        // RAPORT 3: DISTRIBUTIE ORE PE PROGRAME
        // FIX: adaugat @spec in WHERE — filtrul era ignorat anterior.
        // =====================================================================
        private string SqlDistrib() => $@"
            WITH {IdenCte}, {DedupCte},
            OrePerProg AS (
                SELECT i.IdentificatorUnic, i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra,
                       ppm.SpecCurata AS Program, SUM(ppm.OreConv) AS OreProgram
                FROM PpmDedup ppm INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE 1=1 {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@spec = N'Toti' OR ppm.SpecCurata COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
                GROUP BY i.IdentificatorUnic, i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra, ppm.SpecCurata
            ),
            TotUniv AS (SELECT IdentificatorUnic, SUM(OreProgram) AS TotalOreUniv FROM OrePerProg GROUP BY IdentificatorUnic)
            SELECT o.NumeIntreg, o.DenumireFacultate, o.DenumireCatedra, o.Program,
                   CAST(o.OreProgram AS DECIMAL(10,2)) AS OreProgram,
                   CAST(t.TotalOreUniv AS DECIMAL(10,2)) AS TotalOreUniv,
                   CAST(CASE WHEN t.TotalOreUniv > 0 THEN ROUND(o.OreProgram/t.TotalOreUniv*100,2) ELSE 0 END AS DECIMAL(10,2)) AS Procent
            FROM OrePerProg o INNER JOIN TotUniv t ON t.IdentificatorUnic = o.IdentificatorUnic
            ORDER BY o.NumeIntreg COLLATE Romanian_CI_AS, Procent DESC";

        [HttpGet("distributie-ore")]
        public async Task<IActionResult> GetDistrib(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor, [FromQuery] string? specializare)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDistrib(), conn); cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
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
            [FromQuery] int? idCatedra, [FromQuery] string? profesor, [FromQuery] string? specializare,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Program Studiu"),
                new DataColumn("Ore Program", typeof(decimal)), new DataColumn("Total Ore Univ.", typeof(decimal)),
                new DataColumn("Procent %", typeof(decimal))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDistrib(), conn); cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["DenumireFacultate"]?.ToString(),
                    r["DenumireCatedra"]?.ToString(), r["Program"]?.ToString(),
                    r["OreProgram"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreProgram"]),
                    r["TotalOreUniv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreUniv"]),
                    r["Procent"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Procent"]));
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Distributie Ore");
            ws.Cell(1, 1).Value = "Distributie Ore pe Programe"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, null, null, specializare, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Program").TotalsRowFunction = XLTotalsRowFunction.Sum; tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Distributie_Ore", profesor));
        }

        // =====================================================================
        // RAPORT 4: LIMBI STRAINE
        // FIX MAJOR: sursa de adevar este LimbaPredare din View_FDS (atribut
        //   al programului), NU LimbaDePredare din Post_Profesor_Materie
        //   (atribut al orei individuale).
        //   LimbaPredare in View_FDS are valorile: 'EN', 'G', 'FR' etc.
        //   Filtrul: fds.LimbaPredare IS NOT NULL AND fds.LimbaPredare != 'RO'
        //   (sau echivalent) + id_metaspecializare > 0 pentru a evita rors.
        // =====================================================================
        private string SqlLimbi() => $@"
            WITH {IdenCte}, {DedupCte},
            -- Programe oficiale in limbi straine: LimbaPredare din View_FDS
            ProgLimbaStrain AS (
                SELECT DISTINCT DenumireSpecializare, LimbaPredare, DenumireCicluInv, ID_AnUniv
                FROM [AGSIS].[dbo].[View_FDS]
                WHERE ID_AnUniv = @idAn
                  AND id_metaspecializare > 0
                  AND LimbaPredare IS NOT NULL
                  AND LTRIM(RTRIM(LimbaPredare)) NOT IN (N'', N'RO', N'Ro', N'ro', N'Romana', N'Română')
            ),
            Filtrat AS (
                SELECT i.NumeIntreg, ppm.NrSemestruDinAn, ppm.OreConv,
                       pls.LimbaPredare AS LimbaProgram,
                       ppm.SpecCurata AS ProgramStudiu,
                       ISNULL(pls.DenumireCicluInv, N'Neidentificat') AS CicluStudii
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                -- JOIN pe LimbaPredare din FDS (limbA programului, nu a orei)
                INNER JOIN ProgLimbaStrain pls
                    ON ppm.SpecCurata COLLATE DATABASE_DEFAULT
                       = pls.DenumireSpecializare COLLATE DATABASE_DEFAULT
                WHERE 1=1 {FiltruIdent}
                  AND (@prof    = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@tipPost = N'Toti' OR CASE ppm.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
                  AND (@ciclu   = N'Toti' OR ISNULL(pls.DenumireCicluInv, N'Neidentificat') COLLATE DATABASE_DEFAULT = @ciclu COLLATE DATABASE_DEFAULT)
                  AND ppm.OreConv > 0
            ),
            Dedup2 AS (
                SELECT NumeIntreg, NrSemestruDinAn, LimbaProgram, ProgramStudiu, CicluStudii,
                       MAX(OreConv) AS OreD
                FROM Filtrat
                GROUP BY NumeIntreg, NrSemestruDinAn, LimbaProgram, ProgramStudiu, CicluStudii
            )
            SELECT NumeIntreg, NrSemestruDinAn AS Semestru,
                   CAST(SUM(OreD) AS DECIMAL(10,2)) AS TotalOre,
                   -- LimbaProgram = limba specializarii (EN, G, FR etc.)
                   MAX(LimbaProgram) AS LimbaProgram,
                   ProgramStudiu, CicluStudii
            FROM Dedup2
            GROUP BY NumeIntreg, NrSemestruDinAn, ProgramStudiu, CicluStudii
            HAVING SUM(OreD) > 0
            ORDER BY NumeIntreg COLLATE Romanian_CI_AS, NrSemestruDinAn";

        [HttpGet("limbi-straine")]
        public async Task<IActionResult> GetLimbi(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost, [FromQuery] string? cicluStudii)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi(), conn); cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@ciclu", string.IsNullOrWhiteSpace(cicluStudii) ? "Toti" : cicluStudii.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Semestru = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    TotalOre = r["TotalOre"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOre"]),
                    LimbaProgram = r["LimbaProgram"]?.ToString() ?? "",
                    ProgramStudiu = r["ProgramStudiu"]?.ToString() ?? "",
                    CicluStudii = r["CicluStudii"]?.ToString() ?? ""
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
                new DataColumn("Semestru", typeof(int)), new DataColumn("Total Ore (h)", typeof(decimal)),
                new DataColumn("Limba Program"), new DataColumn("Program Studiu"), new DataColumn("Ciclu Studii")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi(), conn); cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@ciclu", string.IsNullOrWhiteSpace(cicluStudii) ? "Toti" : cicluStudii.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(),
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["TotalOre"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOre"]),
                    r["LimbaProgram"]?.ToString(), r["ProgramStudiu"]?.ToString(), r["CicluStudii"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Limbi Straine");
            ws.Cell(1, 1).Value = "Raport Limbi Straine"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, tipPost, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Total Ore (h)").TotalsRowFunction = XLTotalsRowFunction.Sum; tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Limbi_Straine", profesor));
        }

        // =====================================================================
        // RAPORT 5: DISCIPLINE PER PROFESOR
        // FIX: filtru @sem operational prin NrSemestruDinAn.
        // FIX: normalizare whitespace in DISTINCT_Mat.
        // FIX: ZIP respecta filtrul formaInvatamant (genereaza doar fisierul cerut).
        // =====================================================================
        private string SqlDisc(string? formaInvFilter = null) => $@"
            WITH {IdenCte}, {DedupCte},
            Distinct_Mat AS (
                SELECT DISTINCT i.IdentificatorUnic, i.NumeIntreg,
                    i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic,
                    -- Normalizare whitespace pe materie
                    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireMaterie, N'  ', N'><'), N'<>', N''), N'><', N' '))) AS DenumireMaterie,
                    ppm.DenumireFormaInv
                FROM PpmDedup ppm INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE 1=1 {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@sem  = 0       OR ppm.NrSemestruDinAn = @sem)
                  AND ppm.DenumireMaterie IS NOT NULL AND LTRIM(RTRIM(ppm.DenumireMaterie)) != N''
                  {(string.IsNullOrWhiteSpace(formaInvFilter) ? "" :
                    "AND ppm.DenumireFormaInv COLLATE DATABASE_DEFAULT = @formaInvFilter COLLATE DATABASE_DEFAULT")}
            )
            SELECT dm.NumeIntreg, MAX(dm.DenumireFacultate) AS Facultate,
                MAX(dm.DenumireCatedra) AS Departament, MAX(dm.DenumireGradDidactic) AS Grad,
                STUFF((SELECT DISTINCT N' | ' + dm2.DenumireMaterie FROM Distinct_Mat dm2
                       WHERE dm2.IdentificatorUnic = dm.IdentificatorUnic AND dm2.DenumireMaterie IS NOT NULL
                       FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 3, N'') AS Discipline,
                COUNT(DISTINCT dm.DenumireMaterie) AS NrDisc
            FROM Distinct_Mat dm GROUP BY dm.IdentificatorUnic, dm.NumeIntreg
            ORDER BY dm.NumeIntreg COLLATE Romanian_CI_AS";

        [HttpGet("discipline")]
        public async Task<IActionResult> GetDisc([FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor, [FromQuery] int? semestru)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDisc(), conn); cmd.CommandTimeout = 180;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
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
        public async Task<IActionResult> ExportDisciplineZip([FromQuery] int? idAnUniv,
            [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? formaInvatamant, [FromQuery] int? semestru)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            // Daca e specificata o forma, genereaza DOAR acel fisier in ZIP
            var forme = new List<(string Val, string Label)>();
            using (var connF = new SqlConnection(_cs))
            {
                await connF.OpenAsync();
                using var cmdF = new SqlCommand(@"SELECT DISTINCT DenumireFormaInv
                    FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie]
                    WHERE ID_AnUniv = @idAn AND DenumireFormaInv IS NOT NULL ORDER BY DenumireFormaInv", connF);
                cmdF.Parameters.AddWithValue("@idAn", idAn);
                using var rf = await cmdF.ExecuteReaderAsync();
                while (await rf.ReadAsync())
                {
                    var v = rf[0]?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    // Filtru: daca e specificata o forma, include numai aceea
                    if (!string.IsNullOrWhiteSpace(formaInvatamant) && formaInvatamant != "Toti"
                        && !v.Equals(formaInvatamant, StringComparison.OrdinalIgnoreCase)) continue;
                    var lbl = v.Length > 20 ? v.Substring(0, 20).Trim().Replace(" ", "_") : v.Replace(" ", "_");
                    forme.Add((v, lbl));
                }
            }
            using var memZip = new MemoryStream();
            using (var archive = new ZipArchive(memZip, ZipArchiveMode.Create, true))
            {
                foreach (var (formaVal, formaLabel) in forme)
                {
                    var dt = new DataTable();
                    dt.Columns.AddRange(new[]{
                        new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                        new DataColumn("Facultate"), new DataColumn("Departament"),
                        new DataColumn("Grad"), new DataColumn("Discipline"), new DataColumn("Nr. Disc.", typeof(int))
                    });
                    using var conn = new SqlConnection(_cs); await conn.OpenAsync();
                    using var cmd = new SqlCommand(SqlDisc(formaVal), conn); cmd.CommandTimeout = 180;
                    AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
                    cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor!.Trim());
                    cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
                    cmd.Parameters.AddWithValue("@formaInvFilter", formaVal);
                    using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
                    while (await r.ReadAsync())
                        dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["Facultate"]?.ToString(),
                            r["Departament"]?.ToString(), r["Grad"]?.ToString(), r["Discipline"]?.ToString(),
                            r["NrDisc"] == DBNull.Value ? 0 : Convert.ToInt32(r["NrDisc"]));
                    var entry = archive.CreateEntry($"Discipline_{formaLabel}.xlsx");
                    using var es = entry.Open();
                    using var wb2 = new XLWorkbook(); var ws2 = wb2.Worksheets.Add("Disc");
                    ws2.Cell(1, 1).Value = $"Discipline predate — {formaVal} | An: {idAn}";
                    ws2.Cell(1, 1).Style.Font.Bold = true; ws2.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
                    ws2.Range(1, 1, 1, dt.Columns.Count).Merge();
                    var tbl2 = ws2.Cell(3, 1).InsertTable(dt); tbl2.Theme = XLTableTheme.None;
                    StyleHdr(ws2.Range(3, 1, 3, dt.Columns.Count));
                    ws2.Columns(1, 5).AdjustToContents(); ws2.Column(6).Width = 80; ws2.Column(6).Style.Alignment.WrapText = true;
                    using var wbs = new MemoryStream(); wb2.SaveAs(wbs); wbs.Position = 0; wbs.CopyTo(es);
                }
            }
            memZip.Position = 0;
            return File(memZip.ToArray(), "application/zip", "Discipline_Predate.zip");
        }

        // =====================================================================
        // RAPORT 6: TITULARI OFICIALI
        // FIX 37 PERSOANE LIPSA: Scos INNER JOIN implicit cu Post_Profesor_Materie.
        //   Sursa unica: View_Profesori_CF_AnUniv filtrat pe TitularAnUniv=1.
        //   Profesori fara ore alocate (concediu, sabatic, maternitate) inclusi.
        //
        // FIX GRAD: ROW_NUMBER() OVER(PARTITION BY IdentificatorUnic ORDER BY
        //   ID_TipGrad ASC) — gradul cel mai mic numeric = cel mai mare ierarhic.
        //   Profesor(1) > Conferentiar(2) > Lector(3) > Sef Lucrari(10) > Lector(11).
        // =====================================================================
        private string SqlTitulari() => $@"
            WITH {IdenCte},
            TitRanked AS (
                SELECT *, ROW_NUMBER() OVER(PARTITION BY IdentificatorUnic ORDER BY ID_TipGrad ASC) AS Rn
                FROM Identitate i
                WHERE 1=1 {FiltruTitulariOficiali} {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
            )
            SELECT NumeIntreg, DenumireFacultate, DenumireCatedra, DenumireGradDidactic
            FROM TitRanked WHERE Rn = 1
            ORDER BY NumeIntreg COLLATE Romanian_CI_AS";

        [HttpGet("titulari")]
        public async Task<IActionResult> GetTitulari([FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulari(), conn); cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
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
        public async Task<IActionResult> ExportTitulari([FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad Didactic")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulari(), conn); cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["DenumireFacultate"]?.ToString(),
                    r["DenumireCatedra"]?.ToString(), r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Titulari");
            ws.Cell(1, 1).Value = "Cadre Didactice Titulare (Oficiali)"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, null, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowFunction = XLTotalsRowFunction.Count; tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Titulari", profesor));
        }

        // =====================================================================
        // RAPORT 7: COLABORATORI
        // FIX DUPLICATE: GROUP BY pe IdentificatorUnic.
        // Afiliieri: STUFF+FOR XML PATH pentru concatenare departamente multiple.
        // =====================================================================
        private string SqlColab() => $@"
            WITH {IdenCte},
            ColabBase AS (
                SELECT IdentificatorUnic, NumeIntreg, DenumireFacultate, DenumireCatedra, DenumireGradDidactic
                FROM Identitate i
                WHERE (i.TitularAnUniv = 0 OR i.TitularAnUniv IS NULL)
                  AND i.NumeIntreg IS NOT NULL AND LTRIM(RTRIM(i.NumeIntreg)) != N'' {FiltruIdent}
            ),
            ColabDedup AS (
                SELECT IdentificatorUnic, MAX(NumeIntreg) AS NumeIntreg,
                    -- Facultati distincte concatenate
                    STUFF((SELECT DISTINCT N' / ' + c2.DenumireFacultate
                           FROM ColabBase c2 WHERE c2.IdentificatorUnic = ColabBase.IdentificatorUnic
                             AND c2.DenumireFacultate IS NOT NULL
                           FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 3, N'') AS DenumireFacultate,
                    -- Departamente distincte concatenate
                    STUFF((SELECT DISTINCT N' / ' + c3.DenumireCatedra
                           FROM ColabBase c3 WHERE c3.IdentificatorUnic = ColabBase.IdentificatorUnic
                             AND c3.DenumireCatedra IS NOT NULL
                           FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 3, N'') AS DenumireCatedra,
                    MAX(DenumireGradDidactic) AS DenumireGradDidactic
                FROM ColabBase GROUP BY IdentificatorUnic
            )
            SELECT NumeIntreg, DenumireFacultate, DenumireCatedra, DenumireGradDidactic
            FROM ColabDedup ORDER BY NumeIntreg COLLATE Romanian_CI_AS";

        [HttpGet("colaboratori")]
        public async Task<IActionResult> GetColab([FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColab(), conn); cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
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
        public async Task<IActionResult> ExportColab([FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate(i)"), new DataColumn("Departament(e)"), new DataColumn("Grad Didactic")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColab(), conn); cmd.CommandTimeout = 60;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["DenumireFacultate"]?.ToString(),
                    r["DenumireCatedra"]?.ToString(), r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Colaboratori");
            ws.Cell(1, 1).Value = "Asociati / Colaboratori"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, null, null, null, null, 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowFunction = XLTotalsRowFunction.Count; tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Colaboratori.xlsx");
        }

        // =====================================================================
        // RAPORT 8: ANS
        //
        // FIX 1 — SNAPSHOT SEMESTRUL 2:
        //   Raportul oficial ANS = snapshot semestrul 2.
        //   Dovada: Voinea Mihaela oficial=1.01, generat=2.35 (aduna S1+S2).
        //   Filtru: AND vc.NrSemestruDinAn = 2
        //
        // FIX 2 — FORMULA CORECTA (confirmat cu Abrudan):
        //   SUM(NrOreConventionale) S2 / Norma = fractie
        //   Abrudan S2: 13.21 / 11 ≈ 1.20 → dar oficial e 0.88.
        //   Oficial Abrudan: 0.43 (Silvicultura) + 0.45 (Ing. resurse) = 0.88.
        //   9.745 total an / 11 = 0.886 ≈ 0.89 (total an e corect pentru 0.88).
        //   Concluzie: sursa oficiala foloseste TOTALUL ANUAL (S1+S2) / Norma.
        //   Folosim deci SUM fara filtru de semestru, dar NUMAI NrCrtPostProfesor=1.
        //
        // FIX 3 — ORDONARE COLOANE:
        //   Oficial: Matematica (col.1), Informatica (col.2), Fizica (col.3)...
        //   In DB: Matematica ID=1, Informatica ID=40 → fara reordonare ar fi separate.
        //   FIX C#: dupa LoadDomeniiAsync(), mutam Informatica (ID=40) pe pozitia 1
        //   (imediat dupa Matematica ID=1), lasand restul in ordinea naturala.
        //
        // FIX 4 — 741 TITULARI (INNER JOIN eliminat):
        //   SqlTitulariANS nu mai filtreaza pe existenta orelor.
        //   Profesori fara ore primesc fracties goale dar apar in lista.
        // =====================================================================
        private const string SqlCreateTempDomeniu = @"
            IF OBJECT_ID('tempdb..#DomeniuANS') IS NOT NULL DROP TABLE #DomeniuANS;
            CREATE TABLE #DomeniuANS (
                ID_ELEMENT INT, COD_DS_CNATDCU NVARCHAR(20), cod_DS NVARCHAR(20),
                ID_RamuraDeStiinta_ANS INT, DomeniulDeStudiu_ANS NVARCHAR(200),
                RamuraDeStiinta_ANS NVARCHAR(200), DomeniuFundamental NVARCHAR(200)
            );
            INSERT INTO #DomeniuANS EXEC [AGSIS].[dbo].[N_DOMENIU_STUDIUL_ANS_List];";

        // NrCrtPostProfesor = 1 = post de baza. Fara filtru semestru (total anual).
        private const string SqlOreANS = @"
            SELECT vc.ID_Profesor, Map.ID_Ramura AS ID_ANS,
                   CAST(SUM(vc.NrOreConventionale) AS DECIMAL(18,6)) AS OreConvTotal
            FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vc
            CROSS APPLY (
                SELECT TOP 1 d.ID_RamuraDeStiinta_ANS AS ID_Ramura
                FROM [AGSIS].[dbo].[View_FDS] fds
                INNER JOIN #DomeniuANS d ON d.ID_ELEMENT = fds.ID_N_Domeniu_Studiu_ANS
                WHERE fds.DenumireSpecializare COLLATE DATABASE_DEFAULT
                      = vc.DenumireSpecializare COLLATE DATABASE_DEFAULT
                  AND fds.ID_AnUniv = @idAn AND fds.id_metaspecializare > 0
                  AND fds.ID_N_Domeniu_Studiu_ANS IS NOT NULL
            ) Map
            WHERE vc.ID_AnUniv = @idAn AND vc.NrCrtPostProfesor = 1
            GROUP BY vc.ID_Profesor, Map.ID_Ramura
            HAVING SUM(vc.NrOreConventionale) > 0";

        // Sursa: View_Profesori_CF_AnUniv — NU mai e filtru EXISTS pe ore
        private const string SqlTitulariANS = @"
            SELECT DISTINCT p.ID_Profesor, p.NumeIntreg, p.DenumireFacultate, p.DenumireCatedra,
                p.DenumireGradDidactic, p.ID_TipGradDidacticAnUniv AS ID_TipGrad,
                ISNULL(n.NrOreConventionaleTitular, 0)  AS NormaDB,
                ISNULL(n.NrOreConventionaleSuplinitor,0) AS NormaSupl, p.CNP,
                CASE WHEN LEN(LTRIM(RTRIM(ISNULL(p.CNP,'')))) = 13
                          AND ISNUMERIC(SUBSTRING(p.CNP,1,3)) = 1
                     THEN 2026 - (CASE WHEN SUBSTRING(p.CNP,1,1) IN ('1','2') THEN 1900 ELSE 2000 END
                                  + CAST(SUBSTRING(p.CNP,2,2) AS INT))
                     ELSE NULL END AS Varsta,
                CASE WHEN p.TitularAnUniv = 1
                          AND p.DenumireGradDidactic NOT LIKE N'%Asociat%'
                          AND p.DenumireGradDidactic NOT LIKE N'%asociat%' THEN N'1'
                     WHEN p.DenumireGradDidactic LIKE N'%Asociat%'
                       OR p.DenumireGradDidactic LIKE N'%asociat%'          THEN N'2'
                     ELSE N'3' END AS FormaAngajare
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            LEFT JOIN [AGSIS].[pi].[NormaOreConventionale] n
                ON n.ID_TipGradDidactic = p.ID_TipGradDidacticAnUniv AND n.ID_AnUniv = @idAn
            WHERE p.ID_AnUnivCatedra = @idAn AND p.TitularAnUniv = 1
              AND p.ID_Facultate != 41 AND p.ID_TipGradDidacticAnUniv IN (1,2,3,10,11)
              AND (@idFac = 0 OR p.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR p.ID_Catedra = @idCatedra)
            ORDER BY p.NumeIntreg COLLATE Romanian_CI_AS";

        private class TitAns
        {
            public int IdProf { get; set; }
            public string Nume { get; set; } = "";
            public string Fac { get; set; } = "";
            public string Dept { get; set; } = "";
            public string GradD { get; set; } = "";
            public int? IdTip { get; set; }
            public decimal Norma { get; set; }
            public string Cnp { get; set; } = "";
            public int? Varsta { get; set; }
            public string FormaAngajare { get; set; } = "3";
        }

        private async Task<List<(int Id, string Den)>> LoadDomeniiAsync(SqlConnection conn)
        {
            var raw = new List<(int Id, string Den)>();
            using var cmd = new SqlCommand(
                "SELECT ID_Element, Denumire FROM [AGSIS].[dbo].[N_RAMURA_STIINTA_ANS] ORDER BY ID_Element", conn);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                raw.Add((Convert.ToInt32(r["ID_Element"]), r["Denumire"]?.ToString() ?? ""));

            // FIX ORDONARE: oficial Matematica=col1, Informatica=col2, restul in ordine ID
            // ID=1 Matematică, ID=40 Informatică → mutam ID=40 pe pozitia 1
            var reordonat = new List<(int Id, string Den)>();
            var mat = raw.FirstOrDefault(d => d.Id == 1);
            var info = raw.FirstOrDefault(d => d.Id == 40);
            if (mat.Den != null) reordonat.Add(mat);
            if (info.Den != null) reordonat.Add(info);
            foreach (var d in raw)
                if (d.Id != 1 && d.Id != 40) reordonat.Add(d);
            return reordonat;
        }

        private async Task<List<TitAns>> LoadTitulariANSAsync(SqlConnection conn, int idAn, int idFac, int idCat)
        {
            var lst = new List<TitAns>();
            using var cmd = new SqlCommand(SqlTitulariANS, conn); cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFac);
            cmd.Parameters.AddWithValue("@idCatedra", idCat);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                decimal nDb = r["NormaDB"] == DBNull.Value ? 0m : Convert.ToDecimal(r["NormaDB"]);
                decimal nSupl = r["NormaSupl"] == DBNull.Value ? 0m : Convert.ToDecimal(r["NormaSupl"]);
                lst.Add(new TitAns
                {
                    IdProf = Convert.ToInt32(r["ID_Profesor"]),
                    Nume = r["NumeIntreg"]?.ToString() ?? "",
                    Fac = r["DenumireFacultate"]?.ToString() ?? "",
                    Dept = r["DenumireCatedra"]?.ToString() ?? "",
                    GradD = r["DenumireGradDidactic"]?.ToString() ?? "",
                    IdTip = r["ID_TipGrad"] == DBNull.Value ? null : (int?)Convert.ToInt32(r["ID_TipGrad"]),
                    Norma = nDb > 0 ? nDb : nSupl,
                    Cnp = r["CNP"]?.ToString() ?? "",
                    Varsta = r["Varsta"] == DBNull.Value ? null : (int?)Convert.ToInt32(r["Varsta"]),
                    FormaAngajare = r["FormaAngajare"]?.ToString() ?? "3"
                });
            }
            return lst;
        }

        private async Task<Dictionary<int, Dictionary<int, decimal>>> LoadOreANSAsync(SqlConnection conn, int idAn)
        {
            var dict = new Dictionary<int, Dictionary<int, decimal>>();
            using (var cmdT = new SqlCommand(SqlCreateTempDomeniu, conn)) { cmdT.CommandTimeout = 60; await cmdT.ExecuteNonQueryAsync(); }
            using var cmd = new SqlCommand(SqlOreANS, conn); cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                int idP = Convert.ToInt32(r["ID_Profesor"]);
                int idRam = Convert.ToInt32(r["ID_ANS"]);
                decimal ore = r["OreConvTotal"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConvTotal"]);
                if (!dict.ContainsKey(idP)) dict[idP] = new();
                if (!dict[idP].ContainsKey(idRam)) dict[idP][idRam] = 0m;
                dict[idP][idRam] += ore;
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
                if (oreAns.TryGetValue(t.IdProf, out var oreP) && t.Norma > 0)
                    foreach (var kv in oreP)
                    {
                        var dom = domenii.FirstOrDefault(d => d.Id == kv.Key);
                        if (dom.Den == null) continue;
                        decimal f = Math.Round(kv.Value / t.Norma, 2);
                        if (f > 0) frac[dom.Den] = f;
                    }
                profesori.Add(new
                {
                    NrCrt = nrCrt++,
                    NumeComplet = t.Nume,
                    CNP = t.Cnp,
                    Varsta = t.Varsta,
                    FormaAngajare = t.FormaAngajare,
                    GradFunctie = t.GradD,
                    Facultate = t.Fac,
                    Departament = t.Dept,
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
            var domenii = await LoadDomeniiAsync(conn);  // Matematica(0), Informatica(1), rest...
            var titulari = await LoadTitulariANSAsync(conn, idAn, idFacultate ?? 0, idCatedra ?? 0);
            var oreAns = await LoadOreANSAsync(conn, idAn);

            int nrD = domenii.Count, colTot = 9 + nrD + 1;
            var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("CD DRU");
            var gC = XLColor.FromHtml(Green);

            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, colTot).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov";
            ws.Range(3, 1, 3, 6).Merge();

            string[] antet = { "Nr.\nCrt.", "Nume si prenume", "CNP", "Functie",
                               "Forma\nangajare", "Cond.\ndoctorat", "Varsta", "Facultate", "Departament" };
            for (int c = 1; c <= 9; c++) { ws.Cell(5, c).Value = antet[c - 1]; ws.Range(5, c, 7, c).Merge(); }
            ws.Cell(5, colTot).Value = "Total"; ws.Range(5, colTot, 7, colTot).Merge();

            // Coloanele de ramuri — ordinea C#: Matematica(0), Informatica(1), apoi restul
            for (int i = 0; i < nrD; i++) { ws.Cell(6, 10 + i).Value = domenii[i].Den; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }
            for (int c = 1; c <= 9; c++) ws.Cell(8, c).Value = ((char)('A' + c - 1)).ToString();
            for (int i = 0; i < nrD; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, colTot).Value = nrD;

            for (int row = 5; row <= 8; row++)
                for (int col = 1; col <= colTot; col++)
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
                ws.Cell(row2, 3).Value = t.Cnp;
                ws.Cell(row2, 4).Value = t.GradD;
                ws.Cell(row2, 5).Value = t.FormaAngajare;
                ws.Cell(row2, 6).Value = "";
                if (t.Varsta.HasValue) ws.Cell(row2, 7).Value = t.Varsta.Value;
                else ws.Cell(row2, 7).Value = Blank.Value;
                ws.Cell(row2, 8).Value = t.Fac;
                ws.Cell(row2, 9).Value = t.Dept;

                decimal totFrac = 0m;
                if (oreAns.TryGetValue(t.IdProf, out var oreP) && t.Norma > 0)
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