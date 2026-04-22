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

        // =====================================================================
        // An universitar curent - dinamic din DB
        // =====================================================================
        private int GetAnCurent() => _cache.GetOrCreate("AnCurent_v3", e =>
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

        // =====================================================================
        // Norma legala fallback - gradele fara rand in NormaOreConventionale
        // (5=Preparator, 6=Prof.Consultant, 12-21=Cercetatori etc.)
        // =====================================================================
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
        // DEDUP CTE - elimina duplicatele JSONGrupe din PPM
        // Sursa: Post_Profesor_Materie (ppm)
        // =====================================================================
        private const string DedupCte = @"
        PpmDedup AS (
            SELECT
                ID_Profesor,
                DenumireSpecializare,
                CASE WHEN CHARINDEX(N'+', DenumireSpecializare) > 0
                     THEN LEFT(DenumireSpecializare, CHARINDEX(N'+', DenumireSpecializare) - 1)
                     ELSE DenumireSpecializare END                       AS SpecCurata,
                Denumire                                                 AS DenumireMaterie,
                NrSemestruDinAn,
                TitularSauSuplinitor,
                LimbaDePredare,
                DenumireFormaInv,
                DenumireFacultate                                        AS FacPPM,
                MAX(NrOreConventionale)                                  AS OreConv,
                MAX(Nr_Ore_Curs)                                         AS OreCurs,
                MAX(Nr_Ore_Seminar)                                      AS OreSem,
                MAX(Nr_Ore_Laborator)                                    AS OreLab,
                MAX(Nr_Ore_Proiect)                                      AS OreProj
            FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie]
            WHERE ID_AnUniv = @idAn
            GROUP BY
                ID_Profesor, DenumireSpecializare, Denumire,
                NrSemestruDinAn, TitularSauSuplinitor,
                LimbaDePredare, DenumireFormaInv, DenumireFacultate
        )";

        // =====================================================================
        // IDENTITATE CTE - View_Profesori_CF_AnUniv + norma din DB
        // =====================================================================
        private const string IdenCte = @"
        Identitate AS (
            SELECT
                p.ID_Profesor,
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))        AS IdentificatorUnic,
                p.NumeIntreg,
                p.ID_Facultate,
                p.DenumireFacultate,
                p.ID_Catedra,
                p.DenumireCatedra,
                p.TitularAnUniv,
                p.ID_TipGradDidacticAnUniv                               AS ID_TipGrad,
                p.DenumireGradDidactic,
                ISNULL(n.NrOreConventionaleTitular, 0)                   AS NormaDB,
                ISNULL(n.NrOreConventionaleSuplinitor, 18)               AS NormaSupl
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            LEFT JOIN [AGSIS].[pi].[NormaOreConventionale] n
                ON n.ID_TipGradDidactic = p.ID_TipGradDidacticAnUniv
                AND n.ID_AnUniv = @idAn
            WHERE p.ID_AnUnivCatedra = @idAn
        )";

        // =====================================================================
        // REGULA ANS 741 - filtrul complet pentru titularii oficiali ANS
        // Conditii:
        //   p.TitularAnUniv = 1          - angajat titular
        //   p.ID_Facultate != 41         - exclude DPPD
        //   p.ID_TipGradDidacticAnUniv IN (1,2,3,4,10,11) - grade didactice cu norma
        //   SUM(vcm.NrOreConventionale) > 0 - activitate reala
        //   p.CNP IS NOT NULL            - identitate unica verificabila
        // =====================================================================
        private const string FiltruAns741 = @"
            WHERE i.TitularAnUniv = 1
              AND i.ID_Facultate != 41
              AND i.ID_TipGrad IN (1,2,3,4,10,11)
              AND i.CNP IS NOT NULL
              AND (@idFac = 0 OR i.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR i.ID_Catedra = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                  WHERE vcm.ID_Profesor = i.ID_Profesor
                    AND vcm.ID_AnUniv = @idAn
                  HAVING SUM(vcm.NrOreConventionale) > 0
              )";

        // =====================================================================
        // Filtru comun rapoarte generale (toti profesorii, nu doar ANS)
        // =====================================================================
        private const string FiltruIdent = @"
              AND (@idFac = 0 OR i.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR i.ID_Catedra = @idCatedra)";

        private void AddCore(SqlCommand cmd, int idAn, int idFac, int idCat)
        {
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFac);
            cmd.Parameters.AddWithValue("@idCatedra", idCat);
        }

        private void StyleHdr(IXLRange r)
        {
            r.Style.Fill.BackgroundColor = XLColor.FromHtml(Green);
            r.Style.Font.FontColor = XLColor.White;
            r.Style.Font.Bold = true;
        }

        // Nume fisier download cu profesor selectat
        private string FileName(string raport, string? profesor, string? spec)
        {
            if (string.IsNullOrWhiteSpace(profesor) || profesor == "Toti")
                return $"{raport}.xlsx";
            var profSafe = string.Concat(profesor.Take(30)
                .Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
            if (!string.IsNullOrWhiteSpace(spec) && spec != "Toti")
            {
                var specSafe = string.Concat(spec.Take(15)
                    .Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                return $"{profSafe}_{specSafe}_{raport}.xlsx";
            }
            return $"{profSafe}_{raport}.xlsx";
        }

        // =====================================================================
        // LISTE DROPDOWNURI
        // =====================================================================
        [HttpGet("liste/ani-universitari")]
        public IActionResult GetAni() => Ok(_cache.GetOrCreate("AniUniv_v4", e =>
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
        public IActionResult GetFacultati() => Ok(_cache.GetOrCreate("Fac_v5", e =>
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
                WHERE p.ID_AnUnivCatedra = @idAn
                  AND p.DenumireCatedra IS NOT NULL
                  AND LTRIM(RTRIM(p.DenumireCatedra)) != ''
                  AND (@idFac = 0 OR p.ID_Facultate = @idFac)
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
                         THEN LEFT(ppm.DenumireSpecializare, CHARINDEX(N'+', ppm.DenumireSpecializare) - 1)
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

        [HttpGet("liste/profesori")]
        public IActionResult GetProfesori([FromQuery] int? idAnUniv,
            [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? specializare)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = "Toti", nume = "Toti" } };
            using var conn = new SqlConnection(_cs); conn.Open();
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
                            THEN LEFT(ppm.DenumireSpecializare, CHARINDEX(N'+', ppm.DenumireSpecializare) - 1)
                            ELSE ppm.DenumireSpecializare END
                       COLLATE Romanian_CI_AS = @spec COLLATE Romanian_CI_AS)
                ORDER BY p.NumeIntreg", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
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
        // =====================================================================
        private string SqlNorma() => $@"
            WITH
            {IdenCte},
            {DedupCte},
            Baza AS (
                SELECT
                    i.IdentificatorUnic, i.NumeIntreg,
                    i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic,
                    CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END AS TipPost,
                    ppm.SpecCurata AS Specializare, ppm.DenumireMaterie,
                    ppm.NrSemestruDinAn AS Semestru, ppm.DenumireFormaInv,
                    ppm.OreConv, ppm.OreCurs,
                    ppm.OreSem + ppm.OreLab + ppm.OreProj AS OreAplic
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE i.TitularAnUniv IS NOT NULL
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                  AND (@spec = N'Toti' OR ppm.SpecCurata COLLATE Romanian_CI_AS = @spec COLLATE Romanian_CI_AS)
                  AND (@tipPost = N'Toti' OR
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                  AND (@sem = 0 OR ppm.NrSemestruDinAn = @sem)
                  AND (@formaInv = N'Toti' OR ppm.DenumireFormaInv COLLATE Romanian_CI_AS
                       = @formaInv COLLATE Romanian_CI_AS)
            ),
            TotProf AS (
                SELECT IdentificatorUnic, SUM(OreConv) AS TotalOre
                FROM Baza GROUP BY IdentificatorUnic
            )
            SELECT
                b.NumeIntreg, b.DenumireFacultate, b.DenumireCatedra,
                b.DenumireGradDidactic, b.TipPost, b.Specializare,
                b.DenumireMaterie, b.Semestru, b.DenumireFormaInv,
                CAST(b.OreCurs  AS DECIMAL(10,2))  AS OreCurs,
                CAST(b.OreAplic AS DECIMAL(10,2))  AS OreAplic,
                CAST(b.OreConv  AS DECIMAL(10,2))  AS OreConv,
                CAST(t.TotalOre AS DECIMAL(10,2))  AS TotalOreProf,
                CAST(CASE WHEN t.TotalOre > 0
                     THEN ROUND(b.OreConv / t.TotalOre * 100, 2) ELSE 0
                END AS DECIMAL(10,2))              AS ProcDinTotal
            FROM Baza b
            INNER JOIN TotProf t ON t.IdentificatorUnic = b.IdentificatorUnic
            ORDER BY b.NumeIntreg, b.TipPost DESC, b.Specializare, b.Semestru";

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
                    Facultate = r["DenumireFacultate"]?.ToString() ?? "",
                    Departament = r["DenumireCatedra"]?.ToString() ?? "",
                    Grad = r["DenumireGradDidactic"]?.ToString() ?? "",
                    TipPost = r["TipPost"]?.ToString() ?? "",
                    Specializare = r["Specializare"]?.ToString() ?? "",
                    Materie = r["DenumireMaterie"]?.ToString() ?? "",
                    Semestru = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    FormaInv = r["DenumireFormaInv"]?.ToString() ?? "",
                    OreCurs = r["OreCurs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreCurs"]),
                    OreAplicatii = r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
                    OreConv = r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]),
                    TotalOreProf = r["TotalOreProf"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreProf"]),
                    ProcentDinTotal = r["ProcDinTotal"] == DBNull.Value ? 0m : Convert.ToDecimal(r["ProcDinTotal"])
                });
            return Ok(result);
        }

        [HttpGet("export/norma")]
        public async Task<IActionResult> ExportNorma(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? specializare,
            [FromQuery] string? tipPost, [FromQuery] int? semestru, [FromQuery] string? formaInvatamant)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Departament"), new DataColumn("Specializare"),
                new DataColumn("Materie"), new DataColumn("Tip Post"),
                new DataColumn("Sem.", typeof(int)), new DataColumn("Forma Inv."),
                new DataColumn("Ore Curs", typeof(decimal)),
                new DataColumn("Ore Aplic.", typeof(decimal)),
                new DataColumn("Ore Conv.", typeof(decimal)),
                new DataColumn("Total Prof.", typeof(decimal)),
                new DataColumn("% Din Total", typeof(decimal))
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
                    r["NumeIntreg"]?.ToString(), r["DenumireCatedra"]?.ToString(),
                    r["Specializare"]?.ToString(), r["DenumireMaterie"]?.ToString(),
                    r["TipPost"]?.ToString(),
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["DenumireFormaInv"]?.ToString(),
                    r["OreCurs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreCurs"]),
                    r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
                    r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]),
                    r["TotalOreProf"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreProf"]),
                    r["ProcDinTotal"] == DBNull.Value ? 0m : Convert.ToDecimal(r["ProcDinTotal"]));
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme");
            ws.Cell(1, 1).Value = $"Detaliere norme | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var c in new[] { "Ore Curs", "Ore Aplic.", "Ore Conv." })
                tbl.Field(c).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Norme", profesor, specializare));
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
                WHERE i.TitularAnUniv IS NOT NULL
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
            [FromQuery] string? profesor, [FromQuery] string? tipPost)
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
            ws.Cell(1, 1).Value = $"Totaluri norme | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var c in new[] { "Ore IF", "Ore ID", "Ore IFR", "Total Ore Conv.", "Total Anual (×14)" })
                tbl.Field(c).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL GENERAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Totaluri_Norme", profesor, null));
        }

        // =====================================================================
        // RAPORT 3: DISTRIBUTIE ORE PE PROGRAME
        // Procent = (OreProgram / TotalOreUniversitateProfesor) * 100
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
                WHERE i.TitularAnUniv IS NOT NULL
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
                o.NumeIntreg, o.DenumireFacultate, o.DenumireCatedra,
                o.Program,
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
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
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
            ws.Cell(1, 1).Value = $"Distributie ore pe programe | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Program").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Distributie_Ore", profesor, null));
        }

        // =====================================================================
        // RAPORT 4: LIMBI STRAINE
        // =====================================================================
        private string SqlLimbi() => $@"
            WITH
            {IdenCte},
            {DedupCte},
            Filtrat AS (
                SELECT
                    i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra,
                    CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END AS TipPost,
                    ppm.OreConv, ppm.NrSemestruDinAn, ppm.DenumireMaterie, ppm.LimbaDePredare
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE i.TitularAnUniv IS NOT NULL
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                  AND (@tipPost = N'Toti' OR
                       CASE ppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                  AND ppm.LimbaDePredare IS NOT NULL
                  AND ppm.LimbaDePredare NOT IN (N'Romana', N'Română', N'Româna',
                                                 N'Romana/Engleza', N'Româna/Engleza')
                  AND ppm.OreConv > 0
            ),
            Dedup2 AS (
                SELECT NumeIntreg, DenumireMaterie, NrSemestruDinAn, TipPost, LimbaDePredare,
                       MAX(OreConv) AS OreD
                FROM Filtrat
                GROUP BY NumeIntreg, DenumireMaterie, NrSemestruDinAn, TipPost, LimbaDePredare
            )
            SELECT
                NumeIntreg,
                CAST(SUM(CASE WHEN NrSemestruDinAn % 2 = 1 THEN OreD ELSE 0 END) AS DECIMAL(10,2)) AS Sem1,
                CAST(SUM(CASE WHEN NrSemestruDinAn % 2 = 0 THEN OreD ELSE 0 END) AS DECIMAL(10,2)) AS Sem2,
                CAST(SUM(OreD) AS DECIMAL(10,2)) AS Total,
                STUFF((
                    SELECT DISTINCT N', ' + d2.LimbaDePredare
                    FROM Dedup2 d2 WHERE d2.NumeIntreg = Dedup2.NumeIntreg
                    FOR XML PATH(N''), TYPE
                ).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N'') AS LimbiPredate
            FROM Dedup2
            GROUP BY NumeIntreg
            HAVING SUM(OreD) > 0
            ORDER BY NumeIntreg";

        [HttpGet("limbi-straine")]
        public async Task<IActionResult> GetLimbi(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor, [FromQuery] string? tipPost)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi(), conn);
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
                    Sem1 = r["Sem1"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Sem1"]),
                    Sem2 = r["Sem2"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Sem2"]),
                    Total = r["Total"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Total"]),
                    LimbiPredate = r["LimbiPredate"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/limbi-straine")]
        public async Task<IActionResult> ExportLimbi(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor, [FromQuery] string? tipPost)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Sem. 1 (h)", typeof(decimal)),
                new DataColumn("Sem. 2 (h)", typeof(decimal)),
                new DataColumn("Total (h)", typeof(decimal)),
                new DataColumn("Limbi Predate")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi(), conn);
            cmd.CommandTimeout = 120;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(),
                    r["Sem1"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Sem1"]),
                    r["Sem2"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Sem2"]),
                    r["Total"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Total"]),
                    r["LimbiPredate"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Limbi Straine");
            ws.Cell(1, 1).Value = $"Ore predate in limbi straine | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Total (h)").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Limbi_Straine", profesor, null));
        }

        // =====================================================================
        // RAPORT 5: DISCIPLINE PER PROFESOR
        // =====================================================================
        private string SqlDisc(string? formaInv = null) => $@"
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
                WHERE i.TitularAnUniv IS NOT NULL
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE Romanian_CI_AS = @prof COLLATE Romanian_CI_AS)
                  AND ppm.DenumireMaterie IS NOT NULL
                  AND LTRIM(RTRIM(ppm.DenumireMaterie)) != N''
                  {(string.IsNullOrWhiteSpace(formaInv)
                      ? ""
                      : "AND ppm.DenumireFormaInv COLLATE Romanian_CI_AS = @formaInvFilter COLLATE Romanian_CI_AS")}
            )
            SELECT
                dm.NumeIntreg,
                MAX(dm.DenumireFacultate)    AS Facultate,
                MAX(dm.DenumireCatedra)      AS Departament,
                MAX(dm.DenumireGradDidactic) AS Grad,
                MAX(dm.DenumireFormaInv)     AS FormaInv,
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

        // Export ZIP cu 3 fisiere: IF, ID, IFR
        [HttpGet("export/discipline-zip")]
        public async Task<IActionResult> ExportDisciplineZip(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate,
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            // Valorile exacte din DB pentru DenumireFormaInv
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
                    ws.Cell(1, 1).Value = $"Discipline predate - {formaLabel} | An: {idAn}";
                    ws.Cell(1, 1).Style.Font.Bold = true;
                    ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
                    ws.Range(1, 1, 1, dt.Columns.Count).Merge();
                    var tbl = ws.Cell(3, 1).InsertTable(dt);
                    tbl.Theme = XLTableTheme.None;
                    StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
                    ws.Columns(1, 5).AdjustToContents();
                    ws.Column(6).Width = 80; ws.Column(6).Style.Alignment.WrapText = true;
                    using var wbs = new MemoryStream(); wb.SaveAs(wbs); wbs.Position = 0; wbs.CopyTo(es);
                }
            }
            memZip.Position = 0;
            return File(memZip.ToArray(), "application/zip", "Discipline_Predate.zip");
        }

        // =====================================================================
        // RAPORT 6: TITULARI
        // =====================================================================
        private string SqlTitulari() => $@"
            WITH {IdenCte}
            SELECT i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic
            FROM Identitate i
            WHERE i.TitularAnUniv = 1 {FiltruIdent}
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
            [FromQuery] int? idCatedra, [FromQuery] string? profesor)
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
            ws.Cell(1, 1).Value = $"Cadre didactice titulare | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowFunction = XLTotalsRowFunction.Count;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName("Titulari", profesor, null));
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
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
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
            ws.Cell(1, 1).Value = $"Asociati / Colaboratori | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowFunction = XLTotalsRowFunction.Count;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Colaboratori.xlsx");
        }

        // =====================================================================
        // RAPORT 8: ANS — fractiuni norma per domeniu stiinta
        // Regula 741: TitularAnUniv=1 + ID_Facultate!=41 + TipGrad IN(1,2,3,4,10,11)
        //             + CNP NOT NULL + SUM(vcm.NrOreConventionale) > 0
        //
        // JOIN ANS: vcm.DenumireSpecializare + vcm.DenumireFacultate → View_FDS
        //           Ambele cu COLLATE Romanian_CI_AS pentru cross-DB collation
        //           FDS are denumirile UPPERCASE - CI_AS rezolva case mismatch
        // =====================================================================
        [HttpGet("raport-ans")]
        public async Task<IActionResult> GetAns(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();

            // Pas 1: Domenii ANS dinamice din DB
            var domenii = new List<(int Id, string Den)>();
            using var conn = new SqlConnection(_cs);
            await conn.OpenAsync();
            using (var cmdD = new SqlCommand(
                "SELECT ID_Element, Denumire FROM [AGSIS].[dbo].[N_RAMURA_STIINTA_ANS] ORDER BY ID_Element", conn))
            {
                using var rd = await cmdD.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                    domenii.Add((Convert.ToInt32(rd["ID_Element"]), rd["Denumire"]?.ToString() ?? ""));
            }

            // Pas 2: Titulari oficiali (Regula 741)
            var titulari = new List<TitAns>();
            string sqlTit = $@"
                WITH {IdenCte}
                SELECT i.ID_Profesor, i.IdentificatorUnic, i.NumeIntreg,
                       i.DenumireFacultate, i.DenumireCatedra,
                       i.DenumireGradDidactic, i.ID_TipGrad, i.NormaDB
                FROM Identitate i
                {FiltruAns741}
                ORDER BY i.NumeIntreg";
            using (var cmdT = new SqlCommand(sqlTit, conn))
            {
                cmdT.CommandTimeout = 60;
                AddCore(cmdT, idAn, idFacultate ?? 0, idCatedra ?? 0);
                using var rt = await cmdT.ExecuteReaderAsync();
                while (await rt.ReadAsync())
                {
                    int? idTip = rt["ID_TipGrad"] == DBNull.Value ? null : (int?)Convert.ToInt32(rt["ID_TipGrad"]);
                    decimal nDb = rt["NormaDB"] == DBNull.Value ? 0m : Convert.ToDecimal(rt["NormaDB"]);
                    titulari.Add(new TitAns
                    {
                        IdProf = Convert.ToInt32(rt["ID_Profesor"]),
                        CNP = rt["IdentificatorUnic"]?.ToString() ?? "",
                        Nume = rt["NumeIntreg"]?.ToString() ?? "",
                        Fac = rt["DenumireFacultate"]?.ToString() ?? "",
                        Dept = rt["DenumireCatedra"]?.ToString() ?? "",
                        GradD = rt["DenumireGradDidactic"]?.ToString() ?? "",
                        IdTip = idTip,
                        Norma = nDb > 0 ? nDb : NormaFallback(idTip)
                    });
                }
            }

            // Pas 3: Ore per profesor per domeniu ANS
            // JOIN VIEW_CentralizareMateriiProfesor → View_FDS pe (DenumireSpecializare + DenumireFacultate)
            // Ambele COLLATE Romanian_CI_AS (FDS are UPPERCASE, vcm are mixed case)
            var oreAns = new Dictionary<int, Dictionary<int, decimal>>();
            const string sqlOre = @"
                SELECT
                    vcm.ID_Profesor,
                    fds.ID_N_Domeniu_Studiu_ANS AS ID_ANS,
                    SUM(vcm.NrOreConventionale)  AS Ore
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[View_FDS] fds
                    ON vcm.DenumireSpecializare COLLATE Romanian_CI_AS
                       = fds.DenumireSpecializare COLLATE Romanian_CI_AS
                    AND vcm.DenumireFacultate COLLATE Romanian_CI_AS
                       = fds.DenumireFacultate COLLATE Romanian_CI_AS
                    AND fds.ID_AnUniv = @idAn
                    AND fds.id_metaspecializare > 0
                    AND fds.ID_N_Domeniu_Studiu_ANS IS NOT NULL
                WHERE vcm.ID_AnUniv = @idAn
                GROUP BY vcm.ID_Profesor, fds.ID_N_Domeniu_Studiu_ANS
                HAVING SUM(vcm.NrOreConventionale) > 0";
            using (var cmdO = new SqlCommand(sqlOre, conn))
            {
                cmdO.CommandTimeout = 120;
                cmdO.Parameters.AddWithValue("@idAn", idAn);
                using var ro = await cmdO.ExecuteReaderAsync();
                while (await ro.ReadAsync())
                {
                    int idP = Convert.ToInt32(ro["ID_Profesor"]);
                    int idA = Convert.ToInt32(ro["ID_ANS"]);
                    decimal ore = Convert.ToDecimal(ro["Ore"]);
                    if (!oreAns.ContainsKey(idP)) oreAns[idP] = new();
                    oreAns[idP].TryAdd(idA, 0m);
                    oreAns[idP][idA] += ore;
                }
            }

            // Pas 4: Calcul fractiuni
            var profesori = new List<object>();
            int nrCrt = 1;
            foreach (var t in titulari)
            {
                var frac = new Dictionary<string, decimal>();
                if (oreAns.TryGetValue(t.IdProf, out var oreP))
                {
                    decimal totalOre = oreP.Values.Sum();
                    decimal baza = Math.Max(totalOre, t.Norma);
                    if (totalOre > 0 && baza > 0)
                        foreach (var kv in oreP)
                        {
                            var dom = domenii.FirstOrDefault(d => d.Id == kv.Key);
                            if (dom.Den == null) continue;
                            decimal f = Math.Round(kv.Value / baza, 2);
                            if (f > 0) frac[dom.Den] = f;
                        }
                }
                profesori.Add(new
                {
                    NrCrt = nrCrt++,
                    NumeComplet = t.Nume,
                    CNP = t.CNP,
                    Facultate = t.Fac,
                    Departament = t.Dept,
                    GradFunctie = GradANS(t.IdTip, t.GradD),
                    DomeniiMapate = frac
                });
            }
            return Ok(new { Domenii = domenii.Select(d => d.Den).ToList(), Profesori = profesori });
        }

        [HttpGet("export/raport-ans")]
        public IActionResult ExportAns(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var domenii = new List<(int Id, string Den)>();
            using var conn = new SqlConnection(_cs); conn.Open();
            using (var cmdD = new SqlCommand(
                "SELECT ID_Element, Denumire FROM [AGSIS].[dbo].[N_RAMURA_STIINTA_ANS] ORDER BY ID_Element", conn))
            { using var rd = cmdD.ExecuteReader(); while (rd.Read()) domenii.Add((Convert.ToInt32(rd[0]), rd[1]?.ToString() ?? "")); }

            var titulari = new List<TitAns>();
            string sqlTit = $@"WITH {IdenCte} SELECT i.ID_Profesor, i.IdentificatorUnic, i.NumeIntreg,
                   i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic, i.ID_TipGrad, i.NormaDB
                FROM Identitate i {FiltruAns741} ORDER BY i.NumeIntreg";
            using (var cmdT = new SqlCommand(sqlTit, conn))
            {
                cmdT.CommandTimeout = 60;
                AddCore(cmdT, idAn, idFacultate ?? 0, idCatedra ?? 0);
                using var rt = cmdT.ExecuteReader();
                while (rt.Read())
                {
                    int? idTip = rt["ID_TipGrad"] == DBNull.Value ? null : (int?)Convert.ToInt32(rt["ID_TipGrad"]);
                    decimal nDb = rt["NormaDB"] == DBNull.Value ? 0m : Convert.ToDecimal(rt["NormaDB"]);
                    titulari.Add(new TitAns
                    {
                        IdProf = Convert.ToInt32(rt["ID_Profesor"]),
                        CNP = rt["IdentificatorUnic"]?.ToString() ?? "",
                        Nume = rt["NumeIntreg"]?.ToString() ?? "",
                        Fac = rt["DenumireFacultate"]?.ToString() ?? "",
                        Dept = rt["DenumireCatedra"]?.ToString() ?? "",
                        GradD = rt["DenumireGradDidactic"]?.ToString() ?? "",
                        IdTip = idTip,
                        Norma = nDb > 0 ? nDb : NormaFallback(idTip)
                    });
                }
            }
            var oreAns = new Dictionary<int, Dictionary<int, decimal>>();
            const string sqlOre = @"SELECT vcm.ID_Profesor, fds.ID_N_Domeniu_Studiu_ANS AS ID_ANS, SUM(vcm.NrOreConventionale) AS Ore
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[View_FDS] fds
                    ON vcm.DenumireSpecializare COLLATE Romanian_CI_AS = fds.DenumireSpecializare COLLATE Romanian_CI_AS
                    AND vcm.DenumireFacultate COLLATE Romanian_CI_AS = fds.DenumireFacultate COLLATE Romanian_CI_AS
                    AND fds.ID_AnUniv = @idAn AND fds.id_metaspecializare > 0
                    AND fds.ID_N_Domeniu_Studiu_ANS IS NOT NULL
                WHERE vcm.ID_AnUniv = @idAn
                GROUP BY vcm.ID_Profesor, fds.ID_N_Domeniu_Studiu_ANS
                HAVING SUM(vcm.NrOreConventionale) > 0";
            using (var cmdO = new SqlCommand(sqlOre, conn))
            {
                cmdO.CommandTimeout = 120;
                cmdO.Parameters.AddWithValue("@idAn", idAn);
                using var ro = cmdO.ExecuteReader();
                while (ro.Read())
                {
                    int idP = Convert.ToInt32(ro["ID_Profesor"]);
                    int idA = Convert.ToInt32(ro["ID_ANS"]);
                    decimal ore = Convert.ToDecimal(ro["Ore"]);
                    if (!oreAns.ContainsKey(idP)) oreAns[idP] = new();
                    oreAns[idP].TryAdd(idA, 0m);
                    oreAns[idP][idA] += ore;
                }
            }

            int nrD = domenii.Count, colTot = 9 + nrD + 1;
            var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("CD DRU");
            var gC = XLColor.FromHtml(Green);
            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, colTot).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov"; ws.Range(3, 1, 3, 6).Merge();
            string[] antet = { "Nr.\nCrt.", "Nume si prenume", "CNP", "Functie", "Forma\nangajare", "Cond.\ndoctorat", "Varsta", "Facultate", "Departament" };
            for (int c = 1; c <= 9; c++) { ws.Cell(5, c).Value = antet[c - 1]; ws.Range(5, c, 7, c).Merge(); }
            ws.Cell(5, colTot).Value = "Total"; ws.Range(5, colTot, 7, colTot).Merge();
            for (int i = 0; i < nrD; i++) { ws.Cell(6, 10 + i).Value = domenii[i].Den; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }
            for (int c = 1; c <= 9; c++) ws.Cell(8, c).Value = ((char)('A' + c - 1)).ToString();
            for (int i = 0; i < nrD; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, colTot).Value = nrD;
            for (int row = 5; row <= 8; row++) for (int col = 1; col <= colTot; col++)
            {
                ws.Cell(row, col).Style.Font.Bold = true; ws.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, col).Style.Alignment.WrapText = true; ws.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
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
                ws.Cell(row2, 1).Value = nrCrt++; ws.Cell(row2, 2).Value = t.Nume;
                ws.Cell(row2, 3).Value = ""; ws.Cell(row2, 4).Value = GradANS(t.IdTip, t.GradD);
                ws.Cell(row2, 5).Value = ""; ws.Cell(row2, 6).Value = "";
                ws.Cell(row2, 7).Value = ""; ws.Cell(row2, 8).Value = t.Fac; ws.Cell(row2, 9).Value = t.Dept;
                decimal totFrac = 0m;
                if (oreAns.TryGetValue(t.IdProf, out var oreP))
                {
                    decimal tOre = oreP.Values.Sum();
                    decimal baza = Math.Max(tOre, t.Norma);
                    if (tOre > 0 && baza > 0)
                        for (int i = 0; i < nrD; i++)
                            if (oreP.TryGetValue(domenii[i].Id, out decimal oreD))
                            {
                                decimal f = Math.Round(oreD / baza, 2);
                                if (f > 0) { ws.Cell(row2, 10 + i).Value = (double)f; ws.Cell(row2, 10 + i).Style.NumberFormat.Format = "0.00"; totFrac += f; }
                            }
                }
                ws.Cell(row2, colTot).Value = (double)Math.Round(totFrac, 2);
                ws.Cell(row2, colTot).Style.NumberFormat.Format = "0.00";
                for (int c = 1; c <= colTot; c++) ws.Cell(row2, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                row2++;
            }
            using var stream = new MemoryStream(); wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Raport_ANS.xlsx");
        }

        private class TitAns
        {
            public int IdProf { get; set; }
            public string CNP { get; set; } = "";
            public string Nume { get; set; } = "";
            public string Fac { get; set; } = "";
            public string Dept { get; set; } = "";
            public string GradD { get; set; } = "";
            public int? IdTip { get; set; }
            public decimal Norma { get; set; }
        }
    }
}