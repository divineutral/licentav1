using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using System.IO.Compression;
using ClosedXML.Excel;

namespace LicentaV1.Controllers
{
    // Filtru global: orice excepție din controller devine JSON cu detalii.
    // Fără asta, ASP.NET în Development întoarce o pagină HTML cu stack trace,
    // iar fetch().json() pe frontend pică pe primul caracter ('M' din "Microsoft...").
    public class JsonExceptionFilter : IExceptionFilter
    {
        public void OnException(ExceptionContext ctx)
        {
            var msg = ctx.Exception.Message;
            if (ctx.Exception is SqlException sqlEx)
                msg = $"SQL Error {sqlEx.Number}: {sqlEx.Message}";
            ctx.Result = new ObjectResult(new
            {
                error = true,
                message = msg,
                type = ctx.Exception.GetType().Name,
                endpoint = ctx.HttpContext.Request.Path.ToString()
            })
            { StatusCode = 500 };
            ctx.ExceptionHandled = true;
        }
    }

    [Route("api/[controller]")]
    [ApiController]
    [TypeFilter(typeof(JsonExceptionFilter))]
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

        private int GetAnCurent() => _cache.GetOrCreate("AnCurent_v7", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(4);
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT TOP 1 au.ID_AnUniv
                FROM [AGSIS].[dbo].[AnUniversitar] au
                WHERE EXISTS (SELECT 1 FROM [AGSIS].[pi].[View_PostProfesorMaterie] ppm WHERE ppm.ID_AnUniv = au.ID_AnUniv)
                ORDER BY au.Ordine DESC", conn);
            var r = cmd.ExecuteScalar();
            return r != null ? Convert.ToInt32(r) : 45;
        });

        // =====================================================================
        // OreUnice CTE — deduplicare corecta prin ROW_NUMBER
        // PARTITION BY (ID_Profesor, ID_PlanMaterie_Prestator, NrSemestruDinAn)
        // => un singur rand per activitate, indiferent de cate grupe/specializari are cuplajul
        //
        // FIX CUPLAJE: PARTITION BY pe (ID_Profesor, ID_PlanMaterie_Prestator, NrSemestruDinAn)
        //              pastreaza un singur rand reprezentativ per activitate fizica.
        //              Coloana Mentiuni va indica specializarile cuplate.
        //
        // FIX PREGATIRE PEDAGOGICA: LEFT JOIN pe View_FDS + excludere ciclu
        //              'Pregatire Pedagogica%' previne dublarea orelor
        //              prin inregistrarile administrative de pedagogie.
        // =====================================================================
        // FIX: sursa este [AGSIS].[pi].[View_PostProfesorMaterie] — conform query-urilor validate
        // Nu mai filtram pedagogia in CTE — fiecare raport aplica filtrele proprii
        private const string OreUniceCte = @"
OreUnice AS (
    SELECT
        vppm.ID_Profesor,
        vppm.ID_Facultate,
        vppm.ID_Catedra,
        vppm.Denumire                                           AS DenumireMaterie,
        vppm.DenumireSpecializare,
        vppm.DenumireScurtaSpecializare,
        CASE WHEN CHARINDEX(N'+', vppm.DenumireSpecializare) > 0
             THEN LEFT(vppm.DenumireSpecializare, CHARINDEX(N'+', vppm.DenumireSpecializare) - 1)
             ELSE vppm.DenumireSpecializare END                 AS SpecCurata,
        vppm.NrSemestruDinAn,
        vppm.TitularSauSuplinitor,
        
        -- CORECTIE NUME COLOANE PENTRU VIEW_POSTPROFESORMATERIE
        vppm.LimbaDePredare,
        -- In View_PostProfesorMaterie, forma de invatamant este legata de ID_TipFormaInv
        -- Pentru denumire, folosim un CASE rapid sau o aducem din nomenclator
        CASE vppm.ID_TipFormaInv 
            WHEN 1 THEN N'Cu frecvență' 
            WHEN 2 THEN N'Frecvență redusă' 
            WHEN 3 THEN N'Învățământ la distanță' 
            ELSE N'IF' END                                      AS DenumireFormaInv,
            
        -- In acest View, denumirile sunt de obicei fara prefixul 'Denumire'
        -- Daca tot da eroare, le setam ca N'Nespecificat' si le aducem din Identitate
        vppm.DenumireSpecializare                               AS FacPPM, 
        vppm.DenumireSpecializare                               AS DeptOra,
        
        vppm.ApartineDeCuplaj,
        vppm.ID_PlanMaterie_Prestator,
        vppm.NrOreConventionale                                 AS OreConv,
        vppm.Nr_Ore_Curs                                        AS OreCurs,
        vppm.Nr_Ore_Seminar                                     AS OreSem,
        vppm.Nr_Ore_Laborator                                   AS OreLab,
        vppm.Nr_Ore_Proiect                                     AS OreProj,
        ROW_NUMBER() OVER (
            PARTITION BY vppm.ID_Profesor, vppm.ID_PlanMaterie_Prestator, vppm.NrSemestruDinAn
            ORDER BY vppm.DenumireSpecializare
        )                                                       AS Rn
    FROM [AGSIS].[pi].[View_PostProfesorMaterie] vppm
    WHERE vppm.ID_AnUniv = @idAn
)";
        private const string IdenCte = @"
        Identitate AS (
            SELECT
                p.ID_Profesor,
                ISNULL(NULLIF(LTRIM(RTRIM(ISNULL(p.CNP,''))), ''),
                       CAST(p.ID_Profesor AS VARCHAR(20)))          AS IdentificatorUnic,
                p.NumeIntreg,
                p.ID_Facultate,
                p.DenumireFacultate,
                p.ID_Catedra,
                p.DenumireCatedra,
                p.TitularAnUniv,
                p.ID_TipGradDidacticAnUniv                          AS ID_TipGrad,
                p.DenumireGradDidactic,
                ISNULL(n.NrOreConventionaleTitular, 0)              AS NormaDB
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            LEFT JOIN [AGSIS].[pi].[NormaOreConventionale] n
                ON n.ID_TipGradDidactic = p.ID_TipGradDidacticAnUniv AND n.ID_AnUniv = @idAn
            WHERE p.ID_AnUnivCatedra = @idAn
        )";

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
        // DROPDOWN-URI
        // =====================================================================
        [HttpGet("liste/ani-universitari")]
        public IActionResult GetAni() => Ok(_cache.GetOrCreate("AniUniv_v7", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
            var lst = new List<object>();
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT au.ID_AnUniv, LTRIM(RTRIM(au.Denumire)) AS Denumire
                FROM [AGSIS].[dbo].[AnUniversitar] au
                WHERE EXISTS (SELECT 1 FROM [AGSIS].[pi].[View_PostProfesorMaterie] ppm WHERE ppm.ID_AnUniv = au.ID_AnUniv)
                ORDER BY au.Ordine DESC", conn);
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
                    WHERE EXISTS (SELECT 1 FROM [AGSIS].[pi].[View_PostProfesorMaterie] ppm2 WHERE ppm2.ID_AnUniv = au.ID_AnUniv)
                    ORDER BY au.Ordine DESC)
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
                FROM [AGSIS].[pi].[View_PostProfesorMaterie] ppm
                WHERE ppm.ID_AnUniv = @idAn AND ppm.DenumireSpecializare IS NOT NULL
                  AND (@idFac = 0 OR ppm.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR ppm.ID_Catedra = @idCatedra)
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
                INNER JOIN [AGSIS].[pi].[View_PostProfesorMaterie] ppm
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
        // FIX CUPLAJE: OreUnice cu PARTITION BY (ID_Profesor, ID_PlanMaterie_Prestator, NrSemestruDinAn)
        //              => un singur rand per activitate fizica, orele nu se mai sumeaza per specializare cuplatat
        // FIX PREGATIRE PEDAGOGICA: exclus din OreUnice prin LEFT JOIN pe View_FDS + filtru ciclu
        // =====================================================================
        private string SqlNorma() => $@"
            WITH {IdenCte}, {OreUniceCte},
            FdsCiclu AS (
                SELECT DISTINCT DenumireSpecializare, DenumireCicluInv, ID_AnUniv
                FROM [AGSIS].[dbo].[View_FDS] WHERE ID_AnUniv = @idAn AND id_metaspecializare > 0
            ),
            Baza AS (
                SELECT i.IdentificatorUnic, i.NumeIntreg,
                       i.DenumireFacultate AS FacAngajator, ou.DeptOra AS DepartamentOra,
                       i.DenumireGradDidactic,
                       CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END AS TipPost,
                       ou.SpecCurata AS Specializare, ou.DenumireMaterie,
                       ou.NrSemestruDinAn AS Semestru, ou.DenumireFormaInv,
                       ou.OreConv, ou.OreCurs, ou.OreSem + ou.OreLab + ou.OreProj AS OreAplic,
                       ISNULL(fds.DenumireCicluInv, N'Nespecificat') AS CicluInv,
                       CASE WHEN ou.ApartineDeCuplaj IS NOT NULL
                            THEN N'Cuplat cu: ' + ISNULL(STUFF((
                                SELECT DISTINCT N', ' + vdc.DenumireScurtaSpecializare
                                FROM [AGSIS].[pi].[View_DetaliereCuplaje] vdc
                                WHERE vdc.ID_AnUniv = @idAn
                                  AND vdc.ID_Cuplaj = ou.ApartineDeCuplaj
                                  AND vdc.DenumireScurtaSpecializare COLLATE DATABASE_DEFAULT
                                      != ou.DenumireScurtaSpecializare COLLATE DATABASE_DEFAULT
                                FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N''),
                                N'incarcare comuna')
                            ELSE N'' END AS Mentiuni
                FROM OreUnice ou
                INNER JOIN Identitate i ON i.ID_Profesor = ou.ID_Profesor
                LEFT JOIN FdsCiclu fds ON ou.SpecCurata COLLATE DATABASE_DEFAULT
                                       = fds.DenumireSpecializare COLLATE DATABASE_DEFAULT
                WHERE ou.Rn = 1
                  {FiltruIdent}
                  AND (@prof     = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@spec     = N'Toti' OR ou.SpecCurata COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
                  AND (@tipPost  = N'Toti' OR CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
                  AND (@sem      = 0 OR ou.NrSemestruDinAn = @sem)
                  AND (@formaInv = N'Toti' OR ou.DenumireFormaInv COLLATE DATABASE_DEFAULT = @formaInv COLLATE DATABASE_DEFAULT)
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
                    OreAplic = r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
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
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad"),
                new DataColumn("Tip Post"), new DataColumn("Specializare"), new DataColumn("Ciclu"),
                new DataColumn("Materie"), new DataColumn("Semestru", typeof(int)),
                new DataColumn("Forma Inv."), new DataColumn("Ore Curs", typeof(decimal)),
                new DataColumn("Ore Aplic.", typeof(decimal)), new DataColumn("Ore Conv.", typeof(decimal)),
                new DataColumn("Mentiuni")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma(), conn); cmd.CommandTimeout = 180;
            AddNormaParams(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0,
                profesor ?? "", specializare ?? "", tipPost ?? "Toti",
                semestru ?? 0, formaInvatamant ?? "Toti", cicluStudii ?? "Toti");
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["FacAngajator"]?.ToString(),
                    r["DepartamentOra"]?.ToString(), r["DenumireGradDidactic"]?.ToString(),
                    r["TipPost"]?.ToString(), r["Specializare"]?.ToString(), r["CicluInv"]?.ToString(),
                    r["DenumireMaterie"]?.ToString(),
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["DenumireFormaInv"]?.ToString(),
                    r["OreCurs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreCurs"]),
                    r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
                    r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]),
                    r["Mentiuni"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Norma Profesori");
            ws.Cell(1, 1).Value = "Detaliere Norme Profesori"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, tipPost, formaInvatamant, specializare, semestru ?? 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Aplic.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Norma_Profesori", profesor));
        }

        // =====================================================================
        // RAPORT 2: NORMA TOTALURI
        // FIX: OreUnice cu Rn=1 — cuplajele nu mai dubleaza orele
        // =====================================================================
        private string SqlTotaluri() => $@"
            WITH {IdenCte}, {OreUniceCte},
            Baza AS (
                SELECT i.IdentificatorUnic, i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra,
                       CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END AS TipPost,
                       ou.DenumireFormaInv,
                       ou.OreConv
                FROM OreUnice ou INNER JOIN Identitate i ON i.ID_Profesor = ou.ID_Profesor
                WHERE ou.Rn = 1
                  {FiltruIdent}
                  AND (@prof    = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@tipPost = N'Toti' OR CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
            )
            SELECT
                NumeIntreg, MAX(DenumireFacultate) AS Facultate, MAX(DenumireCatedra) AS Departament, TipPost,
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
        public async Task<IActionResult> ExportNormaTotaluri(
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
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Totaluri_Norme.xlsx");
        }

        // =====================================================================
        // RAPORT 3: DISTRIBUTIE ORE PE PROGRAME
        // =====================================================================
        private string SqlDistrib() => $@"
            WITH {IdenCte}, {OreUniceCte},
            OrePerProg AS (
                SELECT i.IdentificatorUnic, i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra,
                       ou.SpecCurata AS Program, SUM(ou.OreConv) AS OreProgram
                FROM OreUnice ou INNER JOIN Identitate i ON i.ID_Profesor = ou.ID_Profesor
                WHERE ou.Rn = 1
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@spec = N'Toti' OR ou.SpecCurata COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
                GROUP BY i.IdentificatorUnic, i.NumeIntreg, i.DenumireFacultate, i.DenumireCatedra, ou.SpecCurata
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
            tbl.Field("Ore Program").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Distributie_Ore", profesor));
        }

        // =====================================================================
        // RAPORT 4: LIMBI STRAINE
        // FIX 1: Filtru ID_Facultate != 17 (Litere)
        // FIX 2: OreUnice exclude deja Pregatire Pedagogica
        // =====================================================================
        private string SqlLimbi() => $@"
            -- FIX: JOIN pe DenumireSpecializare (nenescurtata), nu pe SpecCurata
            -- Conform query-ului master validat: vppm.DenumireSpecializare = ps.NumeSistem
            -- SpecCurata taie dupa '+' si pierde programele cu paranteze din View_FDS
            WITH {IdenCte}, {OreUniceCte},
            ProgLimbaStrain AS (
                SELECT DISTINCT
                    LTRIM(RTRIM(f.DenumireSpecializare)) COLLATE DATABASE_DEFAULT AS NumeSistem,
                    LTRIM(RTRIM(f.LimbaPredare))         AS LimbaPredare,
                    f.DenumireCicluInv,
                    f.CicluDeStudii                      AS Ciclu
                FROM [AGSIS].[dbo].[View_FDS] f
                WHERE f.ID_AnUniv = @idAn
                  AND f.id_metaspecializare > 0
                  AND f.LimbaPredare IS NOT NULL
                  AND LTRIM(RTRIM(f.LimbaPredare)) NOT IN (N'', N'RO', N'Ro', N'ro', N'Romana', N'Română')
            ),
            Filtrat AS (
                SELECT i.NumeIntreg, ou.NrSemestruDinAn, ou.OreConv,
                       pls.LimbaPredare AS LimbaProgram,
                       ou.SpecCurata    AS ProgramStudiu,
                       ISNULL(pls.DenumireCicluInv, N'Neidentificat') AS CicluStudii
                FROM OreUnice ou
                INNER JOIN Identitate i ON i.ID_Profesor = ou.ID_Profesor
                -- FIX: JOIN pe DenumireSpecializare (nenescurtata) = NumeSistem
                INNER JOIN ProgLimbaStrain pls
                    ON LTRIM(RTRIM(ou.DenumireSpecializare)) COLLATE DATABASE_DEFAULT
                       = pls.NumeSistem
                WHERE ou.Rn = 1
                  {FiltruIdent}
                  AND (@prof    = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@tipPost = N'Toti' OR CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
                  AND (@ciclu   = N'Toti' OR ISNULL(pls.DenumireCicluInv, N'Neidentificat') COLLATE DATABASE_DEFAULT = @ciclu COLLATE DATABASE_DEFAULT)
                  AND ou.OreConv > 0
                  -- Filtru ore fizice: eliminam randurile fara activitate reala
                  AND (ISNULL(ou.OreCurs,0) + ISNULL(ou.OreSem,0) + ISNULL(ou.OreLab,0)) > 0
            )
            SELECT NumeIntreg, NrSemestruDinAn AS Semestru,
                   CAST(SUM(OreConv) AS DECIMAL(10,2)) AS TotalOre,
                   MAX(LimbaProgram) AS LimbaProgram,
                   ProgramStudiu, CicluStudii
            FROM Filtrat
            GROUP BY NumeIntreg, NrSemestruDinAn, ProgramStudiu, CicluStudii
            HAVING SUM(OreConv) > 0
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
            tbl.Field("Total Ore (h)").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Limbi_Straine", profesor));
        }

        // =====================================================================
        // RAPORT 5: DISCIPLINE PER PROFESOR
        // =====================================================================
        private string SqlDisc() => $@"
            WITH {IdenCte}, {OreUniceCte},
            Distinct_Mat AS (
                SELECT DISTINCT i.IdentificatorUnic, i.NumeIntreg,
                    i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic,
                    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ou.DenumireMaterie, N'  ', N'><'), N'<>', N''), N'><', N' '))) AS DenumireMaterie
                FROM OreUnice ou INNER JOIN Identitate i ON i.ID_Profesor = ou.ID_Profesor
                WHERE ou.Rn = 1
                  {FiltruIdent}
                  AND (@prof = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@sem  = 0       OR ou.NrSemestruDinAn = @sem)
                  AND ou.DenumireMaterie IS NOT NULL AND LTRIM(RTRIM(ou.DenumireMaterie)) != N''
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
        public async Task<IActionResult> ExportDiscZip(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] int? semestru,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad"),
                new DataColumn("Discipline"), new DataColumn("Nr. Discipline", typeof(int))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDisc(), conn); cmd.CommandTimeout = 180;
            AddCore(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(), r["Grad"]?.ToString(),
                    r["Discipline"]?.ToString(),
                    r["NrDisc"] == DBNull.Value ? 0 : Convert.ToInt32(r["NrDisc"]));
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Discipline");
            ws.Cell(1, 1).Value = "Raport Discipline per Profesor"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green); ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            AddExcelFilters(ws, dt.Columns.Count, idAn, numeFacultate, numeDepartament, profesor, null, null, null, semestru ?? 0);
            var tbl = ws.Cell(4, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr. Discipline").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(4, 1, 4, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName("Discipline", profesor));
        }

        [HttpGet("export/discipline")]
        public async Task<IActionResult> ExportDisc(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] int? semestru,
            [FromQuery] string? numeFacultate, [FromQuery] string? numeDepartament)
            => await ExportDiscZip(idAnUniv, idFacultate, idCatedra, profesor, semestru, numeFacultate, numeDepartament);

        // =====================================================================
        // RAPORT 6: TITULARI
        // =====================================================================
        // Query master fidel: DISTINCT pe (Nume, Prenume, Marca, Grad, Catedra, Facultate)
        // produce 741 rezultate (in loc de 657) deoarece pastreaza afilierile multiple per profesor.
        // Filtrul pe Facultate/Catedra este OPTIONAL (cand 0 -> nu filtreaza, identic cu master).
        private const string SqlTitulariMaster = @"
            SELECT DISTINCT
                V.ID_Profesor,
                V.Nume,
                V.Prenume,
                V.Marca,
                V.NumeIntreg,
                V.DenumireGradDidactic,
                V.DenumireCatedra   AS Departament,
                V.DenumireFacultate AS Facultate
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] AS V
            WHERE V.ID_AnUnivCatedra = @idAn
              AND V.TitularAnUniv = 1
              AND (@idFac     = 0 OR V.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR V.ID_Catedra   = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [AGSIS].[pi].[Post] AS P
                  INNER JOIN [AGSIS].[pi].[Post_Profesor] AS PP ON P.ID_Post = PP.ID_Post
                  INNER JOIN [AGSIS].[pi].[StatDeFunctii] AS SF ON P.ID_StatDeFunctii = SF.ID_StatDeFunctii
                  WHERE PP.ID_Profesor         = V.ID_Profesor
                    AND SF.ID_AnUniv           = @idAn
                    AND P.TitularSauSuplinitor = 1
                    AND P.Deleted              = 0
                    AND PP.Deleted             = 0
              )
            ORDER BY V.Nume COLLATE Romanian_CI_AS, V.Prenume COLLATE Romanian_CI_AS";

        [HttpGet("titulari")]
        public async Task<IActionResult> GetTitulari(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulariMaster, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Nume = r["Nume"]?.ToString() ?? "",
                    Prenume = r["Prenume"]?.ToString() ?? "",
                    Marca = r["Marca"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    Grad = r["DenumireGradDidactic"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/titulari")]
        public async Task<IActionResult> ExportTitulari(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)),
                new DataColumn("Nume"), new DataColumn("Prenume"), new DataColumn("Marca"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad Didactic")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulariMaster, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["Nume"]?.ToString(),
                    r["Prenume"]?.ToString(),
                    r["Marca"]?.ToString(),
                    r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(),
                    r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Titulari");
            ws.Cell(1, 1).Value = $"Cadre Didactice Titulare | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Titulari.xlsx");
        }

        // =====================================================================
        // RAPORT 7: COLABORATORI
        // Query master fidel: DISTINCT pe (Nume, Prenume, Marca, Grad, Catedra, Facultate)
        // pastreaza afilierile multiple per profesor.
        // =====================================================================
        private const string SqlColaboratoriMaster = @"
            SELECT DISTINCT
                V.ID_Profesor,
                V.Nume,
                V.Prenume,
                V.Marca,
                V.NumeIntreg,
                V.DenumireGradDidactic,
                V.DenumireCatedra   AS Departament,
                V.DenumireFacultate AS Facultate
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] AS V
            WHERE V.ID_AnUnivCatedra = @idAn
              AND V.TitularAnUniv = 0
              AND (@idFac     = 0 OR V.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR V.ID_Catedra   = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [AGSIS].[pi].[Post] AS P
                  INNER JOIN [AGSIS].[pi].[Post_Profesor] AS PP ON P.ID_Post = PP.ID_Post
                  INNER JOIN [AGSIS].[pi].[StatDeFunctii] AS SF ON P.ID_StatDeFunctii = SF.ID_StatDeFunctii
                  WHERE PP.ID_Profesor         = V.ID_Profesor
                    AND SF.ID_AnUniv           = @idAn
                    AND P.TitularSauSuplinitor = 0
                    AND P.Deleted              = 0
                    AND PP.Deleted             = 0
              )
            ORDER BY V.Nume COLLATE Romanian_CI_AS, V.Prenume COLLATE Romanian_CI_AS";

        [HttpGet("colaboratori")]
        public async Task<IActionResult> GetColaboratori(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColaboratoriMaster, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["NumeIntreg"]?.ToString() ?? "",
                    Nume = r["Nume"]?.ToString() ?? "",
                    Prenume = r["Prenume"]?.ToString() ?? "",
                    Marca = r["Marca"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    Grad = r["DenumireGradDidactic"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/colaboratori")]
        public async Task<IActionResult> ExportColaboratori(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)),
                new DataColumn("Nume"), new DataColumn("Prenume"), new DataColumn("Marca"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad Didactic")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColaboratoriMaster, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["Nume"]?.ToString(),
                    r["Prenume"]?.ToString(),
                    r["Marca"]?.ToString(),
                    r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(),
                    r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Colaboratori");
            ws.Cell(1, 1).Value = $"Asociati / Colaboratori | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count)); ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Colaboratori.xlsx");
        }

        // =====================================================================
        // RAPORT ANS
        //
        // FIX MAJOR: SqlOreANS rescris complet dupa query-ul validat in SSMS.
        //
        // Probleme din versiunea anterioara:
        // 1. Filtru AND vc.NrSemestruDinAn = 2 — excludea semestrul 1 complet.
        //    Fix: eliminat filtrul de semestru; sumeaza ambele semestre per professor/ramura.
        //
        // 2. Coloana ID_Ramura in temp table se numeste ID_RamuraDeStiinta_ANS,
        //    dar query-ul anterior o redenumea ca ID_Ramura si apoi LoadOreANSAsync
        //    citea cheia "ID_ANS". Acum selectam explicit ID_RamuraDeStiinta_ANS AS ID_ANS.
        //
        // 3. Numitorul fractiei ANS: query-ul validat foloseste
        //    MAX(TotalRealizat, NormaLegala) — daca profesorul depaseste norma,
        //    numitorul devine TotalRealizat, nu norma fixa.
        //    Anterior, C# imparte direct la t.Norma (norma fixa).
        //    Fix: calculul fractiei mutat in SQL cu CASE WHEN TotalRealizat > NormaLegala.
        //
        // 4. Filtrul fds.id_metaspecializare > 0 si IS NOT NULL pe ID_N_Domeniu_Studiu_ANS
        //    erau prezente si raman.
        // =====================================================================
        private const string SqlCreateTempDomeniu = @"
            IF OBJECT_ID('tempdb..#DomeniuANS') IS NOT NULL DROP TABLE #DomeniuANS;
            CREATE TABLE #DomeniuANS (
                ID_ELEMENT INT, COD_DS_CNATDCU NVARCHAR(20), cod_DS NVARCHAR(20),
                ID_RamuraDeStiinta_ANS INT, DomeniulDeStudiu_ANS NVARCHAR(200),
                RamuraDeStiinta_ANS NVARCHAR(200), DomeniuFundamental NVARCHAR(200)
            );
            INSERT INTO #DomeniuANS EXEC [AGSIS].[dbo].[N_DOMENIU_STUDIUL_ANS_List];";

        // =====================================================================
        // SqlOreANS — rescris fidel fata de query-ul validat in SSMS.
        //
        // Logica:
        //   TitulariActivi    — titularii activi cu norma legala
        //   OreRaw            — suma ore conventionale per (profesor, ramura ANS)
        //                       folosind View_CentralizareMateriiProfesor + View_FDS + #DomeniuANS
        //                       JOIN-ul View_FDS->DomeniuANS se face pe ID_N_Domeniu_Studiu_ANS
        //                       si produce ID_RamuraDeStiinta_ANS (indexul coloanei 1-40)
        //   BazaCalcul        — adauga TotalRealizat per profesor si NormaLegala din DB
        //   SELECT final      — fractia = OreRamura / MAX(TotalRealizat, NormaLegala)
        //                       identic cu raportarea ministerului
        // =====================================================================
        private const string SqlOreANS = @"
            WITH TitulariActivi AS (
                SELECT DISTINCT
                    V.ID_Profesor,
                    V.NumeIntreg,
                    V.ID_TipGradDidacticAnUniv,
                    V.DenumireFacultate
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] AS V
                WHERE V.ID_AnUnivCatedra = @idAn
                  AND V.TitularAnUniv = 1
                  AND EXISTS (
                      SELECT 1
                      FROM [AGSIS].[pi].[Post] AS P
                      INNER JOIN [AGSIS].[pi].[Post_Profesor] AS PP ON P.ID_Post = PP.ID_Post
                      INNER JOIN [AGSIS].[pi].[StatDeFunctii] AS SF ON P.ID_StatDeFunctii = SF.ID_StatDeFunctii
                      WHERE PP.ID_Profesor = V.ID_Profesor
                        AND SF.ID_AnUniv = @idAn
                        AND P.TitularSauSuplinitor = 1
                        AND P.Deleted = 0
                        AND PP.Deleted = 0
                  )
            ),
            OreRaw AS (
                SELECT
                    T.ID_Profesor,
                    T.NumeIntreg,
                    T.DenumireFacultate,
                    T.ID_TipGradDidacticAnUniv,
                    M.ID_RamuraDeStiinta_ANS     AS ID_ANS,
                    M.RamuraDeStiinta_ANS         AS RamuraNume,
                    SUM(vc.NrOreConventionale)    AS OreRamura
                FROM TitulariActivi T
                INNER JOIN [AGSIS].[pi].[View_CentralizareMateriiProfesor] vc
                    ON T.ID_Profesor = vc.ID_Profesor
                INNER JOIN [AGSIS].[dbo].[View_FDS] fds
                    ON fds.DenumireSpecializare COLLATE DATABASE_DEFAULT
                       = vc.DenumireSpecializare COLLATE DATABASE_DEFAULT
                   AND fds.ID_AnUniv = @idAn
                   AND fds.id_metaspecializare > 0
                   AND fds.ID_N_Domeniu_Studiu_ANS IS NOT NULL
                INNER JOIN #DomeniuANS M
                    ON M.ID_ELEMENT = fds.ID_N_Domeniu_Studiu_ANS
                WHERE vc.ID_AnUniv = @idAn
                  AND vc.NrCrtPostProfesor = 1
                GROUP BY T.ID_Profesor, T.NumeIntreg, T.DenumireFacultate,
                         T.ID_TipGradDidacticAnUniv, M.ID_RamuraDeStiinta_ANS, M.RamuraDeStiinta_ANS
            ),
            BazaCalcul AS (
                SELECT
                    o.*,
                    SUM(o.OreRamura) OVER (PARTITION BY o.ID_Profesor) AS TotalRealizat,
                    ISNULL(n.NrOreConventionaleTitular, 15.0)           AS NormaLegala
                FROM OreRaw o
                LEFT JOIN [AGSIS].[pi].[NormaOreConventionale] n
                    ON o.ID_TipGradDidacticAnUniv = n.ID_TipGradDidactic
                   AND n.ID_AnUniv = @idAn
            )
            SELECT
                bc.ID_Profesor,
                bc.ID_ANS,
                CAST(bc.OreRamura AS DECIMAL(18,6))                         AS OreConvTotal,
                CAST(
                    ROUND(
                        bc.OreRamura / CASE
                            WHEN bc.TotalRealizat > bc.NormaLegala
                            THEN bc.TotalRealizat
                            ELSE bc.NormaLegala
                        END,
                        2
                    ) AS DECIMAL(10,4)
                )                                                           AS FractieANS
            FROM BazaCalcul bc
            WHERE bc.OreRamura > 0
            ORDER BY bc.NumeIntreg";

        // FIX: aliniat la logica EXISTS din pi.Post (identic cu SqlTitulariMaster).
        // Eliminat:
        //   - filtru ID_Facultate != 41 (DPPD): nu mai excludem DPPD daca nu se cere
        //   - filtru ID_TipGradDidacticAnUniv IN (1,2,3,10,11): tinea afaratura asistentilor
        // Asa raportul ANS vede ACELEASI cadre didactice ca raportul Titulari.
        private const string SqlTitulariANS = @"
            SELECT DISTINCT
                p.ID_Profesor, p.NumeIntreg, p.DenumireFacultate, p.DenumireCatedra,
                p.DenumireGradDidactic, p.ID_TipGradDidacticAnUniv AS ID_TipGrad,
                ISNULL(n.NrOreConventionaleTitular, 0)   AS NormaDB,
                p.CNP,
                CASE WHEN LEN(LTRIM(RTRIM(ISNULL(p.CNP,'')))) = 13
                          AND ISNUMERIC(SUBSTRING(p.CNP,1,3)) = 1
                     THEN 2026 - (CASE WHEN SUBSTRING(p.CNP,1,1) IN ('1','2')
                                       THEN 1900 ELSE 2000 END
                                  + CAST(SUBSTRING(p.CNP,2,2) AS INT))
                     ELSE NULL END AS Varsta,
                CASE WHEN p.TitularAnUniv = 1
                          AND p.DenumireGradDidactic NOT LIKE N'%Asociat%'
                          AND p.DenumireGradDidactic NOT LIKE N'%asociat%' THEN N'1'
                     WHEN p.DenumireGradDidactic LIKE N'%Asociat%'
                       OR p.DenumireGradDidactic LIKE N'%asociat%'         THEN N'2'
                     ELSE N'3' END AS FormaAngajare
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            LEFT JOIN [AGSIS].[pi].[NormaOreConventionale] n
                ON n.ID_TipGradDidactic = p.ID_TipGradDidacticAnUniv AND n.ID_AnUniv = @idAn
            WHERE p.ID_AnUnivCatedra = @idAn
              AND p.TitularAnUniv = 1
              AND (@idFac     = 0 OR p.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR p.ID_Catedra   = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [AGSIS].[pi].[Post] AS P
                  INNER JOIN [AGSIS].[pi].[Post_Profesor] AS PP ON P.ID_Post = PP.ID_Post
                  INNER JOIN [AGSIS].[pi].[StatDeFunctii] AS SF ON P.ID_StatDeFunctii = SF.ID_StatDeFunctii
                  WHERE PP.ID_Profesor         = p.ID_Profesor
                    AND SF.ID_AnUniv           = @idAn
                    AND P.TitularSauSuplinitor = 1
                    AND P.Deleted              = 0
                    AND PP.Deleted             = 0
              )
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

        private static decimal NormaEfectiva(decimal normaDB, int? idTipGrad) =>
            normaDB > 0 ? normaDB : idTipGrad switch
            {
                1 => 11m,
                2 => 12m,
                3 => 14m,
                4 => 15m,
                10 => 14m,
                11 => 14m,
                _ => 15m
            };

        private async Task<List<(int Id, string Den)>> LoadDomeniiAsync(SqlConnection conn)
        {
            var raw = new List<(int Id, string Den)>();
            using var cmd = new SqlCommand(
                "SELECT ID_Element, Denumire FROM [AGSIS].[dbo].[N_RAMURA_STIINTA_ANS] ORDER BY ID_Element", conn);
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                raw.Add((Convert.ToInt32(r["ID_Element"]), r["Denumire"]?.ToString() ?? ""));

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
                int? idTip = r["ID_TipGrad"] == DBNull.Value ? null : (int?)Convert.ToInt32(r["ID_TipGrad"]);
                lst.Add(new TitAns
                {
                    IdProf = Convert.ToInt32(r["ID_Profesor"]),
                    Nume = r["NumeIntreg"]?.ToString() ?? "",
                    Fac = r["DenumireFacultate"]?.ToString() ?? "",
                    Dept = r["DenumireCatedra"]?.ToString() ?? "",
                    GradD = r["DenumireGradDidactic"]?.ToString() ?? "",
                    IdTip = idTip,
                    Norma = NormaEfectiva(nDb, idTip),
                    Cnp = r["CNP"]?.ToString() ?? "",
                    Varsta = r["Varsta"] == DBNull.Value ? null : (int?)Convert.ToInt32(r["Varsta"]),
                    FormaAngajare = r["FormaAngajare"]?.ToString() ?? "3"
                });
            }
            return lst;
        }

        // =====================================================================
        // LoadOreANSAsync — acum citeste si FractieANS calculata in SQL
        // Returneaza doua dictionare:
        //   oreDict  — (idProf -> (idRamura -> OreConvTotal))   pentru export Excel
        //   fracDict — (idProf -> (idRamura -> FractieANS))     pentru API JSON si calcul total
        // =====================================================================
        private async Task<(Dictionary<int, Dictionary<int, decimal>> Ore,
                             Dictionary<int, Dictionary<int, decimal>> Frac)>
            LoadOreANSAsync(SqlConnection conn, int idAn)
        {
            var oreDict = new Dictionary<int, Dictionary<int, decimal>>();
            var fracDict = new Dictionary<int, Dictionary<int, decimal>>();

            using (var cmdT = new SqlCommand(SqlCreateTempDomeniu, conn))
            {
                cmdT.CommandTimeout = 60;
                await cmdT.ExecuteNonQueryAsync();
            }

            using var cmd = new SqlCommand(SqlOreANS, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@idAn", idAn);

            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                int idP = Convert.ToInt32(r["ID_Profesor"]);
                int idR = Convert.ToInt32(r["ID_ANS"]);
                decimal ore = r["OreConvTotal"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConvTotal"]);
                decimal frac = r["FractieANS"] == DBNull.Value ? 0m : Convert.ToDecimal(r["FractieANS"]);

                if (!oreDict.ContainsKey(idP)) oreDict[idP] = new();
                if (!fracDict.ContainsKey(idP)) fracDict[idP] = new();

                if (!oreDict[idP].ContainsKey(idR)) oreDict[idP][idR] = 0m;
                if (!fracDict[idP].ContainsKey(idR)) fracDict[idP][idR] = 0m;

                oreDict[idP][idR] += ore;
                fracDict[idP][idR] += frac;
            }

            return (oreDict, fracDict);
        }

        [HttpGet("raport-ans")]
        public async Task<IActionResult> GetAns(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();

            var domenii = await LoadDomeniiAsync(conn);
            var titulari = await LoadTitulariANSAsync(conn, idAn, idFacultate ?? 0, idCatedra ?? 0);
            var (_, fracDict) = await LoadOreANSAsync(conn, idAn);

            var profesori = new List<object>();
            int nrCrt = 1;
            foreach (var t in titulari)
            {
                var frac = new Dictionary<string, decimal>();
                if (fracDict.TryGetValue(t.IdProf, out var fracP))
                    foreach (var kv in fracP)
                    {
                        var dom = domenii.FirstOrDefault(d => d.Id == kv.Key);
                        if (dom.Den == null) continue;
                        if (kv.Value > 0) frac[dom.Den] = kv.Value;
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
                    NormaEfectiva = t.Norma,
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
            var (_, fracDict) = await LoadOreANSAsync(conn, idAn);

            int nrD = domenii.Count, colTot = 9 + nrD + 1;
            var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("CD DRU");
            var gC = XLColor.FromHtml(Green);

            ws.Cell(1, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(1, 1, 1, colTot).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(2, 1).Value = "Universitatea Transilvania din Brasov";
            ws.Range(2, 1, 2, 6).Merge();

            string[] antet = { "Nr.\nCrt.", "Nume si prenume", "CNP", "Functie",
                               "Forma\nangajare", "Cond.\ndoctorat", "Varsta", "Facultate", "Departament" };
            for (int c = 1; c <= 9; c++) { ws.Cell(4, c).Value = antet[c - 1]; ws.Range(4, c, 6, c).Merge(); }
            ws.Cell(4, colTot).Value = "Total"; ws.Range(4, colTot, 6, colTot).Merge();

            for (int i = 0; i < nrD; i++)
            {
                ws.Cell(5, 10 + i).Value = domenii[i].Den;
                ws.Range(5, 10 + i, 6, 10 + i).Merge();
            }

            for (int c = 1; c <= 9; c++) ws.Cell(7, c).Value = ((char)('A' + c - 1)).ToString();
            for (int i = 0; i < nrD; i++) ws.Cell(7, 10 + i).Value = i + 1;
            ws.Cell(7, colTot).Value = nrD;

            for (int row = 4; row <= 7; row++)
                for (int col = 1; col <= colTot; col++)
                {
                    var cell = ws.Cell(row, col);
                    cell.Style.Font.Bold = true;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.WrapText = true;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            ws.Range(4, 1, 4, colTot).Style.Fill.BackgroundColor = gC;
            ws.Range(4, 1, 4, colTot).Style.Font.FontColor = XLColor.White;

            ws.Column(1).Width = 5; ws.Column(2).Width = 30; ws.Column(3).Width = 14;
            ws.Column(4).Width = 22; ws.Column(5).Width = 10; ws.Column(6).Width = 12;
            ws.Column(7).Width = 8; ws.Column(8).Width = 28; ws.Column(9).Width = 28;
            for (int c = 10; c <= colTot; c++) ws.Column(c).Width = 12;

            int rowExcel = 8, nrCrt = 1;
            foreach (var t in titulari)
            {
                ws.Cell(rowExcel, 1).Value = nrCrt++;
                ws.Cell(rowExcel, 2).Value = t.Nume;
                ws.Cell(rowExcel, 3).Value = t.Cnp;
                ws.Cell(rowExcel, 4).Value = t.GradD;
                ws.Cell(rowExcel, 5).Value = t.FormaAngajare;
                ws.Cell(rowExcel, 6).Value = "";
                if (t.Varsta.HasValue) ws.Cell(rowExcel, 7).Value = t.Varsta.Value;
                else ws.Cell(rowExcel, 7).Value = "";
                ws.Cell(rowExcel, 8).Value = t.Fac;
                ws.Cell(rowExcel, 9).Value = t.Dept;

                decimal totFrac = 0m;
                if (fracDict.TryGetValue(t.IdProf, out var fracP))
                    for (int i = 0; i < nrD; i++)
                        if (fracP.TryGetValue(domenii[i].Id, out decimal f) && f > 0)
                        {
                            ws.Cell(rowExcel, 10 + i).Value = (double)f;
                            ws.Cell(rowExcel, 10 + i).Style.NumberFormat.Format = "0.00";
                            totFrac += f;
                        }

                ws.Cell(rowExcel, colTot).Value = (double)Math.Round(totFrac, 2);
                ws.Cell(rowExcel, colTot).Style.NumberFormat.Format = "0.00";
                for (int c = 1; c <= colTot; c++)
                    ws.Cell(rowExcel, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rowExcel++;
            }

            using var stream = new MemoryStream(); wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Raport_ANS.xlsx");
        }
    }
}