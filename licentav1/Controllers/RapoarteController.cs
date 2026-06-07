using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using System.IO.Compression;
using ClosedXML.Excel;

namespace LicentaV1.Controllers
{
   
    
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
        private const int SaptamaniPerAn = 14;     
        private const decimal NormaLegalaFallback = 15.0m;  
        private const int TimeoutShort = 60;
        private const int TimeoutMedium = 120;
        private const int TimeoutLong = 180;

        public RapoarteController(IConfiguration cfg, IMemoryCache cache)
        {
            _cs = cfg.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        
        private int GetAnCurent() => _cache.GetOrCreate("AnCurent_v9", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(4);
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT TOP 1 au.ID_AnUniv
                FROM [dbo].[AnUniversitar] au
                WHERE EXISTS (SELECT 1 FROM [pi].[View_PostProfesorMaterie] ppm
                              WHERE ppm.ID_AnUniv = au.ID_AnUniv)
                ORDER BY au.Ordine DESC", conn);
            var r = cmd.ExecuteScalar();
            if (r == null || r == DBNull.Value)
                throw new InvalidOperationException(
                    "Nu am putut determina anul universitar curent. " +
                    "Tabelul AnUniversitar nu are randuri legate de View_PostProfesorMaterie.");
            return Convert.ToInt32(r);
        });

        private void StyleHdr(IXLRange r)
        {
            r.Style.Fill.BackgroundColor = XLColor.FromHtml(Green);
            r.Style.Font.FontColor = XLColor.White;
            r.Style.Font.Bold = true;
        }

        private string SafeFile(string raport, string? profesor)
        {
            if (string.IsNullOrWhiteSpace(profesor) || profesor == "Toti") return $"{raport}.xlsx";
            var safe = string.Concat(profesor.Take(35).Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
            return $"{safe}_{raport}.xlsx";
        }

        
        [HttpGet("liste/ani-universitari")]
        public IActionResult GetAni() => Ok(_cache.GetOrCreate("AniUniv_v8", e =>
        {
            e.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
            var lst = new List<object>();
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT au.ID_AnUniv, LTRIM(RTRIM(au.Denumire)) AS Denumire
                FROM [dbo].[AnUniversitar] au
                WHERE EXISTS (SELECT 1 FROM [pi].[View_PostProfesorMaterie] ppm
                              WHERE ppm.ID_AnUniv = au.ID_AnUniv)
                ORDER BY au.Ordine DESC", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) lst.Add(new { id = Convert.ToInt32(r[0]), nume = r[1]?.ToString() ?? "" });
            return lst;
        }));

        [HttpGet("liste/facultati")]
        public IActionResult GetFacultati([FromQuery] int? idAnUniv)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = 0, nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT ID_Facultate, DenumireFacultate
                FROM [dbo].[View_Profesori_CF_AnUniv]
                WHERE ID_AnUnivCatedra = @idAn AND DenumireFacultate IS NOT NULL
                ORDER BY DenumireFacultate", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                lst.Add(new { id = Convert.ToInt32(r["ID_Facultate"]), nume = r["DenumireFacultate"]?.ToString() ?? "" });
            return Ok(lst);
        }

        [HttpGet("liste/departamente")]
        public IActionResult GetDepartamente([FromQuery] int? idAnUniv, [FromQuery] int? idFacultate)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = 0, nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT ID_Catedra, LTRIM(RTRIM(DenumireCatedra)) AS DenumireCatedra
                FROM [dbo].[View_Profesori_CF_AnUniv]
                WHERE ID_AnUnivCatedra = @idAn
                  AND (@idFac = 0 OR ID_Facultate = @idFac)
                  AND DenumireCatedra IS NOT NULL
                  AND LTRIM(RTRIM(DenumireCatedra)) != ''
                ORDER BY DenumireCatedra", conn);
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
                    LTRIM(RTRIM(
                        CASE WHEN CHARINDEX(N'+', DenumireSpecializare) > 0
                             THEN LEFT(DenumireSpecializare, CHARINDEX(N'+', DenumireSpecializare)-1)
                             ELSE DenumireSpecializare END
                    )) AS Spec
                FROM [pi].[View_PostProfesorMaterie]
                WHERE ID_AnUniv = @idAn
                  AND (@idFac = 0 OR ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR ID_Catedra = @idCatedra)
                  AND DenumireSpecializare IS NOT NULL
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

        [HttpGet("liste/cicluri-studii")]
        public IActionResult GetCicluri([FromQuery] int? idAnUniv)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = "Toti", nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT DenumireCicluInv FROM [dbo].[View_FDS]
                WHERE ID_AnUniv = @idAn AND DenumireCicluInv IS NOT NULL
                ORDER BY DenumireCicluInv", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var v = r[0]?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(v)) lst.Add(new { id = v, nume = v });
            }
            return Ok(lst);
        }

        [HttpGet("liste/forme-invatamant")]
        public IActionResult GetFormeInv([FromQuery] int? idAnUniv)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var lst = new List<object> { new { id = "Toti", nume = "Toate" } };
            using var conn = new SqlConnection(_cs); conn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT DenumireFormaInv FROM [dbo].[View_FDS]
                WHERE ID_AnUniv = @idAn AND DenumireFormaInv IS NOT NULL
                ORDER BY DenumireFormaInv", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var v = r[0]?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(v)) lst.Add(new { id = v, nume = v });
            }
            return Ok(lst);
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
                SELECT DISTINCT LTRIM(RTRIM(p.NumeIntreg)) AS NumeIntreg
                FROM [dbo].[View_Profesori_CF_AnUniv] p
                INNER JOIN [pi].[View_PostProfesorMaterie] vppm
                    ON p.ID_Profesor = vppm.ID_Profesor AND p.ID_AnUnivCatedra = vppm.ID_AnUniv
                WHERE p.ID_AnUnivCatedra = @idAn
                  AND p.NumeIntreg IS NOT NULL AND LTRIM(RTRIM(p.NumeIntreg)) != ''
                  AND (@idFac = 0 OR p.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR p.ID_Catedra = @idCatedra)
                  AND (@spec = N'Toti' OR
                        LTRIM(RTRIM(
                            CASE WHEN CHARINDEX(N'+', vppm.DenumireSpecializare) > 0
                                 THEN LEFT(vppm.DenumireSpecializare, CHARINDEX(N'+', vppm.DenumireSpecializare)-1)
                                 ELSE vppm.DenumireSpecializare END
                        )) COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
                  AND (@tipPost = N'Toti' OR
                        CASE vppm.TitularSauSuplinitor WHEN 1 THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                ORDER BY NumeIntreg", conn);
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var n = r[0]?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(n)) lst.Add(new { id = n, nume = n });
            }
            return Ok(lst);
        }

        
        // RAPORT 1
        private const string SqlNorma = @"
            WITH OreUnice AS (
                SELECT
                    vppm.ID_Profesor,
                    vppm.Denumire AS Materie,
                    vppm.DenumireSpecializare,
                    vppm.DenumireScurtaSpecializare,
                    vppm.NrSemestruDinAn,
                    vppm.Nr_Ore_Curs,
                    vppm.Nr_Ore_Seminar,
                    vppm.Nr_Ore_Laborator,
                    vppm.Nr_Ore_Proiect,
                    vppm.NrOreConventionale,
                    vppm.ApartineDeCuplaj,
                    vppm.ID_PlanMaterie_Prestator,
                    vppm.ID_Facultate,
                    vppm.ID_Catedra,
                    vppm.ID_TipFormaInv,
                    vppm.id_specializare,
                    vppm.TitularSauSuplinitor,
                    ROW_NUMBER() OVER (
                        PARTITION BY vppm.ID_Profesor, vppm.ID_PlanMaterie_Prestator, vppm.NrSemestruDinAn
                        ORDER BY vppm.DenumireSpecializare
                    ) AS Rn
                FROM [pi].[View_PostProfesorMaterie] vppm
                WHERE vppm.ID_AnUniv = @idAn
            ),
            -- Maparea specializare → ciclu+forma. View_FDS poate avea acelasi
            -- ID_Specializare in mai multi ani; luam cea mai recenta cu MAX().
            FdsSpec AS (
                SELECT ID_Specializare,
                       MAX(DenumireCicluInv) AS DenCiclu,
                       MAX(ID_TipFormaInv)   AS IdFormaFds
                FROM [dbo].[View_FDS]
                GROUP BY ID_Specializare
            )
            SELECT
                p.NumeIntreg                             AS Profesor,
                p.DenumireFacultate                      AS Facultate,
                p.DenumireCatedra                        AS Departament,
                p.DenumireGradDidactic                   AS Grad,
                CASE ou.TitularSauSuplinitor
                    WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END AS TipPost,
                ou.Materie                               AS Materie,
                ou.NrSemestruDinAn                       AS Semestru,
                ou.DenumireSpecializare                  AS Specializare,
                CAST(ISNULL(ou.Nr_Ore_Curs, 0) AS DECIMAL(10,2))                                            AS OreCurs,
                CAST((ISNULL(ou.Nr_Ore_Seminar, 0) + ISNULL(ou.Nr_Ore_Laborator, 0) + ISNULL(ou.Nr_Ore_Proiect, 0)) AS DECIMAL(10,2)) AS OreAplic,
                CAST(ou.NrOreConventionale AS DECIMAL(10,2)) AS OreConv,
                CASE
                    WHEN ou.ApartineDeCuplaj IS NOT NULL THEN
                        N'Cuplat cu: ' + ISNULL(STUFF((
                            SELECT DISTINCT N', ' + vdc.DenumireScurtaSpecializare
                            FROM [pi].[View_DetaliereCuplaje] vdc
                            WHERE vdc.ID_Cuplaj = ou.ApartineDeCuplaj
                              AND vdc.ID_AnUniv = @idAn
                              AND vdc.DenumireScurtaSpecializare <> ou.DenumireScurtaSpecializare
                            FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N''),
                            N'incarcare comuna')
                    ELSE N'Individual'
                END AS Mentiuni
            FROM OreUnice ou
            INNER JOIN [dbo].[View_Profesori_CF_AnUniv] p
                ON ou.ID_Profesor = p.ID_Profesor AND p.ID_AnUnivCatedra = @idAn
            LEFT JOIN FdsSpec fs
                ON fs.ID_Specializare = ou.id_specializare
            WHERE ou.Rn = 1
              AND (@idFac     = 0 OR ou.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR ou.ID_Catedra   = @idCatedra)
              AND (@spec      = N'Toti' OR
                   LTRIM(RTRIM(
                     CASE WHEN CHARINDEX(N'+', ou.DenumireSpecializare) > 0
                          THEN LEFT(ou.DenumireSpecializare, CHARINDEX(N'+', ou.DenumireSpecializare)-1)
                          ELSE ou.DenumireSpecializare END
                   )) COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
              AND (@prof      = N'Toti' OR p.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
              AND (@tipPost   = N'Toti' OR CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
              AND (@sem       = 0       OR ou.NrSemestruDinAn = @sem)
              -- Forma de invatamant: comparare case+accent insensitive cu denumirea
              -- din BD (vppm.ID_TipFormaInv din pi.View_PostProfesorMaterie e numeric,
              -- dar acelasi ID intoarce View_FDS la coloana ID_TipFormaInv. Comparam
              -- prin sub-select din View_FDS unde avem si denumirea).
              AND (@formaInv  = N'Toti' OR ou.ID_TipFormaInv IN (
                    SELECT DISTINCT fds2.ID_TipFormaInv FROM [dbo].[View_FDS] fds2
                    WHERE fds2.DenumireFormaInv COLLATE Latin1_General_CI_AI = @formaInv COLLATE Latin1_General_CI_AI
                  ))
              -- Ciclu studii: comparare case+accent insensitive cu DenumireCicluInv din View_FDS
              AND (@ciclu     = N'Toti' OR
                   fs.DenCiclu COLLATE Latin1_General_CI_AI = @ciclu COLLATE Latin1_General_CI_AI)
            ORDER BY p.NumeIntreg COLLATE Romanian_CI_AS, ou.NrSemestruDinAn";

        private void AddNormaParams(SqlCommand cmd, int idAn, int idFac, int idCat,
            string prof, string spec, string tipPost, int sem, string formaInv, string ciclu)
        {
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFac);
            cmd.Parameters.AddWithValue("@idCatedra", idCat);
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
            using var cmd = new SqlCommand(SqlNorma, conn); cmd.CommandTimeout = TimeoutLong;
            AddNormaParams(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0,
                profesor ?? "", specializare ?? "", tipPost ?? "Toti", semestru ?? 0,
                formaInvatamant ?? "", cicluStudii ?? "");
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["Profesor"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    Grad = r["Grad"]?.ToString() ?? "",
                    TipPost = r["TipPost"]?.ToString() ?? "",
                    Specializare = r["Specializare"]?.ToString() ?? "",
                    Materie = r["Materie"]?.ToString() ?? "",
                    Semestru = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
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
            [FromQuery] int? semestru, [FromQuery] string? formaInvatamant, [FromQuery] string? cicluStudii)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad"),
                new DataColumn("Tip Post"), new DataColumn("Specializare"),
                new DataColumn("Materie"), new DataColumn("Semestru", typeof(int)),
                new DataColumn("Ore Curs", typeof(decimal)),
                new DataColumn("Ore Aplic.", typeof(decimal)), new DataColumn("Ore Conv.", typeof(decimal)),
                new DataColumn("Mentiuni")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma, conn); cmd.CommandTimeout = TimeoutLong;
            AddNormaParams(cmd, idAn, idFacultate ?? 0, idCatedra ?? 0,
                profesor ?? "", specializare ?? "", tipPost ?? "Toti", semestru ?? 0,
                formaInvatamant ?? "", cicluStudii ?? "");
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["Profesor"]?.ToString(), r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(), r["Grad"]?.ToString(),
                    r["TipPost"]?.ToString(), r["Specializare"]?.ToString(),
                    r["Materie"]?.ToString(),
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["OreCurs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreCurs"]),
                    r["OreAplic"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreAplic"]),
                    r["OreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreConv"]),
                    r["Mentiuni"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Norma Profesori");
            ws.Cell(1, 1).Value = "Detaliere Norme Profesori";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Aplic.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                SafeFile("Norma_Profesori", profesor));
        }

        
        // RAPORT 2
        private const string SqlTotaluri = @"
            WITH OreUnice AS (
                SELECT
                    vppm.ID_Profesor,
                    vppm.ID_PlanMaterie_Prestator,
                    vppm.NrSemestruDinAn,
                    vppm.NrOreConventionale,
                    vppm.DenumireSpecializare,
                    vppm.ID_Facultate,
                    vppm.ID_Catedra,
                    vppm.TitularSauSuplinitor,
                    CASE
                        WHEN UPPER(vppm.DenumireSpecializare) LIKE N'%IFR%'
                          OR UPPER(vppm.DenumireSpecializare) LIKE N'% FR%'
                          OR UPPER(vppm.DenumireSpecializare) LIKE N'%-FR%' THEN N'IFR'
                        WHEN UPPER(vppm.DenumireSpecializare) LIKE N'% ID%'
                          OR UPPER(vppm.DenumireSpecializare) LIKE N'%-ID'
                          OR UPPER(vppm.DenumireSpecializare) LIKE N'%-ID %'
                          OR UPPER(vppm.DenumireSpecializare) LIKE N'% ID' THEN N'ID'
                        ELSE N'IF'
                    END AS FormaInv,
                    ROW_NUMBER() OVER (
                        PARTITION BY vppm.ID_Profesor, vppm.ID_PlanMaterie_Prestator, vppm.NrSemestruDinAn
                        ORDER BY vppm.DenumireSpecializare
                    ) AS Rn
                FROM [pi].[View_PostProfesorMaterie] vppm
                WHERE vppm.ID_AnUniv = @idAn
            ),
            -- Catedra → Facultate (pentru a afisa numele postului unde se preda).
            -- Profesori_CF_AnUniv contine maparea Catedra → Facultate cu denumirile.
            CatedraInfo AS (
                SELECT DISTINCT ID_Catedra, DenumireCatedra, ID_Facultate, DenumireFacultate
                FROM [dbo].[View_Profesori_CF_AnUniv]
                WHERE ID_AnUnivCatedra = @idAn
            )
            SELECT
                p.NumeIntreg                                                         AS Profesor,
                -- Facultate/Departament = unde se preda efectiv ora (din POSTUL vppm),
                -- nu unde e angajat profesorul. Cand un prof are mai multe posturi in
                -- departamente diferite, fiecare apare ca rand separat pe acel dept.
                MAX(ci.DenumireFacultate)                                            AS Facultate,
                MAX(ci.DenumireCatedra)                                              AS Departament,
                MAX(p.DenumireGradDidactic)                                          AS Grad,
                CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END AS TipPost,
                CAST(SUM(CASE WHEN ou.FormaInv = N'IF'  THEN ou.NrOreConventionale ELSE 0 END) AS DECIMAL(10,2)) AS OreIF,
                CAST(SUM(CASE WHEN ou.FormaInv = N'ID'  THEN ou.NrOreConventionale ELSE 0 END) AS DECIMAL(10,2)) AS OreID,
                CAST(SUM(CASE WHEN ou.FormaInv = N'IFR' THEN ou.NrOreConventionale ELSE 0 END) AS DECIMAL(10,2)) AS OreIFR,
                CAST(SUM(ou.NrOreConventionale) AS DECIMAL(10,2))      AS TotalOreConv,
                CAST(SUM(ou.NrOreConventionale) * @sapt AS DECIMAL(10,2)) AS TotalAnual
            FROM OreUnice ou
            INNER JOIN [dbo].[View_Profesori_CF_AnUniv] p
                ON ou.ID_Profesor = p.ID_Profesor AND p.ID_AnUnivCatedra = @idAn
            LEFT JOIN CatedraInfo ci
                ON ci.ID_Catedra = ou.ID_Catedra
            WHERE ou.Rn = 1
              AND (@idFac     = 0 OR ou.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR ou.ID_Catedra   = @idCatedra)
              AND (@prof      = N'Toti' OR p.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
              AND (@tipPost   = N'Toti' OR CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
              AND (@sem       = 0       OR ou.NrSemestruDinAn = @sem)
            -- GROUP BY pe (profesor, post-catedra, tip) astfel incat sa avem
            -- un rand separat per (profesor, departament_postului, Tit/Sup)
            GROUP BY p.NumeIntreg, ou.ID_Catedra, ou.TitularSauSuplinitor
            HAVING SUM(ou.NrOreConventionale) > 0
            ORDER BY p.NumeIntreg COLLATE Romanian_CI_AS";

        [HttpGet("norma-totaluri")]
        public async Task<IActionResult> GetTotaluri(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost, [FromQuery] int? semestru)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri, conn); cmd.CommandTimeout = TimeoutMedium;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@sapt", SaptamaniPerAn);
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["Profesor"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    Grad = r["Grad"]?.ToString() ?? "",
                    TipPost = r["TipPost"]?.ToString() ?? "",
                    OreIF = r["OreIF"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIF"]),
                    OreID = r["OreID"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreID"]),
                    OreIFR = r["OreIFR"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIFR"]),
                    TotalOreConv = r["TotalOreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreConv"]),
                    TotalAnual = r["TotalAnual"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalAnual"])
                });
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public async Task<IActionResult> ExportTotaluri(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? tipPost, [FromQuery] int? semestru)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"),
                new DataColumn("Grad"), new DataColumn("Tip Post"),
                new DataColumn("Ore IF",  typeof(decimal)),
                new DataColumn("Ore ID",  typeof(decimal)),
                new DataColumn("Ore IFR", typeof(decimal)),
                new DataColumn("Total Ore Conv.",   typeof(decimal)),
                new DataColumn("Total Anual (x14)", typeof(decimal))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTotaluri, conn); cmd.CommandTimeout = TimeoutMedium;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor!.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost!.Trim());
            cmd.Parameters.AddWithValue("@sapt", SaptamaniPerAn);
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["Profesor"], r["Facultate"], r["Departament"], r["Grad"], r["TipPost"],
                    r["OreIF"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIF"]),
                    r["OreID"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreID"]),
                    r["OreIFR"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreIFR"]),
                    r["TotalOreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreConv"]),
                    r["TotalAnual"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalAnual"]));
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Totaluri Norme");
            ws.Cell(1, 1).Value = $"Totaluri Norme | An: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            foreach (var c in new[] { "Ore IF", "Ore ID", "Ore IFR", "Total Ore Conv.", "Total Anual (x14)" })
                tbl.Field(c).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Totaluri_Norme.xlsx");
        }

        
        // RAPORT 3
        
        private const string SqlDistrib = @"
            WITH OreUnice AS (
                SELECT
                    vppm.ID_Profesor,
                    -- SpecCurata: denumirea pastrata pentru afisare (cu diacritice, dar
                    -- fara partea dupa '+').
                    LTRIM(RTRIM(
                        CASE WHEN CHARINDEX(N'+', vppm.DenumireSpecializare) > 0
                             THEN LEFT(vppm.DenumireSpecializare, CHARINDEX(N'+', vppm.DenumireSpecializare)-1)
                             ELSE vppm.DenumireSpecializare END
                    )) AS SpecCurata,
                    -- SpecKey: cheie de deduplicare. Aplicam:
                    --   1. UPPER → case-insensitive
                    --   2. COLLATE Latin1_General_CI_AI → ignora diacriticele
                    --      (ex: 'în' = 'in', 'ă' = 'a', 'ș' = 's')
                    --   3. REPLACE paranteze cu spatiu → 'X (Y)' = 'X Y'
                    --   4. REPLACE spatii multiple → unul singur
                    -- Asta uneste 'Informatica aplicata in limba germana' cu
                    -- 'Informatica aplicata (in limba germana)' si cu varianta cu diacritice.
                    LTRIM(RTRIM(
                        REPLACE(REPLACE(REPLACE(REPLACE(
                            UPPER(
                                CASE WHEN CHARINDEX(N'+', vppm.DenumireSpecializare) > 0
                                     THEN LEFT(vppm.DenumireSpecializare, CHARINDEX(N'+', vppm.DenumireSpecializare)-1)
                                     ELSE vppm.DenumireSpecializare END
                            ) COLLATE Latin1_General_CI_AI,
                            N'(', N' '), N')', N' '), N'  ', N' '), N'  ', N' ')
                    )) AS SpecKey,
                    vppm.ID_Facultate, vppm.ID_Catedra,
                    vppm.NrSemestruDinAn, vppm.NrOreConventionale,
                    vppm.TitularSauSuplinitor,
                    ROW_NUMBER() OVER (
                        PARTITION BY vppm.ID_Profesor, vppm.ID_PlanMaterie_Prestator, vppm.NrSemestruDinAn
                        ORDER BY vppm.DenumireSpecializare
                    ) AS Rn
                FROM [pi].[View_PostProfesorMaterie] vppm
                WHERE vppm.ID_AnUniv = @idAn
            ),
            OrePerProg AS (
                SELECT p.NumeIntreg, p.DenumireFacultate, p.DenumireCatedra,
                       ou.SpecKey,
                       MIN(ou.SpecCurata) AS Program,  -- alegem o reprezentare consistenta
                       SUM(ou.NrOreConventionale) AS OreProgram
                FROM OreUnice ou
                INNER JOIN [dbo].[View_Profesori_CF_AnUniv] p
                    ON ou.ID_Profesor = p.ID_Profesor AND p.ID_AnUnivCatedra = @idAn
                WHERE ou.Rn = 1
                  AND (@idFac     = 0 OR ou.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR ou.ID_Catedra   = @idCatedra)
                  AND (@prof      = N'Toti' OR p.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@spec      = N'Toti' OR ou.SpecCurata COLLATE DATABASE_DEFAULT = @spec COLLATE DATABASE_DEFAULT)
                  AND (@tipPost   = N'Toti' OR CASE ou.TitularSauSuplinitor WHEN 1 THEN N'Titular' ELSE N'Suplinitor' END = @tipPost)
                  AND (@sem       = 0       OR ou.NrSemestruDinAn = @sem)
                GROUP BY p.NumeIntreg, p.DenumireFacultate, p.DenumireCatedra, ou.SpecKey
            ),
            TotProf AS (
                SELECT NumeIntreg, SUM(OreProgram) AS TotalProf FROM OrePerProg GROUP BY NumeIntreg
            )
            SELECT o.NumeIntreg AS Profesor, o.DenumireFacultate AS Facultate,
                   o.DenumireCatedra AS Departament, o.Program,
                   CAST(o.OreProgram AS DECIMAL(10,2)) AS OreProgram,
                   CAST(t.TotalProf AS DECIMAL(10,2)) AS TotalProf,
                   CAST(CASE WHEN t.TotalProf > 0 THEN ROUND(o.OreProgram / t.TotalProf * 100, 2) ELSE 0 END AS DECIMAL(10,2)) AS Procent
            FROM OrePerProg o
            INNER JOIN TotProf t ON t.NumeIntreg = o.NumeIntreg
            WHERE o.OreProgram > 0
            ORDER BY o.NumeIntreg COLLATE Romanian_CI_AS, o.OreProgram DESC";

        [HttpGet("distributie-ore")]
        public async Task<IActionResult> GetDistrib(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? specializare,
            [FromQuery] string? tipPost, [FromQuery] int? semestru)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDistrib, conn); cmd.CommandTimeout = TimeoutMedium;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["Profesor"]?.ToString() ?? "",
                    Facultate = r["Facultate"]?.ToString() ?? "",
                    Departament = r["Departament"]?.ToString() ?? "",
                    Program = r["Program"]?.ToString() ?? "",
                    OreProgram = r["OreProgram"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreProgram"]),
                    TotalOreUniv = r["TotalProf"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalProf"]),
                    Procent = r["Procent"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Procent"])
                });
            return Ok(result);
        }

        [HttpGet("export/distributie-ore")]
        public async Task<IActionResult> ExportDistrib(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? specializare,
            [FromQuery] string? tipPost, [FromQuery] int? semestru)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Program Studiu"),
                new DataColumn("Ore Program", typeof(decimal)),
                new DataColumn("Total Ore Profesor", typeof(decimal)),
                new DataColumn("Procent %", typeof(decimal))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDistrib, conn); cmd.CommandTimeout = TimeoutMedium;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            cmd.Parameters.AddWithValue("@tipPost", string.IsNullOrWhiteSpace(tipPost) ? "Toti" : tipPost.Trim());
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["Profesor"]?.ToString(), r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(), r["Program"]?.ToString(),
                    r["OreProgram"] == DBNull.Value ? 0m : Convert.ToDecimal(r["OreProgram"]),
                    r["TotalProf"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalProf"]),
                    r["Procent"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Procent"]));
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Distributie Ore");
            ws.Cell(1, 1).Value = "Distributie Ore pe Programe";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Program").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                SafeFile("Distributie_Ore", profesor));
        }

        
        // RAPORT 4
        
        private const string SqlLimbi = @"
            WITH ProgrameStraine AS (
                SELECT DISTINCT
                    LTRIM(RTRIM(REPLACE(REPLACE(f.Specializare, CHAR(160), N' '), N'  ', N' ')))
                        COLLATE DATABASE_DEFAULT AS ProgramCurat,
                    LTRIM(RTRIM(f.Facultate)) COLLATE DATABASE_DEFAULT AS FacultateCurata,
                    LTRIM(RTRIM(f.DenumireSpecializare)) COLLATE DATABASE_DEFAULT AS NumeSistem,
                    LTRIM(RTRIM(f.LimbaPredare)) AS Limba,
                    -- Folosim DenumireCicluInv (acelasi camp pe care il populeaza
                    -- dropdown-ul din UI prin endpointul liste/cicluri-studii)
                    LTRIM(RTRIM(f.DenumireCicluInv)) AS Ciclu
                FROM [dbo].[View_FDS] f
                WHERE f.ID_AnUniv = @idAn
                  AND f.LimbaPredare IS NOT NULL
                  AND LTRIM(RTRIM(f.LimbaPredare)) NOT IN (N'RO', N'Ro', N'ro', N'Romana', N'Română')
                  AND f.id_metaspecializare > 0
            )
            SELECT
                p.NumeIntreg                                               AS Profesor,
                ps.Limba                                                   AS LimbaProgram,
                ps.ProgramCurat                                            AS ProgramStudiu,
                ps.FacultateCurata                                         AS Facultate,
                ps.Ciclu                                                   AS CicluStudii,
                vppm.NrSemestruDinAn                                       AS Semestru,
                CAST(SUM(vppm.NrOreConventionale) AS DECIMAL(10,2))        AS TotalOreConv
            FROM [pi].[View_PostProfesorMaterie] vppm
            INNER JOIN ProgrameStraine ps
                ON LTRIM(RTRIM(vppm.DenumireSpecializare)) COLLATE DATABASE_DEFAULT = ps.NumeSistem
            INNER JOIN [dbo].[View_Profesori_CF_AnUniv] p
                ON vppm.ID_Profesor = p.ID_Profesor AND p.ID_AnUnivCatedra = @idAn
            WHERE vppm.ID_AnUniv = @idAn
              AND (@idFac     = 0 OR p.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR p.ID_Catedra   = @idCatedra)
              AND (@prof      = N'Toti' OR p.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
              -- Compar case+accent insensitive ca sa nu pierdem din cauza variatiilor
              -- 'MASTER' vs 'Master' vs 'master'
              AND (@ciclu     = N'Toti' OR
                   ps.Ciclu COLLATE Latin1_General_CI_AI = @ciclu COLLATE Latin1_General_CI_AI)
            GROUP BY p.NumeIntreg, ps.Limba, ps.ProgramCurat, ps.FacultateCurata, ps.Ciclu, vppm.NrSemestruDinAn
            HAVING SUM(ISNULL(vppm.Nr_Ore_Curs, 0) + ISNULL(vppm.Nr_Ore_Seminar, 0) + ISNULL(vppm.Nr_Ore_Laborator, 0)) > 0
            ORDER BY ps.FacultateCurata, ps.ProgramCurat, p.NumeIntreg";

        [HttpGet("limbi-straine")]
        public async Task<IActionResult> GetLimbi(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? cicluStudii)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi, conn); cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@ciclu", string.IsNullOrWhiteSpace(cicluStudii) ? "Toti" : cicluStudii.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = r["Profesor"]?.ToString() ?? "",
                    Semestru = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    TotalOre = r["TotalOreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreConv"]),
                    LimbaProgram = r["LimbaProgram"]?.ToString() ?? "",
                    ProgramStudiu = r["ProgramStudiu"]?.ToString() ?? "",
                    CicluStudii = r["CicluStudii"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/limbi-straine")]
        public async Task<IActionResult> ExportLimbi(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] string? cicluStudii)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Semestru", typeof(int)),
                new DataColumn("Total Ore", typeof(decimal)),
                new DataColumn("Limba Program"), new DataColumn("Program Studiu"),
                new DataColumn("Ciclu Studii")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlLimbi, conn); cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@ciclu", string.IsNullOrWhiteSpace(cicluStudii) ? "Toti" : cicluStudii.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["Profesor"]?.ToString(),
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["TotalOreConv"] == DBNull.Value ? 0m : Convert.ToDecimal(r["TotalOreConv"]),
                    r["LimbaProgram"]?.ToString(), r["ProgramStudiu"]?.ToString(),
                    r["CicluStudii"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Limbi Straine");
            ws.Cell(1, 1).Value = "Raport Limbi Straine";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Total Ore").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                SafeFile("Limbi_Straine", profesor));
        }

        
        // RAPORT 5
        
        private const string SqlDisc = @"
            WITH Identitate AS (
                SELECT ID_Profesor, NumeIntreg, DenumireFacultate, DenumireCatedra, DenumireGradDidactic,
                       ID_Facultate, ID_Catedra,
                       CAST(ID_Profesor AS VARCHAR) + '_' + NumeIntreg AS IdentificatorUnic
                FROM [dbo].[View_Profesori_CF_AnUniv]
                WHERE ID_AnUnivCatedra = @idAn
            ),
            PpmDedup AS (
                SELECT ID_Profesor,
                       ISNULL(Denumire, '') AS MaterieBruta,
                       NrSemestruDinAn,
                       -- Aplicam aceeasi normalizare specializare ca la R3:
                       -- scoatem partea dupa '+', UPPER, fara diacritice si paranteze
                       LTRIM(RTRIM(
                           REPLACE(REPLACE(REPLACE(REPLACE(
                               UPPER(
                                   CASE WHEN CHARINDEX(N'+', DenumireSpecializare) > 0
                                        THEN LEFT(DenumireSpecializare, CHARINDEX(N'+', DenumireSpecializare)-1)
                                        ELSE DenumireSpecializare END
                               ) COLLATE Latin1_General_CI_AI,
                               N'(', N' '), N')', N' '), N'  ', N' '), N'  ', N' ')
                       )) AS SpecKey
                FROM [pi].[View_PostProfesorMaterie]
                WHERE ID_AnUniv = @idAn
                  AND (Post_Profesor_Materie_Deleted IS NULL OR Post_Profesor_Materie_Deleted = 0)
            ),
            Distinct_Mat AS (
                SELECT DISTINCT i.IdentificatorUnic, i.NumeIntreg,
                    i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic,
                    i.ID_Facultate, i.ID_Catedra,
                    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.MaterieBruta, N'  ', N'><'), N'<>', N''), N'><', N' '))) AS DenumireMaterie
                FROM PpmDedup ppm
                INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                WHERE (@prof = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                  AND (@sem  = 0       OR ppm.NrSemestruDinAn = @sem)
                  AND (@idFac     = 0 OR i.ID_Facultate = @idFac)
                  AND (@idCatedra = 0 OR i.ID_Catedra   = @idCatedra)
                  -- Filtru specializare: aplicam aceeasi normalizare pe valoarea
                  -- primita de la frontend pentru a fi consistent cu SpecKey
                  AND (@spec = N'Toti' OR ppm.SpecKey =
                       LTRIM(RTRIM(
                           REPLACE(REPLACE(REPLACE(REPLACE(
                               UPPER(@spec) COLLATE Latin1_General_CI_AI,
                               N'(', N' '), N')', N' '), N'  ', N' '), N'  ', N' ')
                       )))
                  AND ppm.MaterieBruta != N''
            )
            SELECT dm.NumeIntreg, MAX(dm.DenumireFacultate) AS Facultate,
                MAX(dm.DenumireCatedra) AS Departament, MAX(dm.DenumireGradDidactic) AS Grad,
                STUFF((SELECT DISTINCT N' | ' + dm2.DenumireMaterie FROM Distinct_Mat dm2
                       WHERE dm2.IdentificatorUnic = dm.IdentificatorUnic
                         AND dm2.DenumireMaterie IS NOT NULL
                       FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 3, N'') AS Discipline,
                COUNT(DISTINCT dm.DenumireMaterie) AS NrDisc
            FROM Distinct_Mat dm
            GROUP BY dm.IdentificatorUnic, dm.NumeIntreg
            ORDER BY dm.NumeIntreg COLLATE Romanian_CI_AS";

        [HttpGet("discipline")]
        public async Task<IActionResult> GetDisc(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] int? semestru, [FromQuery] string? specializare)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var result = new List<object>();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDisc, conn); cmd.CommandTimeout = TimeoutLong;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
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

        [HttpGet("export/discipline")]
        public async Task<IActionResult> ExportDisc(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] int? semestru, [FromQuery] string? specializare)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad"),
                new DataColumn("Discipline"), new DataColumn("Nr. Discipline", typeof(int))
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlDisc, conn); cmd.CommandTimeout = TimeoutLong;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(profesor) ? "Toti" : profesor.Trim());
            cmd.Parameters.AddWithValue("@sem", semestru ?? 0);
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(specializare) ? "Toti" : specializare.Trim());
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++, r["NumeIntreg"]?.ToString(), r["Facultate"]?.ToString(),
                    r["Departament"]?.ToString(), r["Grad"]?.ToString(),
                    r["Discipline"]?.ToString(),
                    r["NrDisc"] == DBNull.Value ? 0 : Convert.ToInt32(r["NrDisc"]));
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Discipline");
            ws.Cell(1, 1).Value = "Discipline per Profesor";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Nr. Discipline").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                SafeFile("Discipline", profesor));
        }

       
        [HttpGet("export/discipline-zip")]
        public async Task<IActionResult> ExportDiscZip(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra,
            [FromQuery] string? profesor, [FromQuery] int? semestru, [FromQuery] string? specializare)
        {
            int idAn = idAnUniv ?? GetAnCurent();

           
            var forme = new[]
            {
                ("IF",  new[] { "CU FRECVENȚĂ", "CU FRECVENTA" }),
                ("ID",  new[] { "ÎNVĂȚĂMÂNT LA DISTANȚĂ", "INVATAMANT LA DISTANTA",
                                "LA DISTANTA", "LA DISTANȚĂ" }),
                ("IFR", new[] { "FRECVENȚĂ REDUSĂ", "FRECVENTA REDUSA" })
            };

            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var (cod, denumiri) in forme)
                {
                    var dt = await BuildDiscDataTable(idAn, idFacultate ?? 0, idCatedra ?? 0,
                        profesor, semestru ?? 0, specializare, denumiri);
                    if (dt.Rows.Count == 0) continue;

                    using var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add($"Discipline {cod}");
                    ws.Cell(1, 1).Value = $"Discipline Predate - {cod}";
                    ws.Cell(1, 1).Style.Font.Bold = true;
                    ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
                    ws.Range(1, 1, 1, dt.Columns.Count).Merge();

                    var tbl = ws.Cell(3, 1).InsertTable(dt);
                    tbl.Theme = XLTableTheme.None;
                    StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
                    ws.Columns().AdjustToContents();

                    using var fileMs = new MemoryStream();
                    wb.SaveAs(fileMs);
                    var entry = zip.CreateEntry($"Discipline_{cod}.xlsx");
                    using var es = entry.Open();
                    fileMs.Position = 0;
                    fileMs.CopyTo(es);
                }
            }
            ms.Position = 0;
            return File(ms.ToArray(), "application/zip",
                SafeFile("Discipline_IF_ID_IFR", profesor) + ".zip");
        }

       
        private async Task<DataTable> BuildDiscDataTable(int idAn, int idFac, int idCat,
            string? prof, int sem, string? spec, string[] formeAcceptate)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[]{
                new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad"),
                new DataColumn("Discipline"), new DataColumn("Nr. Discipline", typeof(int))
            });

            
            var paramsForma = new List<string>();
            for (int i = 0; i < formeAcceptate.Length; i++)
                paramsForma.Add($"@f{i}");
            string inList = string.Join(", ", paramsForma);

            string sqlFiltrat = $@"
                WITH PpmFiltrat AS (
                    SELECT vppm.ID_Profesor, vppm.Denumire AS MaterieBruta, vppm.NrSemestruDinAn,
                           LTRIM(RTRIM(
                               REPLACE(REPLACE(REPLACE(REPLACE(
                                   UPPER(
                                       CASE WHEN CHARINDEX(N'+', vppm.DenumireSpecializare) > 0
                                            THEN LEFT(vppm.DenumireSpecializare, CHARINDEX(N'+', vppm.DenumireSpecializare)-1)
                                            ELSE vppm.DenumireSpecializare END
                                   ) COLLATE Latin1_General_CI_AI,
                                   N'(', N' '), N')', N' '), N'  ', N' '), N'  ', N' ')
                           )) AS SpecKey,
                           vppm.ID_Facultate, vppm.ID_Catedra
                    FROM [pi].[View_PostProfesorMaterie] vppm
                    INNER JOIN [dbo].[View_FDS] f
                        ON f.ID_Specializare = vppm.id_specializare AND f.ID_AnUniv = vppm.ID_AnUniv
                    WHERE vppm.ID_AnUniv = @idAn
                      AND f.DenumireFormaInv COLLATE Latin1_General_CI_AI IN ({inList})
                      AND (vppm.Post_Profesor_Materie_Deleted IS NULL OR vppm.Post_Profesor_Materie_Deleted = 0)
                ),
                Identitate AS (
                    SELECT ID_Profesor, NumeIntreg, DenumireFacultate, DenumireCatedra, DenumireGradDidactic,
                           ID_Facultate, ID_Catedra,
                           CAST(ID_Profesor AS VARCHAR) + '_' + NumeIntreg AS IdentificatorUnic
                    FROM [dbo].[View_Profesori_CF_AnUniv]
                    WHERE ID_AnUnivCatedra = @idAn
                ),
                Distinct_Mat AS (
                    SELECT DISTINCT i.IdentificatorUnic, i.NumeIntreg,
                        i.DenumireFacultate, i.DenumireCatedra, i.DenumireGradDidactic,
                        LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.MaterieBruta, N'  ', N'><'), N'<>', N''), N'><', N' '))) AS DenumireMaterie
                    FROM PpmFiltrat ppm
                    INNER JOIN Identitate i ON i.ID_Profesor = ppm.ID_Profesor
                    WHERE (@prof = N'Toti' OR i.NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT)
                      AND (@sem  = 0       OR ppm.NrSemestruDinAn = @sem)
                      AND (@idFac     = 0 OR i.ID_Facultate = @idFac)
                      AND (@idCatedra = 0 OR i.ID_Catedra   = @idCatedra)
                      AND (@spec = N'Toti' OR ppm.SpecKey =
                           LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(
                               UPPER(@spec) COLLATE Latin1_General_CI_AI,
                               N'(', N' '), N')', N' '), N'  ', N' '), N'  ', N' '))))
                      AND ppm.MaterieBruta != N''
                )
                SELECT dm.NumeIntreg, MAX(dm.DenumireFacultate) AS Facultate,
                    MAX(dm.DenumireCatedra) AS Departament, MAX(dm.DenumireGradDidactic) AS Grad,
                    STUFF((SELECT DISTINCT N' | ' + dm2.DenumireMaterie FROM Distinct_Mat dm2
                           WHERE dm2.IdentificatorUnic = dm.IdentificatorUnic
                             AND dm2.DenumireMaterie IS NOT NULL
                           FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 3, N'') AS Discipline,
                    COUNT(DISTINCT dm.DenumireMaterie) AS NrDisc
                FROM Distinct_Mat dm
                GROUP BY dm.IdentificatorUnic, dm.NumeIntreg
                ORDER BY dm.NumeIntreg COLLATE Romanian_CI_AS";

            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(sqlFiltrat, conn);
            cmd.CommandTimeout = TimeoutLong;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFac);
            cmd.Parameters.AddWithValue("@idCatedra", idCat);
            cmd.Parameters.AddWithValue("@prof", string.IsNullOrWhiteSpace(prof) ? "Toti" : prof.Trim());
            cmd.Parameters.AddWithValue("@sem", sem);
            cmd.Parameters.AddWithValue("@spec", string.IsNullOrWhiteSpace(spec) ? "Toti" : spec.Trim());
            for (int i = 0; i < formeAcceptate.Length; i++)
                cmd.Parameters.AddWithValue($"@f{i}", formeAcceptate[i]);

            using var r = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await r.ReadAsync())
            {
                dt.Rows.Add(nr++,
                    r["NumeIntreg"]?.ToString() ?? "",
                    r["Facultate"]?.ToString() ?? "",
                    r["Departament"]?.ToString() ?? "",
                    r["Grad"]?.ToString() ?? "",
                    r["Discipline"]?.ToString() ?? "",
                    r["NrDisc"] == DBNull.Value ? 0 : Convert.ToInt32(r["NrDisc"]));
            }
            return dt;
        }

        
        // RAPORT 6
      
        private const string SqlTitulari = @"
            SELECT DISTINCT
                V.ID_Profesor,
                V.Nume,
                V.Prenume,
                V.Marca,
                V.NumeIntreg,
                V.DenumireGradDidactic,
                V.DenumireCatedra   AS Departament,
                V.DenumireFacultate AS Facultate
            FROM [dbo].[View_Profesori_CF_AnUniv] AS V
            WHERE V.ID_AnUnivCatedra = @idAn
              AND V.TitularAnUniv = 1
              AND (@idFac     = 0 OR V.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR V.ID_Catedra   = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [pi].[Post] AS P
                  INNER JOIN [pi].[Post_Profesor] AS PP ON P.ID_Post = PP.ID_Post
                  INNER JOIN [pi].[StatDeFunctii] AS SF ON P.ID_StatDeFunctii = SF.ID_StatDeFunctii
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
            using var cmd = new SqlCommand(SqlTitulari, conn); cmd.CommandTimeout = 60;
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
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulari, conn); cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["Nume"]?.ToString(), r["Prenume"]?.ToString(), r["Marca"]?.ToString(),
                    r["Facultate"]?.ToString(), r["Departament"]?.ToString(),
                    r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Titulari");
            ws.Cell(1, 1).Value = $"Cadre Didactice Titulare | An ID: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true; tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Titulari.xlsx");
        }

        
        // RAPORT 7
     
        private const string SqlColaboratori = @"
            SELECT DISTINCT
                V.ID_Profesor,
                V.Nume,
                V.Prenume,
                V.Marca,
                V.NumeIntreg,
                V.DenumireGradDidactic,
                V.DenumireCatedra   AS Departament,
                V.DenumireFacultate AS Facultate
            FROM [dbo].[View_Profesori_CF_AnUniv] AS V
            WHERE V.ID_AnUnivCatedra = @idAn
              AND V.TitularAnUniv = 0
              AND (@idFac     = 0 OR V.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR V.ID_Catedra   = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [pi].[Post] AS P
                  INNER JOIN [pi].[Post_Profesor] AS PP ON P.ID_Post = PP.ID_Post
                  INNER JOIN [pi].[StatDeFunctii] AS SF ON P.ID_StatDeFunctii = SF.ID_StatDeFunctii
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
            using var cmd = new SqlCommand(SqlColaboratori, conn); cmd.CommandTimeout = 60;
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
                new DataColumn("Facultate"), new DataColumn("Departament"), new DataColumn("Grad")
            });
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColaboratori, conn); cmd.CommandTimeout = 60;
            cmd.Parameters.AddWithValue("@idAn", idAn);
            cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
            cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
            using var r = await cmd.ExecuteReaderAsync(); int nr = 1;
            while (await r.ReadAsync())
                dt.Rows.Add(nr++,
                    r["Nume"]?.ToString(), r["Prenume"]?.ToString(), r["Marca"]?.ToString(),
                    r["Facultate"]?.ToString(), r["Departament"]?.ToString(),
                    r["DenumireGradDidactic"]?.ToString());
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Colaboratori");
            ws.Cell(1, 1).Value = $"Asociati / Colaboratori | An ID: {idAn}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(Green);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true; tbl.Field("Nr.Crt.").TotalsRowLabel = "TOTAL";
            StyleHdr(ws.Range(3, 1, 3, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var s = new MemoryStream(); wb.SaveAs(s);
            return File(s.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Colaboratori.xlsx");
        }


        // RAPORT 8
       
        private const string SqlAnsMaster = @"
            -- Pas 1: snapshot StatDeFunctiiPeSpecializare cu maparea ramura ANS
            IF OBJECT_ID('tempdb..#DM') IS NOT NULL DROP TABLE #DM;
            CREATE TABLE #DM (
                ID_ELEMENT INT,
                COD_DS_CNATDCU NVARCHAR(20),
                cod_DS NVARCHAR(20),
                ID_RamuraDeStiinta_ANS INT,
                DomeniulDeStudiu_ANS NVARCHAR(200),
                RamuraDeStiinta_ANS NVARCHAR(200),
                DomeniuFundamental NVARCHAR(200)
            );
            INSERT INTO #DM EXEC [dbo].[N_DOMENIU_STUDIUL_ANS_List];

            -- Pas 2: SFS curatat de cuplaje defecte si imbogatit cu ramura ANS
            IF OBJECT_ID('tempdb..#sfs') IS NOT NULL DROP TABLE #sfs;
            SELECT
                sfs.ID_Profesor,
                sfs.ID_TipGradDidactic,
                sfs.NrCrtPost,
                sfs.TitularSauSuplinitor,
                sfs.DenTitularSauSuplinitor,
                sfs.NrOreConventionale,
                sfs.id_specializare,
                sfs.ID_Facultate,
                CASE WHEN sfs.xTipCuplaj IN (
                        N'CuplajeCareNuMaiExista',
                        N'AplicDinCuplajCurs',
                        N'AplicDinCuplajApp')
                     THEN NULL
                     ELSE sfs.ApartineDeCuplaj
                END AS ApartineDeCuplaj,
                fds.ID_N_Domeniu_Studiu_ANS,
                dm.ID_RamuraDeStiinta_ANS,
                dm.RamuraDeStiinta_ANS
            INTO #sfs
            FROM [pi].[StatDeFunctiiPeSpecializare] sfs
            INNER JOIN (
                -- Maparea specializare → domeniu ANS este stabila in timp
                -- (e nomenclator national). Specializari continuate din ani vechi
                -- (ex. Autovehicule rutiere ID_AnUniv=43) NU apar in View_FDS pentru
                -- anul curent dar APAR in StatDeFunctii. Luam maparea din orice an
                -- in care exista, nu doar din anul curent, ca sa nu pierdem ore.
                SELECT ID_Specializare,
                       MAX(ID_N_Domeniu_Studiu_ANS) AS ID_N_Domeniu_Studiu_ANS
                FROM [dbo].[View_FDS]
                WHERE ID_N_Domeniu_Studiu_ANS IS NOT NULL
                GROUP BY ID_Specializare
            ) fds
                ON sfs.id_specializare = fds.ID_Specializare
            INNER JOIN #DM dm
                ON dm.ID_ELEMENT = fds.ID_N_Domeniu_Studiu_ANS
            WHERE sfs.ID_AnUniv = @idAn
              AND sfs.TitularSauSuplinitor = 1
              AND sfs.DenTitularSauSuplinitor <> N'SupTit'
              AND (@idFac     = 0 OR sfs.ID_Facultate = @idFac);

            -- Pas 3: cnt pentru deduplicarea cuplajelor (formula profei)
            IF OBJECT_ID('tempdb..#sfs_cnt') IS NOT NULL DROP TABLE #sfs_cnt;
            SELECT
                s.*,
                CASE WHEN ISNULL(s.ApartineDeCuplaj, -1) = -1 THEN 1
                     ELSE COUNT(*) OVER (
                            PARTITION BY s.ID_Profesor, s.ID_TipGradDidactic,
                                         ISNULL(s.ApartineDeCuplaj, -1), s.NrCrtPost)
                END AS cnt
            INTO #sfs_cnt
            FROM #sfs s;

            -- Pas 4: fractia per profesor x ramura, cu norma individuala
            -- (ExceptiiNormaOreConventionale > NormaOreConventionale standard)
            -- Folosim denumirea OFICIALA din N_RAMURA_STIINTA_ANS (cu paranteze
            -- complete: 'Stiinte economice (fara Cibernetica...)' nu varianta scurta
            -- 'Stiinte economice' din N_DOMENIU_STUDIUL_ANS_List).
            SELECT
                v.ID_Profesor,
                v.NumeIntreg                            AS NumeIntreg,
                v.DenumireFacultate                     AS Facultate,
                v.DenumireCatedra                       AS Departament,
                v.DenumireGradDidactic                  AS Grad,
                s.ID_RamuraDeStiinta_ANS                AS IdRamura,
                ISNULL(rsa.Denumire, s.RamuraDeStiinta_ANS) AS RamuraNume,
                CAST(SUM(
                    s.NrOreConventionale / s.cnt /
                    ISNULL(enoc.NrOreTitular, ISNULL(noc.NrOreConventionaleTitular, @normaFallback))
                ) AS DECIMAL(10,4))                     AS Fractie
            FROM #sfs_cnt s
            INNER JOIN [dbo].[View_Profesori_CF_AnUniv] v
                ON s.ID_Profesor = v.ID_Profesor
                AND v.ID_AnUnivCatedra = @idAn
                AND v.TitularAnUniv = 1
            LEFT JOIN [dbo].[N_RAMURA_STIINTA_ANS] rsa
                ON rsa.ID_Element = s.ID_RamuraDeStiinta_ANS
            LEFT JOIN [pi].[NormaOreConventionale] noc
                ON noc.ID_TipGradDidactic = s.ID_TipGradDidactic
                AND noc.ID_AnUniv = @idAn
            LEFT JOIN [pi].[ExceptiiNormaOreConventionale] enoc
                ON enoc.ID_Profesor = s.ID_Profesor
                AND enoc.ID_AnUniv = @idAn
            WHERE (@idFac     = 0 OR v.ID_Facultate = @idFac)
              AND (@idCatedra = 0 OR v.ID_Catedra   = @idCatedra)
              AND EXISTS (
                  SELECT 1
                  FROM [pi].[Post] AS P
                  INNER JOIN [pi].[Post_Profesor] AS PP ON P.ID_Post = PP.ID_Post
                  INNER JOIN [pi].[StatDeFunctii] AS SF ON P.ID_StatDeFunctii = SF.ID_StatDeFunctii
                  WHERE PP.ID_Profesor = v.ID_Profesor
                    AND SF.ID_AnUniv   = @idAn
                    AND P.TitularSauSuplinitor = 1
                    AND P.Deleted = 0
                    AND PP.Deleted = 0
              )
            GROUP BY v.ID_Profesor, v.NumeIntreg,
                     v.DenumireFacultate, v.DenumireCatedra, v.DenumireGradDidactic,
                     s.ID_RamuraDeStiinta_ANS, s.RamuraDeStiinta_ANS, rsa.Denumire
            HAVING SUM(s.NrOreConventionale / s.cnt /
                       ISNULL(enoc.NrOreTitular, ISNULL(noc.NrOreConventionaleTitular, @normaFallback))) > 0
            ORDER BY v.NumeIntreg COLLATE Romanian_CI_AS, s.ID_RamuraDeStiinta_ANS;";

       
        private List<(int Id, string Den)> GetRamuriAns(SqlConnection conn)
        {
            return _cache.GetOrCreate("RamuriANS_v4", e =>
            {
                e.AbsoluteExpirationRelativeToNow = TimeSpan.FromDays(1);

                
                var rsa = new Dictionary<int, string>();
                using (var cmd = new SqlCommand(
                    "SELECT ID_Element, Denumire FROM [dbo].[N_RAMURA_STIINTA_ANS]", conn))
                {
                    cmd.CommandTimeout = TimeoutShort;
                    using var r = cmd.ExecuteReader();
                    while (r.Read())
                        rsa[Convert.ToInt32(r[0])] = r[1]?.ToString() ?? "";
                }

                
                using var cmd2 = new SqlCommand(@"
                    IF OBJECT_ID('tempdb..#DM_ramuri') IS NOT NULL DROP TABLE #DM_ramuri;
                    CREATE TABLE #DM_ramuri (
                        ID_ELEMENT INT, COD_DS_CNATDCU NVARCHAR(20), cod_DS NVARCHAR(20),
                        ID_RamuraDeStiinta_ANS INT, DomeniulDeStudiu_ANS NVARCHAR(200),
                        RamuraDeStiinta_ANS NVARCHAR(200), DomeniuFundamental NVARCHAR(200)
                    );
                    INSERT INTO #DM_ramuri EXEC [dbo].[N_DOMENIU_STUDIUL_ANS_List];
                    SELECT DISTINCT ID_RamuraDeStiinta_ANS, RamuraDeStiinta_ANS
                    FROM #DM_ramuri
                    WHERE ID_RamuraDeStiinta_ANS IS NOT NULL
                    ORDER BY ID_RamuraDeStiinta_ANS;", conn);
                cmd2.CommandTimeout = TimeoutShort;
                var lst = new List<(int, string)>();
                using var rd = cmd2.ExecuteReader();
                while (rd.Read())
                {
                    int id = Convert.ToInt32(rd[0]);
                  
                    string den = rsa.TryGetValue(id, out var d) ? d : (rd[1]?.ToString() ?? "");
                   
                    if (den.Contains("Stiinta Sportului", StringComparison.OrdinalIgnoreCase) ||
                        den.Contains("Ştiinţa Sportului", StringComparison.OrdinalIgnoreCase))
                    {
                        den = "Stiintele Sportului si Educatiei Fizice";
                    }
                    lst.Add((id, den));
                }
                return lst;
            })!;
        }

        [HttpGet("raport-ans")]
        public async Task<IActionResult> GetAns(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();

            var ramuri = GetRamuriAns(conn);

           
            var profMap = new Dictionary<int, (string Nume, string Fac, string Dept, string Grad,
                Dictionary<int, decimal> Frac)>();

            using (var cmd = new SqlCommand(SqlAnsMaster, conn))
            {
                cmd.CommandTimeout = TimeoutLong;
                cmd.Parameters.AddWithValue("@idAn", idAn);
                cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
                cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
                cmd.Parameters.AddWithValue("@normaFallback", NormaLegalaFallback);
                using var rr = await cmd.ExecuteReaderAsync();
                while (await rr.ReadAsync())
                {
                    int idP = Convert.ToInt32(rr["ID_Profesor"]);
                    int idR = Convert.ToInt32(rr["IdRamura"]);
                    decimal fr = rr["Fractie"] == DBNull.Value ? 0m : Convert.ToDecimal(rr["Fractie"]);
                    if (fr <= 0) continue;
                    if (!profMap.TryGetValue(idP, out var entry))
                    {
                        entry = (
                            rr["NumeIntreg"]?.ToString() ?? "",
                            rr["Facultate"]?.ToString() ?? "",
                            rr["Departament"]?.ToString() ?? "",
                            rr["Grad"]?.ToString() ?? "",
                            new Dictionary<int, decimal>()
                        );
                        profMap[idP] = entry;
                    }
                    entry.Frac[idR] = Math.Round(fr, 2);
                }
            }

            int nrCrt = 1;
            var profesori = profMap
                .OrderBy(kv => kv.Value.Nume,
                    StringComparer.Create(new System.Globalization.CultureInfo("ro-RO"), true))
                .Select(kv =>
                {
                    var domeniiMapate = new Dictionary<string, decimal>();
                    foreach (var (id, den) in ramuri)
                        if (kv.Value.Frac.TryGetValue(id, out var f) && f > 0)
                            domeniiMapate[den] = f;
                    return (object)new
                    {
                        NrCrt = nrCrt++,
                        NumeComplet = kv.Value.Nume,
                        Facultate = kv.Value.Fac,
                        Departament = kv.Value.Dept,
                        GradFunctie = kv.Value.Grad,
                        DomeniiMapate = domeniiMapate
                    };
                })
                .ToList();

            return Ok(new
            {
                Domenii = ramuri.Select(r => r.Den).ToList(),
                Profesori = profesori
            });
        }

        [HttpGet("export/raport-ans")]
        public async Task<IActionResult> ExportAns(
            [FromQuery] int? idAnUniv, [FromQuery] int? idFacultate, [FromQuery] int? idCatedra)
        {
            int idAn = idAnUniv ?? GetAnCurent();
            using var conn = new SqlConnection(_cs); await conn.OpenAsync();

            var ramuri = GetRamuriAns(conn);

            var profMap = new Dictionary<int, (string Nume, string Fac, string Dept, string Grad,
                Dictionary<int, decimal> Frac)>();
            using (var cmd = new SqlCommand(SqlAnsMaster, conn))
            {
                cmd.CommandTimeout = TimeoutLong;
                cmd.Parameters.AddWithValue("@idAn", idAn);
                cmd.Parameters.AddWithValue("@idFac", idFacultate ?? 0);
                cmd.Parameters.AddWithValue("@idCatedra", idCatedra ?? 0);
                cmd.Parameters.AddWithValue("@normaFallback", NormaLegalaFallback);
                using var rr = await cmd.ExecuteReaderAsync();
                while (await rr.ReadAsync())
                {
                    int idP = Convert.ToInt32(rr["ID_Profesor"]);
                    int idR = Convert.ToInt32(rr["IdRamura"]);
                    decimal fr = rr["Fractie"] == DBNull.Value ? 0m : Convert.ToDecimal(rr["Fractie"]);
                    if (fr <= 0) continue;
                    if (!profMap.TryGetValue(idP, out var entry))
                    {
                        entry = (
                            rr["NumeIntreg"]?.ToString() ?? "",
                            rr["Facultate"]?.ToString() ?? "",
                            rr["Departament"]?.ToString() ?? "",
                            rr["Grad"]?.ToString() ?? "",
                            new Dictionary<int, decimal>()
                        );
                        profMap[idP] = entry;
                    }
                    entry.Frac[idR] = Math.Round(fr, 2);
                }
            }

            int nrR = ramuri.Count, colTot = 6 + nrR + 1;
            var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("Raport ANS");
            var gC = XLColor.FromHtml(Green);

            ws.Cell(1, 1).Value = "Anexa 1. Tabel institutional - Raport ANS";
            ws.Range(1, 1, 1, colTot).Merge();
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = gC;
            ws.Cell(2, 1).Value = $"An universitar (ID): {idAn}";
            ws.Range(2, 1, 2, colTot).Merge();
            ws.Cell(2, 1).Style.Font.Italic = true;

            string[] antet = { "Nr.\nCrt.", "Nume si prenume", "Functie",
                               "Facultate", "Departament", "Norma" };
            for (int c = 1; c <= 6; c++)
            {
                ws.Cell(4, c).Value = antet[c - 1];
                ws.Cell(4, c).Style.Font.Bold = true;
                ws.Cell(4, c).Style.Alignment.WrapText = true;
                ws.Cell(4, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            for (int i = 0; i < nrR; i++)
            {
                ws.Cell(4, 7 + i).Value = ramuri[i].Den;
                ws.Cell(4, 7 + i).Style.Font.Bold = true;
                ws.Cell(4, 7 + i).Style.Alignment.WrapText = true;
                ws.Cell(4, 7 + i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(4, 7 + i).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            ws.Cell(4, colTot).Value = "Total";
            ws.Cell(4, colTot).Style.Font.Bold = true;
            ws.Cell(4, colTot).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(4, 1, 4, colTot).Style.Fill.BackgroundColor = gC;
            ws.Range(4, 1, 4, colTot).Style.Font.FontColor = XLColor.White;

            ws.Column(1).Width = 5; ws.Column(2).Width = 30;
            ws.Column(3).Width = 22; ws.Column(4).Width = 28; ws.Column(5).Width = 28;
            ws.Column(6).Width = 10;
            for (int c = 7; c <= colTot; c++) ws.Column(c).Width = 12;

            int row = 5, nr = 1;
            foreach (var kv in profMap.OrderBy(p => p.Value.Nume,
                StringComparer.Create(new System.Globalization.CultureInfo("ro-RO"), true)))
            {
                ws.Cell(row, 1).Value = nr++;
                ws.Cell(row, 2).Value = kv.Value.Nume;
                ws.Cell(row, 3).Value = kv.Value.Grad;
                ws.Cell(row, 4).Value = kv.Value.Fac;
                ws.Cell(row, 5).Value = kv.Value.Dept;
                decimal totFrac = 0m;
                for (int i = 0; i < nrR; i++)
                {
                    if (kv.Value.Frac.TryGetValue(ramuri[i].Id, out decimal f) && f > 0)
                    {
                        ws.Cell(row, 7 + i).Value = (double)f;
                        ws.Cell(row, 7 + i).Style.NumberFormat.Format = "0.00";
                        totFrac += f;
                    }
                }
                ws.Cell(row, 6).Value = (double)Math.Round(totFrac, 2);
                ws.Cell(row, 6).Style.NumberFormat.Format = "0.00";
                ws.Cell(row, colTot).Value = (double)Math.Round(totFrac, 2);
                ws.Cell(row, colTot).Style.NumberFormat.Format = "0.00";
                for (int c = 1; c <= colTot; c++)
                    ws.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                row++;
            }

            using var stream = new MemoryStream(); wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Raport_ANS.xlsx");
        }
    }
}