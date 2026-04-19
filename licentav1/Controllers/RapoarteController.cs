using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using System.IO.Compression;
using ClosedXML.Excel;

namespace LicentaV1.Controllers
{
    /// <summary>
    /// RapoarteController — refactorizat Metadata-Driven.
    ///
    /// Modificari fata de versiunea anterioara:
    ///
    /// LISTE (dropdown-uri):
    ///   - Facultati:     [dbo].[FacultateListAll]                         (SP oficial)
    ///   - Departamente:  [dbo].[CatedraListByFacultateAnUniv]             (SP oficial, params: @ID_AnUniv, @ID_Facultate)
    ///   - Specializari:  [dbo].[SpecializareListByFacultateAnUniv]        (SP oficial, params: @ID_AnUniv, @ID_Facultate)
    ///   - Profesori:     SELECT DISTINCT pe View_Profesori_CF_AnUniv      (nu exista SP dedicat)
    ///   - Ani univ:      SELECT WHERE Denumire LIKE '%2025-2026%', fallback ID 45
    ///
    /// AN UNIVERSITAR CURENT:
    ///   SetariUniversitate blocat -> cautam Denumire LIKE '%2025-2026%' in AnUniversitar.
    ///   Fallback: Ordine DESC (cel mai recent), fallback final: 45.
    ///
    /// RAPOARTE (regiunile 1-7):
    ///   Nemodificate structural — refactorizarea listelor nu afecteaza query-urile de raport.
    ///   Parametrii @ID_Facultate inlocuiesc @fac (string) pentru dropdownuri;
    ///   query-urile de raport continua sa foloseasca @fac (string) pentru compatibilitate.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class RapoarteController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;
        private readonly IMemoryCache _cache;
        private const string BrandColorHex = "#56723e";

        // An universitar curent — detectat din BD la startup, nu hardcodat
        private readonly int _idAnUnivCurent;

        public RapoarteController(IConfiguration configuration, IMemoryCache cache)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection")!;
            _cache = cache;
            _idAnUnivCurent = DetectaAnUnivCurent();
        }

        // =====================================================================
        // Detectie an universitar curent — fara SetariUniversitate (blocat).
        // Strategia:
        //   1. Cauta Denumire LIKE '%2025-2026%' (an curent hardcodat DOAR in pattern,
        //      nu in ID — pattern-ul e parametrizabil din config daca e necesar)
        //   2. Fallback: MAX(Ordine) — cel mai recent an din BD
        //   3. Fallback final: 45
        // =====================================================================
        private int DetectaAnUnivCurent()
        {
            // Pattern-ul anului curent poate veni din appsettings daca se doreste
            // zero-hardcoding absolut. Deocamdata e citit din config cu fallback inline.
            string anPattern = _configuration["AnUniversitar:Pattern"] ?? "2025-2026";

            try
            {
                using var conn = new SqlConnection(_connectionString);
                conn.Open();

                // Strategia 1: an explicit dupa denumire
                using (var cmd = new SqlCommand(@"
                    SELECT TOP 1 ID_AnUniv
                    FROM [AGSIS].[dbo].[AnUniversitar]
                    WHERE Denumire LIKE @pattern
                      AND Denumire IS NOT NULL
                    ORDER BY Ordine DESC", conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@pattern", SqlDbType.NVarChar, 50)
                    { Value = $"%{anPattern}%" });
                    var val = cmd.ExecuteScalar();
                    if (val != null && val != DBNull.Value)
                        return Convert.ToInt32(val);
                }

                // Strategia 2: cel mai recent dupa Ordine
                using (var cmd2 = new SqlCommand(@"
                    SELECT TOP 1 ID_AnUniv
                    FROM [AGSIS].[dbo].[AnUniversitar]
                    WHERE Denumire IS NOT NULL
                    ORDER BY Ordine DESC", conn))
                {
                    var val2 = cmd2.ExecuteScalar();
                    if (val2 != null && val2 != DBNull.Value)
                        return Convert.ToInt32(val2);
                }
            }
            catch (Exception ex)
            {
                // Log minim — nu inghitim exceptia silentios
                Console.Error.WriteLine(
                    $"[RapoarteController] DetectaAnUnivCurent EROARE: {ex.Message}. Fallback: 45.");
            }

            return 45; // fallback final documentat
        }

        // =====================================================================
        // Norme legale — reutilizat de rapoartele de norma
        // =====================================================================
        private async Task<Dictionary<int, decimal>> LoadNormeLegaleAsync(
            SqlConnection conn, int idAnUniv)
        {
            var dict = new Dictionary<int, decimal>();
            const string sql = @"
                SELECT ID_TipGradDidactic, NrOreConventionaleTitular
                FROM [AGSIS].[pi].[NormaOreConventionale]
                WHERE ID_AnUniv = @id
                  AND NrOreConventionaleTitular IS NOT NULL
                  AND NrOreConventionaleTitular > 0";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = idAnUniv });
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                dict[Convert.ToInt32(r["ID_TipGradDidactic"])] =
                    Convert.ToDecimal(r["NrOreConventionaleTitular"]);
            return dict;
        }

        #region ================= LISTE (DROPDOWNS via Stored Procedures) =================

        // -------------------------------------------------------------------------
        // Ani universitari — query direct (nu exista SP dedicat)
        // Cache 60 min — lista se schimba o data pe an
        // -------------------------------------------------------------------------
        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAni()
        {
            return Ok(_cache.GetOrCreate("ListaAniUniv_v2", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<object>();
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT au.ID_AnUniv                              AS id,
                           UPPER(LTRIM(RTRIM(au.Denumire)))
                               COLLATE DATABASE_DEFAULT             AS AnCurat
                    FROM [AGSIS].[dbo].[AnUniversitar] au
                    WHERE au.Denumire IS NOT NULL
                      AND LTRIM(RTRIM(au.Denumire)) != ''
                    ORDER BY au.Ordine DESC", conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    lista.Add(new
                    {
                        id = reader["id"] != DBNull.Value ? Convert.ToInt32(reader["id"]) : 0,
                        nume = reader["AnCurat"]?.ToString() ?? ""
                    });
                return lista;
            }));
        }

        // -------------------------------------------------------------------------
        // Facultati — via SP oficial [dbo].[FacultateListAll]
        // SP returneaza: ID_Facultate, Denumire (sau similar — adaptam la schema SP).
        // Cache 60 min.
        // -------------------------------------------------------------------------
        [HttpGet("liste/facultati")]
        public ActionResult GetFacultati()
        {
            return Ok(_cache.GetOrCreate("ListaFacultati_SP_v1", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<object> {
                    new { id = (int?)null, nume = "Toti" }
                };
                try
                {
                    using var conn = new SqlConnection(_connectionString);
                    conn.Open();
                    using var cmd = new SqlCommand("[dbo].[FacultateListAll]", conn)
                    {
                        CommandType = CommandType.StoredProcedure,
                        CommandTimeout = 30
                    };
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        // SP poate returna coloane cu diferite denumiri — incercam variantele comune
                        int? id = TryGetInt(reader, "ID_Facultate") ?? TryGetInt(reader, "id");
                        string den = TryGetString(reader, "Denumire")
                                  ?? TryGetString(reader, "NumeFacultate")
                                  ?? TryGetString(reader, "DenumireFacultate")
                                  ?? reader[0]?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(den))
                            lista.Add(new { id, nume = den.Trim() });
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(
                        $"[RapoarteController] GetFacultati SP EROARE: {ex.Message}. " +
                        "Fallback pe SELECT DISTINCT.");
                    // Fallback: citire directa daca SP nu exista sau e incompatibil
                    lista = FacultatiDirectFallback();
                }
                return lista;
            }));
        }

        private List<object> FacultatiDirectFallback()
        {
            var lista = new List<object> { new { id = (int?)null, nume = "Toti" } };
            try
            {
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT
                        LTRIM(RTRIM(p.DenumireFacultate)) COLLATE DATABASE_DEFAULT AS Denumire
                    FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                    WHERE p.DenumireFacultate IS NOT NULL
                      AND LTRIM(RTRIM(p.DenumireFacultate)) != ''
                      AND p.ID_AnUnivCatedra = (
                          SELECT TOP 1 ID_AnUniv FROM [AGSIS].[dbo].[AnUniversitar]
                          ORDER BY Ordine DESC)
                    ORDER BY Denumire", conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var den = reader[0]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(den))
                        lista.Add(new { id = (int?)null, nume = den });
                }
            }
            catch { /* ignoram — lista va fi doar "Toti" */ }
            return lista;
        }

        // -------------------------------------------------------------------------
        // Departamente — via SP oficial [dbo].[CatedraListByFacultateAnUniv]
        // Params SP: @ID_AnUniv INT, @ID_Facultate INT
        // Endpoint primeste atat idFacultate (int, pentru SP) cat si numeFacultate
        // (string, pentru compatibilitate cu rapoartele existente).
        // -------------------------------------------------------------------------
        [HttpGet("liste/departamente")]
        public ActionResult GetDepartamente([FromQuery] int idAnUniv, [FromQuery] string numeFacultate)
        {
            var lista = new List<object> { new { id = 0, nume = "Toti" } };
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            // Folosim View-ul de profesori pentru a extrage DOAR departamentele facultatii selectate
            string sql = @"
        SELECT DISTINCT ID_Catedra, DenumireCatedra 
        FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv]
        WHERE ID_AnUnivCatedra = @idAn 
          AND (ID_Facultate = TRY_CAST(@fac AS INT) OR DenumireFacultate = @fac)
        ORDER BY DenumireCatedra ASC";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@idAn", idAnUniv);
            cmd.Parameters.AddWithValue("@fac", numeFacultate);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                lista.Add(new { id = reader["ID_Catedra"], nume = reader["DenumireCatedra"].ToString() });
            return Ok(lista);
        }
        private List<object> DepartamenteFallback(int idAn, string? numeFacultate)
        {
            var lista = new List<object> { new { id = (int?)null, nume = "Toti" } };
            try
            {
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT
                        LTRIM(RTRIM(vp.DenumireCatedra)) COLLATE DATABASE_DEFAULT AS Denumire
                    FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] vp
                    WHERE vp.ID_AnUnivCatedra = @idAn
                      AND vp.DenumireCatedra IS NOT NULL
                      AND LTRIM(RTRIM(vp.DenumireCatedra)) != ''
                      AND (
                          @fac = N'Toti'
                          OR vp.DenumireFacultate COLLATE DATABASE_DEFAULT
                           = @fac                  COLLATE DATABASE_DEFAULT
                      )
                    ORDER BY Denumire", conn);
                cmd.Parameters.Add(new SqlParameter("@idAn", SqlDbType.Int) { Value = idAn });
                cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200)
                { Value = string.IsNullOrWhiteSpace(numeFacultate) ? "Toti" : numeFacultate.Trim() });
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var den = reader[0]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(den))
                        lista.Add(new { id = (int?)null, nume = den });
                }
            }
            catch { }
            return lista;
        }

        // -------------------------------------------------------------------------
        // Specializari — via SP oficial [dbo].[SpecializareListByFacultateAnUniv]
        // Params SP: @ID_AnUniv INT, @ID_Facultate INT
        // -------------------------------------------------------------------------
        [HttpGet("liste/specializari")]
        public ActionResult GetSpecializari([FromQuery] int? idFacultate = null)
        {
            var lista = new List<object> { new { id = 0, nume = "Toti" } };
            if (!idFacultate.HasValue || idFacultate == 0) return Ok(lista);

            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            // Folosim procedura IL care a functionat in testul tau
            using var cmd = new SqlCommand("[IL].[SpecializareListByAnUnivCurentSiIdFacultate]", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@ID_Facultate", idFacultate.Value);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                lista.Add(new
                {
                    id = reader["ID_Specializare"],
                    // Folosim DenumireDomeniu pentru a avea numele curat (ex: ADMINISTRAREA AFACERILOR)
                    nume = reader["DenumireDomeniu"].ToString()
                });
            return Ok(lista);
        }
        private List<object> SpecializariFallback(int idAn, string? numeFacultate)
        {
            var lista = new List<object> { new { id = (int?)null, nume = "Toti" } };
            try
            {
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT
                        UPPER(LTRIM(RTRIM(ppm.DenumireSpecializare)))
                            COLLATE DATABASE_DEFAULT AS Denumire
                    FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                    WHERE ppm.ID_AnUniv = @idAn
                      AND ppm.DenumireSpecializare IS NOT NULL
                      AND LTRIM(RTRIM(ppm.DenumireSpecializare)) != ''
                      AND (
                          @fac = N'Toti'
                          OR ppm.DenumireFacultate COLLATE DATABASE_DEFAULT
                           = @fac                   COLLATE DATABASE_DEFAULT
                      )
                    ORDER BY Denumire", conn);
                cmd.Parameters.Add(new SqlParameter("@idAn", SqlDbType.Int) { Value = idAn });
                cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200)
                { Value = string.IsNullOrWhiteSpace(numeFacultate) ? "Toti" : numeFacultate.Trim() });
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var den = reader[0]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(den) && !lista.Any(x => ((dynamic)x).nume == den))
                        lista.Add(new { id = (int?)null, nume = den });
                }
            }
            catch { }
            return lista;
        }

        // -------------------------------------------------------------------------
        // Profesori — SELECT DISTINCT din View_Profesori_CF_AnUniv
        // Filtrat optional dupa facultate si departament (string, pentru compatibilitate rapoarte)
        // -------------------------------------------------------------------------
        [HttpGet("liste/profesori")]
        public ActionResult GetProfesori([FromQuery] int idAnUniv, [FromQuery] string facultate, [FromQuery] string departament)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            // Folosim TRY_CAST pentru a vedea dacă parametrul primit e ID sau Nume
            string sql = @"
        SELECT DISTINCT LTRIM(RTRIM(p.NumeIntreg)) COLLATE DATABASE_DEFAULT AS Nume
        FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
        WHERE p.ID_AnUnivCatedra = @idAn
          AND (@fac = 'Toti' OR p.ID_Facultate = TRY_CAST(@fac AS INT) OR p.DenumireFacultate = @fac)
          AND (@dept = 'Toti' OR p.ID_Catedra = TRY_CAST(@dept AS INT) OR p.DenumireCatedra = @dept)
        ORDER BY Nume";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@idAn", idAnUniv);
            cmd.Parameters.AddWithValue("@fac", facultate);
            cmd.Parameters.AddWithValue("@dept", departament);

            using var reader = cmd.ExecuteReader();
            while (reader.Read()) lista.Add(reader[0].ToString());
            return Ok(lista);
        }

        // -------------------------------------------------------------------------
        // Forme de invatamant — din DenumireFormaInv din ppm (nu lista statica)
        // -------------------------------------------------------------------------
        [HttpGet("liste/forme-invatamant")]
        public ActionResult GetFormeInvatamant([FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            return Ok(_cache.GetOrCreate($"FormeInv_{idAn}", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<string> { "Toti" };
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT
                        LTRIM(RTRIM(ppm.DenumireFormaInv)) COLLATE DATABASE_DEFAULT AS FormaInv
                    FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                    WHERE ppm.ID_AnUniv = @idAn
                      AND ppm.DenumireFormaInv IS NOT NULL
                      AND LTRIM(RTRIM(ppm.DenumireFormaInv)) != ''
                    ORDER BY FormaInv", conn);
                cmd.Parameters.Add(new SqlParameter("@idAn", SqlDbType.Int) { Value = idAn });
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var f = reader[0]?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(f)) lista.Add(f);
                }
                return lista;
            }));
        }

        // -------------------------------------------------------------------------
        // An universitar curent — expus pentru UI
        // -------------------------------------------------------------------------
        [HttpGet("liste/an-curent")]
        public ActionResult GetAnCurent() => Ok(new { idAnUniv = _idAnUnivCurent });

        #endregion

        #region ================= HELPER SQL COMUN =================

        private const string BaseDataSql = @"
            WITH VcmDedup AS (
                SELECT DISTINCT
                    ppm.ID_Profesor,
                    ppm.NumeIntreg                                       AS NumeIntregProfesor,
                    ppm.ID_AnUniv,
                    ppm.DenumireSpecializare,
                    ppm.DenumireMaterie,
                    ppm.NrSemestruDinAn,
                    ppm.Nr_Ore_Curs,
                    ppm.Nr_Ore_Seminar,
                    ppm.Nr_Ore_Laborator,
                    ppm.Nr_Ore_Proiect,
                    ppm.NrOreConventionale,
                    ppm.DenumireFacultate,
                    ppm.DenumireCatedraProfesor                          AS DenumireCatedra,
                    ppm.TitularSauSuplinitor,
                    ppm.DenumireFormaInv
                FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                WHERE ppm.ID_AnUniv = @ID_AnUniv
            ),
            BaseData AS (
                SELECT
                    vcm.NumeIntregProfesor                                          AS NumeIntreg,
                    vcm.ID_Profesor,
                    UPPER(LTRIM(RTRIM(
                        CASE WHEN CHARINDEX('+', vcm.DenumireSpecializare) > 0
                             THEN LEFT(vcm.DenumireSpecializare, CHARINDEX('+', vcm.DenumireSpecializare) - 1)
                             ELSE vcm.DenumireSpecializare END
                    ))) COLLATE DATABASE_DEFAULT                                    AS SpecializareCurata,
                    vcm.DenumireSpecializare                                        AS NumeSpecOriginal,
                    ISNULL(vcm.DenumireMaterie, 'Nedefinit')                        AS DenumireMaterie,
                    CASE WHEN vcm.TitularSauSuplinitor = 1 THEN 'Titular' ELSE 'Suplinitor' END AS TipPost,
                    ISNULL(vcm.NrSemestruDinAn, 0)                                 AS Semestru,
                    CAST(ISNULL(vcm.NrOreConventionale, 0) AS DECIMAL(10,4))       AS OreConvLinie,
                    ISNULL(vcm.Nr_Ore_Curs, 0)                                     AS OreCursLinie,
                    ISNULL(vcm.Nr_Ore_Seminar,   0)
                  + ISNULL(vcm.Nr_Ore_Laborator, 0)
                  + ISNULL(vcm.Nr_Ore_Proiect,   0)                                AS OreAplicatiiLinie,
                    LTRIM(RTRIM(ISNULL(vcm.DenumireFacultate, '')))                AS FacultateCurata,
                    vcm.DenumireCatedra,
                    UPPER(LTRIM(RTRIM(au.Denumire))) COLLATE DATABASE_DEFAULT       AS AnCurat,
                    -- FormaInv din coloana DenumireFormaInv (metadata BD) — fara LIKE pe spec
                    LTRIM(RTRIM(ISNULL(vcm.DenumireFormaInv, 'IF')))               AS FormaInv
                FROM VcmDedup vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv = au.ID_AnUniv
                WHERE vcm.ID_AnUniv = @ID_AnUniv
            )";

        #endregion

        #region ================= RAPORT 1: NORMA PROFESORI =================

        private string BuildNormaSql() => BaseDataSql + @",
            Filtrat AS (
                SELECT bd.NumeIntreg, bd.SpecializareCurata, bd.NumeSpecOriginal,
                       bd.DenumireMaterie, bd.TipPost, bd.Semestru,
                       bd.DenumireCatedra, bd.ID_Profesor,
                       bd.OreCursLinie, bd.OreAplicatiiLinie, bd.OreConvLinie, bd.FormaInv
                FROM BaseData bd
                WHERE (@an       = 'Toti' OR bd.AnCurat = @an)
                  AND (@fac      = 'Toti' OR bd.FacultateCurata COLLATE DATABASE_DEFAULT
                                           = @fac                COLLATE DATABASE_DEFAULT)
                  AND (@prof     = 'Toti' OR bd.NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR bd.FormaInv   COLLATE DATABASE_DEFAULT
                                           = @formaInv      COLLATE DATABASE_DEFAULT)
                  AND (@specs    = 'Toti' OR bd.SpecializareCurata IN
                       (SELECT UPPER(LTRIM(RTRIM(value))) FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR bd.Semestru = @semestru)
                  AND (@tipPost  = 'Toti' OR bd.TipPost = @tipPost)
                  AND (@dept     = 'Toti' OR bd.DenumireCatedra COLLATE DATABASE_DEFAULT
                                           = @dept               COLLATE DATABASE_DEFAULT)
            ),
            Dedup AS (
                SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, Semestru, TipPost,
                       MAX(OreCursLinie)      AS OreCurs,
                       MAX(OreAplicatiiLinie) AS OreAplicatii,
                       MAX(OreConvLinie)      AS OreConv,
                       MAX(DenumireCatedra)   AS DenumireCatedra,
                       MAX(FormaInv)          AS FormaInv
                FROM Filtrat
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, Semestru, TipPost
            )
            SELECT NumeIntreg, SpecializareCurata AS Specializare, DenumireMaterie AS Materie,
                   TipPost, Semestru,
                   CAST(OreCurs      AS DECIMAL(10,2)) AS OreCurs,
                   CAST(OreAplicatii AS DECIMAL(10,2)) AS OreAplicatii,
                   CAST(OreConv      AS DECIMAL(10,2)) AS OreConv,
                   DenumireCatedra, FormaInv
            FROM Dedup
            ORDER BY NumeIntreg, Specializare, Materie, Semestru
            OPTION (RECOMPILE)";

        [HttpGet("norma-profesori")]
        public async Task<IActionResult> GetNormaProfesori(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] string? specializari,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildNormaSql(), conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                result.Add(new
                {
                    Profesor = reader["NumeIntreg"]?.ToString() ?? "",
                    Specializare = reader["Specializare"]?.ToString() ?? "",
                    Materie = reader["Materie"]?.ToString() ?? "",
                    TipPost = reader["TipPost"]?.ToString() ?? "",
                    Semestru = reader["Semestru"] != DBNull.Value ? Convert.ToInt32(reader["Semestru"]) : 0,
                    OreCurs = reader["OreCurs"] != DBNull.Value ? Convert.ToDecimal(reader["OreCurs"]) : 0m,
                    OreAplicatii = reader["OreAplicatii"] != DBNull.Value ? Convert.ToDecimal(reader["OreAplicatii"]) : 0m,
                    OreConv = reader["OreConv"] != DBNull.Value ? Convert.ToDecimal(reader["OreConv"]) : 0m,
                    Departament = reader["DenumireCatedra"]?.ToString() ?? "",
                    FormaInv = reader["FormaInv"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/norma-profesori")]
        public async Task<IActionResult> ExportNormaProfesori(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] string? specializari,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"),       new DataColumn("Specializare"),
                new DataColumn("Materie"),        new DataColumn("Tip Post"),
                new DataColumn("Sem.", typeof(int)),
                new DataColumn("Ore Curs",   typeof(decimal)),
                new DataColumn("Ore Aplic.", typeof(decimal)),
                new DataColumn("Ore Conv.",  typeof(decimal)),
                new DataColumn("Departament"),    new DataColumn("Forma Inv.")
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildNormaSql(), conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                dt.Rows.Add(
                    reader["NumeIntreg"]?.ToString() ?? "",
                    reader["Specializare"]?.ToString() ?? "",
                    reader["Materie"]?.ToString() ?? "",
                    reader["TipPost"]?.ToString() ?? "",
                    reader["Semestru"] != DBNull.Value ? Convert.ToInt32(reader["Semestru"]) : 0,
                    reader["OreCurs"] != DBNull.Value ? Convert.ToDecimal(reader["OreCurs"]) : 0m,
                    reader["OreAplicatii"] != DBNull.Value ? Convert.ToDecimal(reader["OreAplicatii"]) : 0m,
                    reader["OreConv"] != DBNull.Value ? Convert.ToDecimal(reader["OreConv"]) : 0m,
                    reader["DenumireCatedra"]?.ToString() ?? "",
                    reader["FormaInv"]?.ToString() ?? "");
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme Detaliate");
            ws.Cell(1, 1).Value = $"Detaliere norme | An: {idAn} | Facultate: {facultate ?? "Toti"}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            ws.Cell(2, 1).Value = "NOTA: Ore Conv. reflecta norma totala din universitate, independent de filtrele aplicate.";
            ws.Cell(2, 1).Style.Font.Italic = true;
            ws.Cell(2, 1).Style.Font.FontColor = XLColor.Gray;
            ws.Range(2, 1, 2, dt.Columns.Count).Merge();
            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Aplic.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
            StyleHeader(ws.Range(4, 1, 4, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"NormaDetaliata_{idAn}.xlsx");
        }

        #endregion

        #region ================= RAPORT 2: ORE PER PROGRAM =================

        private string BuildOreProfProgramSql() => BaseDataSql + @",
            OreProgram AS (
                SELECT NumeIntreg, ID_Profesor, SpecializareCurata,
                       DenumireMaterie, Semestru, TipPost,
                       MAX(OreConvLinie)      AS OreConv,
                       MAX(OreCursLinie)      AS OreCurs,
                       MAX(OreAplicatiiLinie) AS OreAplicatii
                FROM BaseData
                WHERE (@an       = 'Toti' OR AnCurat = @an)
                  AND (@fac      = 'Toti' OR FacultateCurata COLLATE DATABASE_DEFAULT
                                           = @fac             COLLATE DATABASE_DEFAULT)
                  AND (@prof     = 'Toti' OR NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR FormaInv   COLLATE DATABASE_DEFAULT
                                           = @formaInv   COLLATE DATABASE_DEFAULT)
                  AND (@specs    = 'Toti' OR SpecializareCurata IN
                       (SELECT UPPER(LTRIM(RTRIM(value))) FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost  = 'Toti' OR TipPost = @tipPost)
                  AND (@dept     = 'Toti' OR DenumireCatedra COLLATE DATABASE_DEFAULT
                                           = @dept            COLLATE DATABASE_DEFAULT)
                GROUP BY NumeIntreg, ID_Profesor, SpecializareCurata, DenumireMaterie, Semestru, TipPost
            ),
            AgregatProgram AS (
                SELECT NumeIntreg, ID_Profesor, SpecializareCurata AS ProgramStudiu,
                       SUM(OreConv)      AS OreConvProgram,
                       SUM(OreCurs)      AS OreCursProgram,
                       SUM(OreAplicatii) AS OreAplicatiiProgram
                FROM OreProgram
                GROUP BY NumeIntreg, ID_Profesor, SpecializareCurata
                HAVING SUM(OreConv) > 0
            ),
            TotalProfesor AS (
                SELECT NumeIntreg, SUM(OreConvProgram) AS TotalPost
                FROM AgregatProgram GROUP BY NumeIntreg
            )
            SELECT a.NumeIntreg                                              AS Profesor,
                   ISNULL(a.ProgramStudiu, 'Nespecificat')                   AS ProgramStudiu,
                   CAST(a.OreConvProgram      AS DECIMAL(10,2))              AS NrOreConv,
                   CAST(a.OreCursProgram      AS DECIMAL(10,2))              AS NrOreCurs,
                   CAST(a.OreAplicatiiProgram AS DECIMAL(10,2))              AS NrOreAplicatii,
                   CAST(t.TotalPost           AS DECIMAL(10,2))              AS TotalPost,
                   CAST(CASE WHEN ISNULL(t.TotalPost,0)=0 THEN 0
                             ELSE (a.OreConvProgram/t.TotalPost)*100
                        END AS DECIMAL(10,2))                                AS ProcentPost
            FROM AgregatProgram a
            LEFT JOIN TotalProfesor t ON t.NumeIntreg = a.NumeIntreg
            ORDER BY a.NumeIntreg, a.OreConvProgram DESC
            OPTION (RECOMPILE)";

        [HttpGet("ore-profesor-program")]
        public async Task<IActionResult> GetOreProfProgram(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? specializari, [FromQuery] string? profesor,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] string? departament = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildOreProfProgramSql(), conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                result.Add(new
                {
                    Profesor = reader["Profesor"]?.ToString() ?? "",
                    ProgramStudiu = reader["ProgramStudiu"]?.ToString() ?? "",
                    OreConvProgram = reader["NrOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["NrOreConv"]) : 0m,
                    TotalPost = reader["TotalPost"] != DBNull.Value ? Convert.ToDecimal(reader["TotalPost"]) : 0m,
                    ProcentPost = reader["ProcentPost"] != DBNull.Value ? Convert.ToDecimal(reader["ProcentPost"]) : 0m
                });
            return Ok(result);
        }

        [HttpGet("export/ore-program")]
        public async Task<IActionResult> ExportOreProgramExcel(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? specializari, [FromQuery] string? profesor,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] string? departament = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"),        new DataColumn("Program Studiu"),
                new DataColumn("Ore Conv.",  typeof(decimal)),
                new DataColumn("Total Post", typeof(decimal)),
                new DataColumn("Procent %",  typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildOreProfProgramSql(), conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                dt.Rows.Add(
                    reader["Profesor"]?.ToString() ?? "",
                    reader["ProgramStudiu"]?.ToString() ?? "",
                    reader["NrOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["NrOreConv"]) : 0m,
                    reader["TotalPost"] != DBNull.Value ? Convert.ToDecimal(reader["TotalPost"]) : 0m,
                    reader["ProcentPost"] != DBNull.Value ? Convert.ToDecimal(reader["ProcentPost"]) : 0m);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Distributie Ore");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"OreProgram_{idAn}.xlsx");
        }

        #endregion

        #region ================= RAPORT 3: NORME TOTALURI =================

        private string BuildNormaTotaluriSql() => @"
            WITH ProfDept AS (
                SELECT p.ID_Profesor,
                       MIN(p.DenumireCatedra)   AS DeptProfesor,
                       MIN(p.DenumireFacultate) AS FacProfesor
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                WHERE p.ID_AnUnivCatedra = @ID_AnUniv
                GROUP BY p.ID_Profesor
            ),
            DateBrute AS (
                SELECT
                    ppm.NumeIntreg                                                  AS NumeComplet,
                    ppm.ID_Profesor,
                    ISNULL(pd.DeptProfesor, 'Nespecificat')                         AS Departament,
                    ISNULL(pd.FacProfesor,  'Nespecificat')                         AS Facultate,
                    ppm.ID_TipGradDidacticPost                                      AS ID_TipGrad,
                    CAST(ISNULL(ppm.NrOreConventionale, 0) AS DECIMAL(10,4))        AS OreConv,
                    ppm.DenumireMaterie,
                    ISNULL(ppm.NrSemestruDinAn, 0)                                  AS Semestru,
                    CASE WHEN ppm.TitularSauSuplinitor = 1 THEN 'Titular' ELSE 'Suplinitor' END AS TipPost,
                    -- FormaInv din coloana BD, nu din LIKE pe spec (metadata-driven)
                    LTRIM(RTRIM(ISNULL(ppm.DenumireFormaInv, 'IF')))               AS FormaInv
                FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                LEFT JOIN ProfDept pd ON pd.ID_Profesor = ppm.ID_Profesor
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON ppm.ID_AnUniv = au.ID_AnUniv
                WHERE ppm.ID_AnUniv = @ID_AnUniv
                  AND (@an   = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac  = 'Toti' OR ISNULL(pd.FacProfesor,  '') COLLATE DATABASE_DEFAULT
                                       = @fac                          COLLATE DATABASE_DEFAULT)
                  AND (@prof = 'Toti' OR ppm.NumeIntreg = @prof)
                  AND (@dept = 'Toti' OR ISNULL(pd.DeptProfesor, '') COLLATE DATABASE_DEFAULT
                                       = @dept                         COLLATE DATABASE_DEFAULT)
            ),
            Dedup AS (
                SELECT NumeComplet, ID_Profesor, TipPost, FormaInv,
                       DenumireMaterie, Semestru,
                       MAX(OreConv)     AS OreConvDedup,
                       MAX(Departament) AS Departament,
                       MAX(Facultate)   AS Facultate,
                       MAX(ID_TipGrad)  AS ID_TipGrad
                FROM DateBrute
                GROUP BY NumeComplet, ID_Profesor, TipPost, FormaInv, DenumireMaterie, Semestru
            ),
            Agreg AS (
                SELECT NumeComplet, ID_Profesor, MAX(Departament) AS Departament,
                       MAX(Facultate) AS Facultate, MAX(ID_TipGrad) AS ID_TipGrad, TipPost,
                       -- Sumele pe FormaInv vin din metadata BD (DenumireFormaInv normalizat)
                       CAST(ISNULL(SUM(CASE WHEN FormaInv LIKE '%IF%'  AND FormaInv NOT LIKE '%IFR%'
                                            THEN OreConvDedup ELSE 0 END), 0) AS DECIMAL(10,2)) AS OreIF,
                       CAST(ISNULL(SUM(CASE WHEN FormaInv LIKE '%ID%'  AND FormaInv NOT LIKE '%IFR%'
                                            THEN OreConvDedup ELSE 0 END), 0) AS DECIMAL(10,2)) AS OreID,
                       CAST(ISNULL(SUM(CASE WHEN FormaInv LIKE '%IFR%'
                                            THEN OreConvDedup ELSE 0 END), 0) AS DECIMAL(10,2)) AS OreIFR,
                       CAST(ISNULL(SUM(OreConvDedup), 0) AS DECIMAL(10,2))                       AS TotalOreConv
                FROM Dedup
                GROUP BY NumeComplet, ID_Profesor, TipPost
            ),
            Norme AS (
                -- Norme din BD — fara fallback hardcodat in SQL
                SELECT ID_TipGradDidactic, NrOreConventionaleTitular AS Norma
                FROM [AGSIS].[pi].[NormaOreConventionale]
                WHERE ID_AnUniv = @ID_AnUniv
                  AND NrOreConventionaleTitular IS NOT NULL
                  AND NrOreConventionaleTitular > 0
            )
            SELECT ag.NumeComplet, ag.ID_Profesor, ag.Departament, ag.Facultate, ag.TipPost,
                   ag.OreIF, ag.OreID, ag.OreIFR, ag.TotalOreConv,
                   -- TotalAnual = TotalOreConv * Norma din BD (NULL daca gradul nu are norma)
                   CAST(ag.TotalOreConv * n.Norma AS DECIMAL(10,2)) AS TotalAnual,
                   n.Norma                                            AS NormaLegala
            FROM Agreg ag
            LEFT JOIN Norme n ON n.ID_TipGradDidactic = ag.ID_TipGrad
            ORDER BY ag.NumeComplet, ag.TipPost DESC";

        [HttpGet("norma-totaluri")]
        public async Task<IActionResult> GetNormaTotaluri(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildNormaTotaluriSql(), conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAn });
            cmd.Parameters.Add(new SqlParameter("@an", SqlDbType.NVarChar, 200) { Value = anUniv ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@prof", SqlDbType.NVarChar, 200) { Value = profesor ?? "Toti" });
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                result.Add(new
                {
                    Profesor = reader["NumeComplet"]?.ToString() ?? "",
                    Departament = reader["Departament"]?.ToString() ?? "",
                    Facultate = reader["Facultate"]?.ToString() ?? "",
                    TipNorma = reader["TipPost"]?.ToString() ?? "",
                    OreIF = reader["OreIF"] != DBNull.Value ? Convert.ToDecimal(reader["OreIF"]) : 0m,
                    OreID = reader["OreID"] != DBNull.Value ? Convert.ToDecimal(reader["OreID"]) : 0m,
                    OreIFR = reader["OreIFR"] != DBNull.Value ? Convert.ToDecimal(reader["OreIFR"]) : 0m,
                    TotalOreConv = reader["TotalOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalOreConv"]) : 0m,
                    TotalAnual = reader["TotalAnual"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAnual"]) : 0m,
                    NormaLegala = reader["NormaLegala"] != DBNull.Value ? (decimal?)Convert.ToDecimal(reader["NormaLegala"]) : null
                });
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public async Task<IActionResult> ExportNormaTotaluri(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"),       new DataColumn("Departament"),
                new DataColumn("Facultate"),      new DataColumn("Tip Post"),
                new DataColumn("Norma Legala", typeof(decimal)),
                new DataColumn("Ore IF",    typeof(decimal)), new DataColumn("Ore ID",    typeof(decimal)),
                new DataColumn("Ore IFR",   typeof(decimal)),
                new DataColumn("Total Ore Conv.", typeof(decimal)),
                new DataColumn("Total Anual",     typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildNormaTotaluriSql(), conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAn });
            cmd.Parameters.Add(new SqlParameter("@an", SqlDbType.NVarChar, 200) { Value = anUniv ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@prof", SqlDbType.NVarChar, 200) { Value = profesor ?? "Toti" });
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                dt.Rows.Add(
                    reader["NumeComplet"]?.ToString() ?? "",
                    reader["Departament"]?.ToString() ?? "",
                    reader["Facultate"]?.ToString() ?? "",
                    reader["TipPost"]?.ToString() ?? "",
                    reader["NormaLegala"] != DBNull.Value ? Convert.ToDecimal(reader["NormaLegala"]) : DBNull.Value,
                    reader["OreIF"] != DBNull.Value ? Convert.ToDecimal(reader["OreIF"]) : 0m,
                    reader["OreID"] != DBNull.Value ? Convert.ToDecimal(reader["OreID"]) : 0m,
                    reader["OreIFR"] != DBNull.Value ? Convert.ToDecimal(reader["OreIFR"]) : 0m,
                    reader["TotalOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalOreConv"]) : 0m,
                    reader["TotalAnual"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAnual"]) : DBNull.Value);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme Totaluri");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            foreach (var col in new[] { "Ore IF", "Ore ID", "Ore IFR", "Total Ore Conv.", "Total Anual" })
                tbl.Field(col).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Totaluri_Norme_{idAn}.xlsx");
        }

        #endregion

        #region ================= RAPORT 4: LIMBI STRAINE =================

        // Sursa oficiala pentru LimbaDePredare: coloana ppm.LimbaDePredare (nu filtre pe denumire materie)
        private string BuildLimbiSql() => BaseDataSql + @",
            LimbiBase AS (
                SELECT DISTINCT
                    bd.NumeIntreg AS NumeComplet,
                    bd.DenumireCatedra,
                    bd.FacultateCurata AS Facultate,
                    ppm2.LimbaDePredare,
                    bd.DenumireMaterie,
                    bd.Semestru,
                    bd.TipPost,
                    MAX(bd.OreConvLinie) AS OreConvDedup
                FROM BaseData bd
                INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm2
                    ON  ppm2.ID_Profesor  = bd.ID_Profesor
                    AND ppm2.ID_AnUniv    = @ID_AnUniv
                    AND ppm2.LimbaDePredare IS NOT NULL
                    AND LTRIM(RTRIM(ppm2.LimbaDePredare)) != ''
                WHERE (@an       = 'Toti' OR bd.AnCurat = @an)
                  AND (@fac      = 'Toti' OR bd.FacultateCurata COLLATE DATABASE_DEFAULT
                                           = @fac                COLLATE DATABASE_DEFAULT)
                  AND (@prof     = 'Toti' OR bd.NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR bd.FormaInv   COLLATE DATABASE_DEFAULT
                                           = @formaInv      COLLATE DATABASE_DEFAULT)
                  AND (@specs    = 'Toti' OR bd.SpecializareCurata IN
                       (SELECT UPPER(LTRIM(RTRIM(value))) FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR bd.Semestru = @semestru)
                  AND (@tipPost  = 'Toti' OR bd.TipPost = @tipPost)
                  AND (@dept     = 'Toti' OR bd.DenumireCatedra COLLATE DATABASE_DEFAULT
                                           = @dept               COLLATE DATABASE_DEFAULT)
                GROUP BY bd.NumeIntreg, bd.DenumireCatedra, bd.FacultateCurata,
                         ppm2.LimbaDePredare, bd.DenumireMaterie, bd.Semestru, bd.TipPost
            ),
            Dedup AS (
                SELECT NumeComplet, DenumireMaterie, Semestru, TipPost, LimbaDePredare,
                       MAX(OreConvDedup) AS OreConvDedup
                FROM LimbiBase
                GROUP BY NumeComplet, DenumireMaterie, Semestru, TipPost, LimbaDePredare
            )
            SELECT NumeComplet,
                   LimbaDePredare,
                   CAST(ROUND(SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConvDedup ELSE 0 END)*14,2) AS DECIMAL(10,2)) AS Sem1,
                   CAST(ROUND(SUM(CASE WHEN Semestru IN (2,4,6,8,10,12)THEN OreConvDedup ELSE 0 END)*14,2) AS DECIMAL(10,2)) AS Sem2,
                   CAST(ROUND(SUM(OreConvDedup)*14, 2) AS DECIMAL(10,2))                                                       AS Total
            FROM Dedup
            GROUP BY NumeComplet, LimbaDePredare
            HAVING SUM(OreConvDedup) > 0
            ORDER BY LimbaDePredare, NumeComplet";

        private void AddLimbiParams(SqlCommand cmd, int idAn,
            string? anUniv, string? facultate, string? departament, string? formaInvatamant,
            string? profesor, string? specializari, int semestru, string? tipPost)
        {
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAn });
            cmd.Parameters.Add(new SqlParameter("@an", SqlDbType.NVarChar, 200) { Value = anUniv ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@formaInv", SqlDbType.NVarChar, 200) { Value = formaInvatamant ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@prof", SqlDbType.NVarChar, 200) { Value = profesor ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@specs", SqlDbType.NVarChar, 2000) { Value = string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari });
            cmd.Parameters.Add(new SqlParameter("@semestru", SqlDbType.Int) { Value = semestru });
            cmd.Parameters.Add(new SqlParameter("@tipPost", SqlDbType.NVarChar, 50) { Value = tipPost ?? "Toti" });
        }

        [HttpGet("limbi-straine")]
        public async Task<IActionResult> GetLimbiStraine(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] string? specializari,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildLimbiSql(), conn);
            cmd.CommandTimeout = 120;
            AddLimbiParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await reader.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    NumeIntreg = reader["NumeComplet"]?.ToString() ?? "",
                    Limba = reader["LimbaDePredare"]?.ToString() ?? "",
                    OreSem1 = reader["Sem1"] != DBNull.Value ? Convert.ToDecimal(reader["Sem1"]) : 0m,
                    OreSem2 = reader["Sem2"] != DBNull.Value ? Convert.ToDecimal(reader["Sem2"]) : 0m,
                    Total = reader["Total"] != DBNull.Value ? Convert.ToDecimal(reader["Total"]) : 0m
                });
            return Ok(result);
        }

        [HttpGet("export/limbi-straine")]
        public async Task<IActionResult> ExportLimbiStraine(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] string? specializari,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nr.",   typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Limba"),
                new DataColumn("Sem. 1", typeof(decimal)), new DataColumn("Sem. 2", typeof(decimal)),
                new DataColumn("Total",  typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildLimbiSql(), conn);
            cmd.CommandTimeout = 120;
            AddLimbiParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            int nr = 1;
            while (await reader.ReadAsync())
                dt.Rows.Add(nr++,
                    reader["NumeComplet"]?.ToString() ?? "",
                    reader["LimbaDePredare"]?.ToString() ?? "",
                    reader["Sem1"] != DBNull.Value ? Convert.ToDecimal(reader["Sem1"]) : 0m,
                    reader["Sem2"] != DBNull.Value ? Convert.ToDecimal(reader["Sem2"]) : 0m,
                    reader["Total"] != DBNull.Value ? Convert.ToDecimal(reader["Total"]) : 0m);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Limbi Straine");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            foreach (var col in new[] { "Sem. 1", "Sem. 2", "Total" })
                tbl.Field(col).TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Raport_Limbi_Straine_{idAn}.xlsx");
        }

        #endregion

        #region ================= RAPORT 5: DISCIPLINE PREDATE =================

        private string BuildDisciplineSql() => BaseDataSql + @",
            DU AS (
                SELECT DISTINCT
                    bd.NumeIntreg, bd.DenumireCatedra, bd.ID_Profesor, bd.FormaInv,
                    bd.DenumireMaterie, bd.OreCursLinie, bd.OreAplicatiiLinie
                FROM BaseData bd
                WHERE (@an       = 'Toti' OR bd.AnCurat = @an)
                  AND (@fac      = 'Toti' OR bd.FacultateCurata COLLATE DATABASE_DEFAULT
                                           = @fac                COLLATE DATABASE_DEFAULT)
                  AND (@prof     = 'Toti' OR bd.NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR bd.FormaInv   COLLATE DATABASE_DEFAULT
                                           = @formaInv      COLLATE DATABASE_DEFAULT)
                  AND (@specs    = 'Toti' OR bd.SpecializareCurata IN
                       (SELECT UPPER(LTRIM(RTRIM(value))) FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR bd.Semestru = @semestru)
                  AND (@tipPost  = 'Toti' OR bd.TipPost = @tipPost)
                  AND (@dept     = 'Toti' OR bd.DenumireCatedra COLLATE DATABASE_DEFAULT
                                           = @dept               COLLATE DATABASE_DEFAULT)
            ),
            DU_Tip AS (
                SELECT NumeIntreg, FormaInv, DenumireMaterie,
                       MAX(OreCursLinie)      AS MaxCurs,
                       MAX(OreAplicatiiLinie) AS MaxAplicatii,
                       MIN(DenumireCatedra)   AS DenumireCatedra,
                       MAX(ID_Profesor)       AS ID_Profesor
                FROM DU GROUP BY NumeIntreg, FormaInv, DenumireMaterie
            ),
            Prof AS (
                SELECT NumeIntreg, MIN(DenumireCatedra) AS DenumireCatedra,
                       MAX(ID_Profesor) AS ID_Profesor, FormaInv
                FROM DU_Tip GROUP BY NumeIntreg, FormaInv
            )
            SELECT p.NumeIntreg, p.DenumireCatedra AS Departament, p.ID_Profesor, p.FormaInv,
                   STUFF((
                       SELECT N', ' + CAST(d2.DenumireMaterie AS NVARCHAR(MAX))
                           + N' ('
                           + STUFF(
                               CASE WHEN d2.MaxCurs      > 0 THEN N', Curs'        ELSE N'' END
                             + CASE WHEN d2.MaxAplicatii > 0 THEN N', Seminar/Lab' ELSE N'' END,
                             1, 2, N'')
                           + N')'
                       FROM DU_Tip d2
                       WHERE d2.NumeIntreg COLLATE DATABASE_DEFAULT = p.NumeIntreg COLLATE DATABASE_DEFAULT
                         AND d2.FormaInv   COLLATE DATABASE_DEFAULT = p.FormaInv   COLLATE DATABASE_DEFAULT
                       ORDER BY d2.DenumireMaterie
                       FOR XML PATH(''), TYPE
                   ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Discipline
            FROM Prof p
            ORDER BY p.FormaInv, p.NumeIntreg";

        [HttpGet("discipline-predate")]
        public async Task<IActionResult> GetDisciplinePredate(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] string? specializari,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildDisciplineSql(), conn);
            cmd.CommandTimeout = 180;
            AddBaseParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                result.Add(new
                {
                    Profesor = reader["NumeIntreg"]?.ToString() ?? "",
                    Departament = reader["Departament"]?.ToString() ?? "",
                    FormaInvatamant = reader["FormaInv"]?.ToString() ?? "",
                    Discipline = reader["Discipline"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/discipline-predate")]
        public async Task<IActionResult> ExportDisciplinePredate(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] string? profesor,
            [FromQuery] string? specializari,
            [FromQuery] int semestru = 0,
            [FromQuery] string tipPost = "Toti",
            [FromQuery] string? formaInvatamant = "Toti",
            [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            // Colecteaza toate formele din BD (nu lista statica IF/ID/IFR)
            var datePerForma = new Dictionary<string, (DataTable Dt, int Nr)>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(BuildDisciplineSql(), conn);
            cmd.CommandTimeout = 180;
            AddBaseParams(cmd, idAn, anUniv, facultate, departament, formaInvatamant,
                profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                string forma = reader["FormaInv"]?.ToString() ?? "Necunoscuta";
                if (!datePerForma.ContainsKey(forma))
                {
                    var dt = new DataTable();
                    dt.Columns.AddRange(new[] {
                        new DataColumn("Nr.", typeof(int)), new DataColumn("Profesor"),
                        new DataColumn("Departament"),      new DataColumn("Discipline Predate")
                    });
                    datePerForma[forma] = (dt, 1);
                }
                var (table, nr) = datePerForma[forma];
                table.Rows.Add(nr, reader["NumeIntreg"]?.ToString() ?? "",
                    reader["Departament"]?.ToString() ?? "",
                    reader["Discipline"]?.ToString() ?? "");
                datePerForma[forma] = (table, nr + 1);
            }

            using var memZip = new MemoryStream();
            using (var archive = new ZipArchive(memZip, ZipArchiveMode.Create, leaveOpen: true))
            {
                if (datePerForma.Count == 0)
                {
                    var re = archive.CreateEntry("README.txt");
                    using var rw = new StreamWriter(re.Open());
                    await rw.WriteAsync($"Nu exista date pentru parametrii selectati (An: {idAn}).");
                }
                foreach (var kvp in datePerForma)
                {
                    if (kvp.Value.Dt.Rows.Count == 0) continue;
                    var safeName = string.Concat(kvp.Key.Take(30).Select(c =>
                        Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                    var entry = archive.CreateEntry($"Discipline_Predate_{safeName}.xlsx");
                    using var entryStream = entry.Open();
                    using var wb = new XLWorkbook();
                    string sheetName = $"Discipline {kvp.Key}";
                    if (sheetName.Length > 31) sheetName = sheetName[..31];
                    var ws = wb.Worksheets.Add(sheetName);
                    ws.Cell(1, 1).Value = $"Discipline Predate - {kvp.Key} | An: {idAn}";
                    ws.Cell(1, 1).Style.Font.Bold = true;
                    ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
                    ws.Range(1, 1, 1, 4).Merge();
                    var tbl = ws.Cell(3, 1).InsertTable(kvp.Value.Dt);
                    tbl.Theme = XLTableTheme.None;
                    ws.Columns(1, 3).AdjustToContents();
                    ws.Column(4).Width = 80;
                    ws.Column(4).Style.Alignment.WrapText = true;
                    StyleHeader(ws.Range(3, 1, 3, 4));
                    using var wbStream = new MemoryStream(); wb.SaveAs(wbStream);
                    wbStream.Position = 0;
                    await wbStream.CopyToAsync(entryStream);
                }
            }
            memZip.Position = 0;
            return File(memZip.ToArray(), "application/zip",
                $"Discipline_Predate_{idAn}.zip");
        }

        #endregion

        #region ================= RAPORT 6: TITULARI =================

        private const string SqlTitulariRaport = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                MIN(p.NumeIntreg)                                    AS NumeIntreg,
                MIN(p.DenumireFacultate)                             AS Facultate,
                MIN(p.DenumireCatedra)                               AS Departament,
                MIN(p.DenumireGradDidacticAnUniv)                    AS GradDidactic
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                ON  ppm.ID_Profesor         = p.ID_Profesor
                AND ppm.ID_AnUniv           = @ID_AnUniv
                AND ppm.NrOreConventionale   > 0
            WHERE p.ID_AnUnivCatedra = @ID_AnUniv
              AND p.TitularAnUniv    = 1
              AND (@fac  = 'Toti' OR p.DenumireFacultate COLLATE DATABASE_DEFAULT
                                   = @fac                 COLLATE DATABASE_DEFAULT)
              AND (@dept = 'Toti' OR p.DenumireCatedra   COLLATE DATABASE_DEFAULT
                                   = @dept                COLLATE DATABASE_DEFAULT)
            GROUP BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))
            HAVING SUM(ppm.NrOreConventionale) > 0
            ORDER BY MIN(p.NumeIntreg)";

        [HttpGet("titulari")]
        public async Task<IActionResult> GetTitulari(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulariRaport, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAn });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            int nr = 1;
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = reader["NumeIntreg"]?.ToString() ?? "",
                    Departament = reader["Departament"]?.ToString() ?? "",
                    Facultate = reader["Facultate"]?.ToString() ?? "",
                    Grad = reader["GradDidactic"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/titulari")]
        public async Task<IActionResult> ExportTitulari(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nr.",  typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Departament"),       new DataColumn("Facultate"),
                new DataColumn("Grad")
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlTitulariRaport, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAn });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            int nr = 1;
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                dt.Rows.Add(nr++,
                    reader["NumeIntreg"]?.ToString() ?? "",
                    reader["Departament"]?.ToString() ?? "",
                    reader["Facultate"]?.ToString() ?? "",
                    reader["GradDidactic"]?.ToString() ?? "");
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Titulari");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Nr.").TotalsRowFunction = XLTotalsRowFunction.Count;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Cadre_Didactice_Titulare_{idAn}.xlsx");
        }

        #endregion

        #region ================= RAPORT 7: COLABORATORI =================

        private const string SqlColaboratori = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                MIN(p.NumeIntreg)                                    AS NumeIntreg,
                MIN(p.DenumireFacultate)                             AS Facultate,
                MIN(p.DenumireCatedra)                               AS Departament,
                MIN(p.DenumireGradDidacticAnUniv)                    AS GradDidactic
            FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
            INNER JOIN [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
                ON  ppm.ID_Profesor         = p.ID_Profesor
                AND ppm.ID_AnUniv           = @ID_AnUniv
                AND ppm.NrOreConventionale   > 0
            WHERE p.ID_AnUnivCatedra = @ID_AnUniv
              AND p.TitularAnUniv    = 0
              AND (@fac  = 'Toti' OR p.DenumireFacultate COLLATE DATABASE_DEFAULT
                                   = @fac                 COLLATE DATABASE_DEFAULT)
              AND (@dept = 'Toti' OR p.DenumireCatedra   COLLATE DATABASE_DEFAULT
                                   = @dept                COLLATE DATABASE_DEFAULT)
            GROUP BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))
            HAVING SUM(ppm.NrOreConventionale) > 0
            ORDER BY MIN(p.NumeIntreg)";

        [HttpGet("colaboratori")]
        public async Task<IActionResult> GetColaboratori(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColaboratori, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAn });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            int nr = 1;
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = reader["NumeIntreg"]?.ToString() ?? "",
                    Departament = reader["Departament"]?.ToString() ?? "",
                    Facultate = reader["Facultate"]?.ToString() ?? "",
                    Grad = reader["GradDidactic"]?.ToString() ?? ""
                });
            return Ok(result);
        }

        [HttpGet("export/colaboratori")]
        public async Task<IActionResult> ExportColaboratori(
            [FromQuery] string? anUniv, [FromQuery] string? facultate,
            [FromQuery] string? departament, [FromQuery] int? idAnUniv = null)
        {
            int idAn = idAnUniv ?? _idAnUnivCurent;
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nr.",  typeof(int)), new DataColumn("Profesor"),
                new DataColumn("Departament"),       new DataColumn("Facultate"),
                new DataColumn("Grad")
            });
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlColaboratori, conn);
            cmd.CommandTimeout = 60;
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAn });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            int nr = 1;
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                dt.Rows.Add(nr++,
                    reader["NumeIntreg"]?.ToString() ?? "",
                    reader["Departament"]?.ToString() ?? "",
                    reader["Facultate"]?.ToString() ?? "",
                    reader["GradDidactic"]?.ToString() ?? "");
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Colaboratori");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Nr.").TotalsRowFunction = XLTotalsRowFunction.Count;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL";
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            ws.Columns().AdjustToContents();
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Cadre_Didactice_Colaboratori_{idAn}.xlsx");
        }

        #endregion

        #region ================= HELPERS =================

        private void AddBaseParams(SqlCommand cmd, int idAnUniv,
            string? anUniv, string? facultate, string? departament,
            string? formaInvatamant, string? profesor, string? specializari,
            int semestru, string? tipPost)
        {
            cmd.Parameters.Add(new SqlParameter("@ID_AnUniv", SqlDbType.Int) { Value = idAnUniv });
            cmd.Parameters.Add(new SqlParameter("@an", SqlDbType.NVarChar, 200) { Value = anUniv ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@fac", SqlDbType.NVarChar, 200) { Value = facultate ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@dept", SqlDbType.NVarChar, 200) { Value = departament ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@formaInv", SqlDbType.NVarChar, 200) { Value = formaInvatamant ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@prof", SqlDbType.NVarChar, 200) { Value = profesor ?? "Toti" });
            cmd.Parameters.Add(new SqlParameter("@specs", SqlDbType.NVarChar, 2000) { Value = string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari });
            cmd.Parameters.Add(new SqlParameter("@semestru", SqlDbType.Int) { Value = semestru });
            cmd.Parameters.Add(new SqlParameter("@tipPost", SqlDbType.NVarChar, 50) { Value = tipPost ?? "Toti" });
        }

        private void StyleHeader(IXLRange range)
        {
            range.Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColorHex);
            range.Style.Font.FontColor = XLColor.White;
            range.Style.Font.Bold = true;
        }

        private static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }
            return result;
        }

        // Helpers pentru citire sigura din SqlDataReader (coloana poate lipsi in SP)
        private static int? TryGetInt(SqlDataReader r, string col)
        {
            try { return r[col] != DBNull.Value ? (int?)Convert.ToInt32(r[col]) : null; }
            catch { return null; }
        }

        private static string? TryGetString(SqlDataReader r, string col)
        {
            try { return r[col]?.ToString(); }
            catch { return null; }
        }

        #endregion
    }
}