using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RapoarteController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;
        private readonly IMemoryCache _cache;

        public RapoarteController(IConfiguration configuration, IMemoryCache cache)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        // =========================================================================
        // 1. ENDPOINT-URI PENTRU LISTE (FILTRE) - LOGICĂ VALIDATĂ
        // =========================================================================

        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAni()
        {
            return Ok(_cache.GetOrCreate("ListaAniUniv", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<object>();

                lista.Add(new { id = "AN UNIVERSITAR 2025-2026", nume = "AN UNIVERSITAR 2025-2026" });

                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    string sql = @"
                        SELECT DISTINCT UPPER(LTRIM(RTRIM(REPLACE(DenumireAnUniv, CHAR(9), '')))) COLLATE DATABASE_DEFAULT as AnCurat 
                        FROM [agsis_dw].[dbo].[Cazare] 
                        WHERE DenumireAnUniv IS NOT NULL 
                        ORDER BY AnCurat DESC";

                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var an = reader["AnCurat"].ToString();
                        if (an != "AN UNIVERSITAR 2025-2026")
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
                            UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT as FacCurata
                        FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
                        WHERE ppm.DenumireFacultate IS NOT NULL
                        ORDER BY FacCurata ASC";

                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read()) { lista.Add(reader["FacCurata"].ToString()!); }
                }
                return lista;
            }));
        }
        [HttpGet("liste/specializari-per-facultate")]
        public ActionResult GetSpecializari(string anUniv, string numeFacultate)
        {
            var lista = new List<string> { "Toti" };

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql;

                // Dacă e selectată o facultate, folosim logica "inteligentă" cu ID-ul
                if (!string.IsNullOrEmpty(numeFacultate) && numeFacultate != "Toti")
                {
                    sql = @"
                    DECLARE @TargetFacId INT;
                    
                    SELECT TOP 1 @TargetFacId = ID_FacultateSpecializare
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie]
                    WHERE UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) = @fac
                    GROUP BY ID_FacultateSpecializare
                    ORDER BY COUNT(*) DESC;

                    -- LOGICA DE CURĂȚARE EXTINSĂ
                    SELECT DISTINCT 
                        UPPER(LTRIM(RTRIM(
                            REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                                CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                     THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                     ELSE sf.DenumireSpecializare END, 
                            ' - CORECT', ''),   -- Stergem ' - CORECT'
                            ' CORECT', ''),     -- Stergem ' CORECT'
                            ' - COPIE', ''),    -- Stergem ' - COPIE'
                            'Ș', 'S'), 'Ț', 'T')
                        ))) as SpecCurata
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    WHERE sf.ID_Facultate = @TargetFacId 
                      AND sf.DenumireSpecializare IS NOT NULL
                    ORDER BY SpecCurata";
                }
                else
                {
                    // Logică pentru "Toti"
                    sql = @"
                    SELECT DISTINCT 
                        UPPER(LTRIM(RTRIM(
                            REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                                CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                     THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                     ELSE sf.DenumireSpecializare END, 
                            ' - CORECT', ''),
                            ' CORECT', ''),
                            ' - COPIE', ''),
                            'Ș', 'S'), 'Ț', 'T')
                        ))) as SpecCurata
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    WHERE sf.DenumireSpecializare IS NOT NULL
                    ORDER BY SpecCurata";
                }

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                if (!string.IsNullOrEmpty(numeFacultate) && numeFacultate != "Toti")
                {
                    cmd.Parameters.AddWithValue("@fac", numeFacultate);
                }

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string val = reader["SpecCurata"].ToString()!;
                    // Un ultim filtru de siguranță: nu adăugăm duplicate sau goale
                    if (!string.IsNullOrWhiteSpace(val) && !lista.Contains(val))
                        lista.Add(val);
                }
            }
            return Ok(lista);
        }

        [HttpGet("liste/profesori-per-specializari")]
        public ActionResult GetProfesori(string anUniv, string facultate, string specializari)
        {
            var lista = new List<string> { "Toti" };
            bool toateSpecializarile = (string.IsNullOrEmpty(specializari) || specializari == "Toti");

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                // La profesori păstrăm logica veche (unde predau ei), dar e ok acum că specializările sunt filtrate corect
                string sql = @"
                SELECT DISTINCT ppm.NumeIntreg
                FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                INNER JOIN (
                    SELECT DISTINCT ID_AnUniv, DenumireAnUniv FROM [agsis_dw].[dbo].[Cazare]
                ) cz ON ppm.ID_AnUniv = cz.ID_AnUniv
                WHERE 
                   (@an = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(cz.DenumireAnUniv, CHAR(9), '')))) = @an)
                   AND
                   (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) = @fac)
                   AND
                   (@allSpecs = 1 OR 
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                        CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                             THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                             ELSE sf.DenumireSpecializare END, 
                    'Ș', 'S'), 'Ț', 'T')))) IN (SELECT value FROM STRING_SPLIT(@listaSpecs, ','))
                   )
                ORDER BY ppm.NumeIntreg";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@allSpecs", toateSpecializarile ? 1 : 0);
                cmd.Parameters.AddWithValue("@listaSpecs", specializari ?? "");

                using var reader = cmd.ExecuteReader();
                while (reader.Read()) { lista.Add(reader["NumeIntreg"].ToString()!); }
            }
            return Ok(lista);
        }

        // =========================================================================
        // 2. RAPORTUL PRINCIPAL (NORMA) - Păstrat intact
        // =========================================================================

        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string anUniv, string facultate, string specializari, string profesor)
        {
            var result = new List<object>();

            // AICI E SCHIMBAREA: Definim numărul de săptămâni ca variabilă
            // Dacă profesoara zice că sunt 15, modifici doar aici.
            int nrSaptamani = 14;

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH BaseData AS (
                    SELECT 
                        ppm.NumeIntreg,
                        sf.DenumireMaterie,
                        ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat') as TipPost, 
                        ISNULL(sf.Nr_Ore_Curs, 0) + ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) as TotalOre,
                        
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                            CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                 THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                 ELSE sf.DenumireSpecializare END, 
                        'Ș', 'S'), 'Ț', 'T')))) AS SpecializareCurata,

                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) AS FacultateCurata,
                        
                        UPPER(LTRIM(RTRIM(REPLACE(cz.DenumireAnUniv, CHAR(9), '')))) AS AnCurat

                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                    INNER JOIN (
                        SELECT DISTINCT ID_AnUniv, DenumireAnUniv FROM [agsis_dw].[dbo].[Cazare]
                    ) cz ON ppm.ID_AnUniv = cz.ID_AnUniv
                )
                SELECT 
                    NumeIntreg as Profesor,
                    SpecializareCurata as Specializare,
                    DenumireMaterie as Materie,
                    TipPost,
                    SUM(TotalOre) as NormaSaptamana,
                    -- AICI FOLOSIM PARAMETRUL @saptamani
                    SUM(TotalOre * @saptamani) as NormaSemestru 
                FROM BaseData
                WHERE 
                    (@an = 'Toti' OR AnCurat = @an)
                    AND
                    (@fac = 'Toti' OR FacultateCurata = @fac)
                    AND
                    (@prof = 'Toti' OR NumeIntreg = @prof)
                    AND
                    (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost
                ORDER BY NumeIntreg, Materie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                string specsParam = string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari;
                cmd.Parameters.AddWithValue("@specs", specsParam);

                // Trimitem parametrul în SQL
                cmd.Parameters.AddWithValue("@saptamani", nrSaptamani);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        Profesor = reader["Profesor"],
                        Specializare = reader["Specializare"],
                        Materie = reader["Materie"],
                        Status = reader["TipPost"],
                        NormaSapt = reader["NormaSaptamana"],
                        NormaSem = reader["NormaSemestru"]
                    });
                }
            }
            return Ok(result);
        }

        // =========================================================================
        // 3. RAPORT SECUNDAR (STAT FUNCTII) - Păstrat intact
        // =========================================================================

        [HttpGet("stat-functii-multi")]
        public ActionResult GetStatFunctiiMulti(string specializari, string profesor)
        {
            var listaResult = new List<object>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
                WITH DateNormalizate AS (
                    SELECT 
                        sf.DenTitularSauSuplinitor,
                        sf.NrOreConventionale,
                        ppm.NumeIntreg,
                        UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                            CASE 
                                WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                ELSE sf.DenumireSpecializare 
                            END, 
                        'Ș', 'S'), 'Ț', 'T')))) AS SpecializareCurata
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                )
                SELECT 
                    ISNULL(DenTitularSauSuplinitor, 'Altele') as TipPost, 
                    COUNT(*) as TitularOcupate, 
                    SUM(CAST(ISNULL(NrOreConventionale, 0) AS INT)) as PlataCuOra
                FROM DateNormalizate
                WHERE 
                    (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                    AND 
                    (@prof = 'Toti' OR NumeIntreg = @prof)
                GROUP BY DenTitularSauSuplinitor";

                SqlCommand cmd = new SqlCommand(sql, conn);
                string specsParam = string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari;

                cmd.Parameters.AddWithValue("@specs", specsParam);
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");

                conn.Open();
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    listaResult.Add(new
                    {
                        TipPost = reader["TipPost"].ToString(),
                        TitularOcupate = Convert.ToInt32(reader["TitularOcupate"]),
                        PlataCuOra = Convert.ToInt32(reader["PlataCuOra"])
                    });
                }
            }
            return Ok(listaResult);
        }
    }
}