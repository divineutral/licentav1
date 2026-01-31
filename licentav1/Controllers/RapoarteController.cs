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

        // FUNCȚIA CRITICĂ: Normalizează inputul pentru a preveni erori și a unifica datele
        private string NormalizeazaInput(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            // 1. Elimină tot ce este după '+' (ex: "Informatica+ID" -> "Informatica")
            int plusIndex = input.IndexOf('+');
            if (plusIndex > 0) input = input.Substring(0, plusIndex);

            // 2. Elimină spații, diacritice și face majuscule
            return input.Trim().ToUpper()
                .Replace("Ș", "S").Replace("Ț", "T")
                .Replace("Ă", "A").Replace("Â", "A").Replace("Î", "I")
                .Replace("CHAR(9)", "");
        }

        [HttpGet("liste/facultati")]
        public ActionResult GetFacultati()
        {
            // Cache 60 minute pentru viteză
            return Ok(_cache.GetOrCreate("ListaFacultati", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<string> { "Toti" };
                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT DISTINCT Denumire FROM [agsis_dw].[dbo].[facultate] 
                                   WHERE Denumire LIKE 'Facultatea de%' OR Denumire IN ('FACULTATEA DE DREPT', 'FACULTATEA DE MEDICINĂ', 'FACULTATEA DE MUZICA')
                                   ORDER BY Denumire ASC";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    conn.Open();
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read()) { lista.Add(reader["Denumire"].ToString()!); }
                }
                return lista;
            }));
        }

        [HttpGet("liste/specializari-per-facultate")]
        public ActionResult GetSpecializari(string numeFacultate)
        {
            var listaRaw = new List<string>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                // SQL curăță denumirile direct la sursă pentru Dropdown
                string sql = @"SELECT DISTINCT UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(
                                CASE 
                                    WHEN CHARINDEX('+', DenumireSpecializare) > 0 
                                    THEN LEFT(DenumireSpecializare, CHARINDEX('+', DenumireSpecializare) - 1)
                                    ELSE DenumireSpecializare 
                                END, 
                                CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) as DenumireCurata
                               FROM [agsis_dw].[dbo].[View_Student_Bursa_AnUniv_Semestru]
                               WHERE (DenumireFacultate = @fac OR @fac = 'Toti') AND DenumireSpecializare IS NOT NULL";

                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@fac", numeFacultate);
                conn.Open();
                using var reader = cmd.ExecuteReader();
                while (reader.Read()) { listaRaw.Add(reader["DenumireCurata"].ToString()!); }
            }
            var listaFinala = listaRaw.Distinct().OrderBy(s => s).ToList();
            listaFinala.Insert(0, "Toti");
            return Ok(listaFinala);
        }

        [HttpGet("liste/profesori-per-specializari")]
        public ActionResult GetProfesori(string specializari)
        {
            var lista = new List<string> { "Toti" };
            if (string.IsNullOrEmpty(specializari) || specializari == "Toti") return Ok(lista);

            // Pregătim lista curată pentru SQL
            var specsList = specializari.Split(',').Select(NormalizeazaInput);
            var specsParam = string.Join(",", specsList);

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"SELECT DISTINCT ppm.NumeIntreg 
                               FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                               INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                               WHERE EXISTS (
                                   SELECT 1 FROM STRING_SPLIT(@specs, ',') s 
                                   WHERE UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                                       CASE WHEN CHARINDEX('+', sf.DenumireSpecializare) > 0 
                                            THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare) - 1)
                                            ELSE sf.DenumireSpecializare END, 
                                   'Ș', 'S'), 'Ț', 'T')))) = s.value
                               )
                               ORDER BY ppm.NumeIntreg ASC";

                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@specs", specsParam);
                conn.Open();
                using var reader = cmd.ExecuteReader();
                while (reader.Read()) { lista.Add(reader["NumeIntreg"].ToString()!); }
            }
            return Ok(lista);
        }

        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string specializari, string profesor)
        {
            var lista = new List<object>();
            if (string.IsNullOrEmpty(specializari)) return Ok(lista);

            var specsList = specializari.Split(',').Select(NormalizeazaInput);
            var specsParam = string.Join(",", specsList);

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                // CTE pentru performanță maximă și unificare date
                string sql = @"
                WITH DateNormalizate AS (
                    SELECT 
                        ppm.NumeIntreg,
                        sf.DenumireMaterie,
                        sf.Nr_Ore_Curs, sf.Nr_Ore_Seminar, sf.Nr_Ore_Laborator,
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
                    NumeIntreg AS NumeProfesor, 
                    SpecializareCurata,
                    DenumireMaterie AS NumeMaterie, 
                    SUM(ISNULL(Nr_Ore_Curs,0) + ISNULL(Nr_Ore_Seminar,0) + ISNULL(Nr_Ore_Laborator,0)) AS TotalOre
                FROM DateNormalizate
                WHERE (@specs = 'TOTI' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (NumeIntreg = @prof OR @prof = 'Toti')
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie
                ORDER BY NumeIntreg ASC";

                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@specs", specsParam);
                cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
                conn.Open();

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    lista.Add(new
                    {
                        Profesor = reader["NumeProfesor"].ToString(),
                        Specializare = reader["SpecializareCurata"].ToString(),
                        Materie = reader["NumeMaterie"].ToString(),
                        Norma = reader["TotalOre"].ToString()
                    });
                }
            }
            return Ok(lista);
        }

        [HttpGet("stat-functii-multi")]
        public ActionResult GetStatFunctiiMulti(string specializari, string profesor)
        {
            var listaResult = new List<object>();
            if (string.IsNullOrEmpty(specializari)) return Ok(listaResult);

            var specsList = specializari.Split(',').Select(NormalizeazaInput);
            var specsParam = string.Join(",", specsList);

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
                WHERE (@specs = 'TOTI' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (NumeIntreg = @prof OR @prof = 'Toti')
                GROUP BY DenTitularSauSuplinitor";

                SqlCommand cmd = new SqlCommand(sql, conn);
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

        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAni()
        {
            var lista = new List<object>();
            lista.Add(new { id = "45", nume = "AN UNIVERSITAR 2025-2026" });

            using (SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection")!))
            {
                string sql = "SELECT DISTINCT ID_AnUniv, DenumireAnUniv FROM [agsis_dw].[dbo].[Cazare] ORDER BY DenumireAnUniv DESC";
                SqlCommand cmd = new SqlCommand(sql, conn);
                conn.Open();
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    lista.Add(new { id = reader["ID_AnUniv"].ToString(), nume = reader["DenumireAnUniv"].ToString() });
                }
            }
            return Ok(lista);
        }
    }
}