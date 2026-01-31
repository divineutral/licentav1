using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using System.Data;
using System.Text.RegularExpressions;

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

        // ==========================================
        // 1. LOGICA DE CURĂȚARE CENTRALIZATĂ
        // ==========================================
        private string NormalizeazaText(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";

            // 1. Elimină tot ce e după '+' (ex: "Informatica+ID" -> "Informatica")
            int plusIndex = input.IndexOf('+');
            if (plusIndex > 0) input = input.Substring(0, plusIndex);

            // 2. Majuscule, Trim și eliminare caractere invizibile (Tab, Newline)
            input = input.ToUpper().Trim()
                         .Replace("\t", "")
                         .Replace("\n", "")
                         .Replace("\r", "");

            // 3. Uniformizare diacritice
            return input
                .Replace("Ș", "S").Replace("Ț", "T")
                .Replace("Ă", "A").Replace("Â", "A").Replace("Î", "I")
                .Replace("CHAR(9)", ""); // Siguranță pentru string-uri vechi
        }

        // ==========================================
        // 2. ENDPOINT-URI PENTRU FILTRE (CASCADĂ)
        // ==========================================

        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAni()
        {
            // Cache 60 min
            return Ok(_cache.GetOrCreate("ListaAniUniv", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<object>();

                // Opțional: Adăugăm manual anul viitor/curent dacă nu e încă în bază
                lista.Add(new { id = "AN UNIVERSITAR 2025-2026", nume = "AN UNIVERSITAR 2025-2026" });

                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    // FIX CRITIC: Returnăm "Denumire" ca ID, nu "ID_AnUniv", pentru a matcha parametrul din SSRS
                    string sql = @"SELECT DISTINCT UPPER(LTRIM(RTRIM(REPLACE(DenumireAnUniv, CHAR(9), '')))) COLLATE DATABASE_DEFAULT as AnCurat 
                                   FROM [agsis_dw].[dbo].[Cazare] 
                                   WHERE DenumireAnUniv IS NOT NULL 
                                   ORDER BY AnCurat DESC";

                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var an = reader["AnCurat"].ToString();
                        // Evităm duplicatele dacă anul manual există deja
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
                    // FIX: Luăm datele din tabela de fapte (Post_Profesor_Materie)
                    // Astfel afișăm doar facultățile care au activitate reală
                    string sql = @"
                        SELECT DISTINCT 
                            UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT as FacCurata
                        FROM [agsis_dw].[dbo].[Post_Profesor_Materie]
                        WHERE DenumireFacultate IS NOT NULL
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
                // NOTA: Am simplificat query-ul pentru a folosi doar Post_Profesor_Materie
                // deoarece este sursa sigură pentru care avem legătura cu Anul Universitar.
                string sql = @"
            SELECT DISTINCT 
                UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                    CASE WHEN CHARINDEX('+', ppm.DenumireSpecializare COLLATE DATABASE_DEFAULT) > 0 
                         THEN LEFT(ppm.DenumireSpecializare, CHARINDEX('+', ppm.DenumireSpecializare COLLATE DATABASE_DEFAULT) - 1)
                         ELSE ppm.DenumireSpecializare END, 
                'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT as SpecCurata
            FROM [agsis_dw].[dbo].[Post_Profesor_Materie] ppm
            INNER JOIN [agsis_dw].[dbo].[Cazare] cz 
                ON ppm.ID_AnUniv = cz.ID_AnUniv
            WHERE 
                ppm.DenumireSpecializare IS NOT NULL 
                AND
                -- Filtru FACULTATE
                (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT = @fac COLLATE DATABASE_DEFAULT)
                AND
                -- Filtru AN UNIVERSITAR (Cascada)
                (@an = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(cz.DenumireAnUniv, CHAR(9), '')))) = @an)
            ORDER BY SpecCurata";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@fac", NormalizeazaText(numeFacultate ?? "Toti"));
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string val = reader["SpecCurata"].ToString()!;
                    if (!string.IsNullOrWhiteSpace(val)) lista.Add(val);
                }
            }
            return Ok(lista);
        }

        [HttpGet("liste/profesori-per-specializari")]
        public ActionResult GetProfesori(string anUniv, string facultate, string specializari)
        {
            var lista = new List<string> { "Toti" };

            // Convertim lista de specializări într-un format SQL-safe
            var specs = (specializari ?? "Toti").Split(',')
                        .Select(NormalizeazaText)
                        .Where(s => !string.IsNullOrEmpty(s))
                        .ToList();
            bool toateSpecializarile = (specializari == "Toti" || !specs.Any());

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                string sql = @"
            SELECT DISTINCT ppm.NumeIntreg
            FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
            INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm 
                ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
            INNER JOIN [agsis_dw].[dbo].[Cazare] cz 
                ON ppm.ID_AnUniv = cz.ID_AnUniv
            WHERE 
                -- 1. Filtru AN
                (@an = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(cz.DenumireAnUniv, CHAR(9), '')))) = @an)
                AND
                -- 2. Filtru FACULTATE (Important când Specializarea e 'Toti')
                (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT = @fac COLLATE DATABASE_DEFAULT)
                AND
                -- 3. Filtru SPECIALIZĂRI
                (@allSpecs = 1 OR 
                 UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                    CASE WHEN CHARINDEX('+', sf.DenumireSpecializare COLLATE DATABASE_DEFAULT) > 0 
                         THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare COLLATE DATABASE_DEFAULT) - 1)
                         ELSE sf.DenumireSpecializare END, 
                'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT IN (SELECT value COLLATE DATABASE_DEFAULT FROM STRING_SPLIT(@listaSpecs, ','))
                )
            ORDER BY ppm.NumeIntreg";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
                cmd.Parameters.AddWithValue("@fac", NormalizeazaText(facultate ?? "Toti"));
                cmd.Parameters.AddWithValue("@allSpecs", toateSpecializarile ? 1 : 0);
                cmd.Parameters.AddWithValue("@listaSpecs", string.Join(",", specs));

                using var reader = cmd.ExecuteReader();
                while (reader.Read()) { lista.Add(reader["NumeIntreg"].ToString()!); }
            }
            return Ok(lista);
        }

        // ==========================================
        // 3. RAPORT PRINCIPAL (NORMA)
        // ==========================================
        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string anUniv, string facultate, string specializari, string profesor)
        {
            var result = new List<object>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                // 1. CONSTRUIRE FILTRE INTERIOARE (Se aplică pe tabelele sursă cz și ppm)
                string whereInner = "WHERE 1=1";

                if (!string.IsNullOrEmpty(anUniv) && anUniv != "Toti")
                {
                    whereInner += " AND UPPER(LTRIM(RTRIM(REPLACE(cz.DenumireAnUniv, CHAR(9), '')))) = @anUniv";
                }

                if (!string.IsNullOrEmpty(facultate) && facultate != "Toti")
                {
                    whereInner += " AND UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ppm.DenumireFacultate, CHAR(9), ''), 'Ș', 'S'), 'Ț', 'T')))) = @facultate";
                }

                if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                {
                    whereInner += " AND ppm.NumeIntreg = @profesor";
                }

                // 2. CONSTRUIRE FILTRE EXTERIOARE (Se aplică pe coloana calculată SpecializareCurata)
                string whereOuter = "WHERE 1=1";

                if (!string.IsNullOrEmpty(specializari) && specializari != "Toti")
                {
                    whereOuter += " AND SpecializareCurata IN (SELECT value COLLATE DATABASE_DEFAULT FROM STRING_SPLIT(@specs, ','))";
                }

                string sql = $@"
        WITH BaseData AS (
            SELECT 
                ppm.NumeIntreg,
                sf.DenumireMaterie,
                ISNULL(sf.Nr_Ore_Curs, 0) + ISNULL(sf.Nr_Ore_Seminar, 0) + ISNULL(sf.Nr_Ore_Laborator, 0) as TotalOre,
                -- Curățare Specializare (Logică Complexă)
                UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                    CASE WHEN CHARINDEX('+', sf.DenumireSpecializare COLLATE DATABASE_DEFAULT) > 0 
                         THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare COLLATE DATABASE_DEFAULT) - 1)
                         ELSE sf.DenumireSpecializare END, 
                'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT AS SpecializareCurata
            FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
            INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm 
                ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
            INNER JOIN [agsis_dw].[dbo].[Cazare] cz 
                ON ppm.ID_AnUniv = cz.ID_AnUniv
            
            {whereInner} -- <--- AICI SE PUN FILTRELE PENTRU AN, FACULTATE, PROFESOR
        )
        SELECT 
            NumeIntreg as Profesor,
            SpecializareCurata as Specializare,
            DenumireMaterie as Materie,
            SUM(TotalOre) as Norma
        FROM BaseData
        {whereOuter} -- <--- AICI SE PUNE FILTRUL PENTRU SPECIALIZARE
        GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie
        ORDER BY NumeIntreg, Materie";

                conn.Open();
                using var cmd = new SqlCommand(sql, conn);

                // Adăugare parametri
                if (!string.IsNullOrEmpty(anUniv) && anUniv != "Toti")
                    cmd.Parameters.AddWithValue("@anUniv", anUniv);

                if (!string.IsNullOrEmpty(facultate) && facultate != "Toti")
                    cmd.Parameters.AddWithValue("@facultate", NormalizeazaText(facultate));

                if (!string.IsNullOrEmpty(profesor) && profesor != "Toti")
                    cmd.Parameters.AddWithValue("@profesor", profesor);

                if (!string.IsNullOrEmpty(specializari) && specializari != "Toti")
                    cmd.Parameters.AddWithValue("@specs", specializari);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        Profesor = reader["Profesor"],
                        Specializare = reader["Specializare"],
                        Materie = reader["Materie"],
                        Norma = reader["Norma"]
                    });
                }
            }
            return Ok(result);
        }

        // ==========================================
        // 4. RAPORT SECUNDAR (STAT FUNCTII)
        // ==========================================
        [HttpGet("stat-functii-multi")]
        public ActionResult GetStatFunctiiMulti(string specializari, string profesor)
        {
            var listaResult = new List<object>();
            if (string.IsNullOrEmpty(specializari)) return Ok(listaResult);

            var specsList = specializari.Split(',').Select(NormalizeazaText);
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
                                WHEN CHARINDEX('+', sf.DenumireSpecializare COLLATE DATABASE_DEFAULT) > 0 
                                THEN LEFT(sf.DenumireSpecializare, CHARINDEX('+', sf.DenumireSpecializare COLLATE DATABASE_DEFAULT) - 1)
                                ELSE sf.DenumireSpecializare 
                            END, 
                        'Ș', 'S'), 'Ț', 'T')))) COLLATE DATABASE_DEFAULT AS SpecializareCurata
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    INNER JOIN [agsis_dw].[dbo].[Post_Profesor_Materie] ppm ON sf.ID_Post_Profesor_Materie = ppm.ID_Post_Profesor_Materie
                )
                SELECT 
                    ISNULL(DenTitularSauSuplinitor, 'Altele') as TipPost, 
                    COUNT(*) as TitularOcupate, 
                    SUM(CAST(ISNULL(NrOreConventionale, 0) AS INT)) as PlataCuOra
                FROM DateNormalizate
                WHERE (@specs = 'TOTI' OR SpecializareCurata IN (SELECT value COLLATE DATABASE_DEFAULT FROM STRING_SPLIT(@specs, ',')))
                  AND (NumeIntreg COLLATE DATABASE_DEFAULT = @prof COLLATE DATABASE_DEFAULT OR @prof = 'Toti')
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
    }
}