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
        private const string BrandColorHex = "#56723e";

        private readonly string[] DomeniiExcel = new string[]
        {
            "Matematica", "Informatica", "Fizica", "Chimie si inginerie chimica", "Stiintele pamantului si atmosferei",
            "Inginerie civila", "Inginerie electrica, electronica si telecomunicatii", "Inginerie geologica, mine, petrol si gaze",
            "Ingineria transporturilor", "Ingineria resurselor vegetale si animale",
            "Ingineria sistemelor, calculatoare si tehnologia informatiei",
            "Inginerie mecanica, mecatronica, inginerie industriala si management",
            "Biologie", "Biochimie", "Medicina", "Medicina veterinara", "Medicina dentara", "Farmacie",
            "Stiinte juridice", "Stiinte administrative", "Stiinte ale comunicarii", "Sociologie",
            "Stiinte politice", "Stiinte militare, informatii si ordine publica",
            "Stiinte economice (doar Cibernetica, statistica si informatica economica)",
            "Stiinte economice (fara Cibernetica, statistica si informatica economica)",
            "Psihologie si stiinte comportamentale",
            "Filologie", "Filosofie", "Istorie", "Teologie", "Studii culturale",
            "Arhitectura si urbanism", "Arte vizuale (fara Istoria si teoria artei)",
            "Arte vizuale (doar Istoria si teoria artei)", "Teatru si artele spectacolului",
            "Cinematografie si media", "Muzica (doar Interpretare muzicala)",
            "Muzica (fara Interpretare muzicala)", "Stiintele Sportului si Educatiei Fizice"
        };

        private static readonly Dictionary<int, int> MappingMetaspec = new Dictionary<int, int>
        {
            { 20, 5 }, { 34, 5 }, { 44, 5 }, { 182, 5 }, { 837, 5 }, { 847, 5 }, { 848, 5 },
            { 43, 9 }, { 148, 8 }, { 171, 11 }, { 306, 11 }, { 358, 11 }, { 464, 11 },
            { 466, 11 }, { 615, 11 }, { 823, 11 }, { 828, 11 },
            { 35, 11 }, { 82, 7 }, { 84, 7 }, { 139, 11 }, { 162, 7 }, { 310, 7 },
            { 315, 7 }, { 316, 7 }, { 362, 7 }, { 372, 7 }, { 418, 7 }, { 597, 7 }, { 617, 7 },
            { 116, 12 }, { 126, 11 }, { 129, 11 }, { 156, 11 }, { 228, 11 }, { 229, 12 },
            { 404, 11 }, { 529, 11 }, { 142, 11 }, { 326, 11 }, { 531, 11 }, { 806, 11 }, { 819, 11 },
            { 53, 11 }, { 138, 11 }, { 446, 11 }, { 448, 11 }, { 449, 11 }, { 450, 11 },
            { 451, 11 }, { 496, 11 }, { 497, 11 }, { 821, 11 },
            { 46, 9 }, { 122, 9 }, { 176, 9 }, { 178, 9 }, { 731, 9 },
            { 72, 9 }, { 90, 9 }, { 118, 9 }, { 226, 9 }, { 235, 9 }, { 249, 9 },
            { 307, 9 }, { 437, 9 }, { 458, 9 }, { 566, 9 }, { 845, 9 },
            { 101, 40 }, { 102, 40 }, { 104, 40 }, { 251, 1 }, { 317, 40 },
            { 340, 1 }, { 368, 40 }, { 369, 40 }, { 477, 40 },
            { 45, 25 }, { 73, 25 }, { 93, 25 }, { 112, 24 }, { 221, 25 }, { 223, 25 },
            { 227, 25 }, { 242, 25 }, { 283, 25 }, { 288, 25 }, { 299, 25 },
            { 181, 31 }, { 196, 27 }, { 197, 27 }, { 200, 27 }, { 205, 27 }, { 207, 27 },
            { 209, 27 }, { 217, 27 }, { 218, 27 }, { 341, 27 }, { 343, 27 }, { 463, 27 },
            { 511, 27 }, { 512, 27 }, { 513, 27 }, { 514, 27 }, { 726, 27 }, { 798, 27 }, { 801, 27 },
            { 60, 18 }, { 64, 18 }, { 331, 18 }, { 485, 18 }, { 555, 18 },
            { 41, 20 }, { 98, 20 }, { 100, 25 }, { 322, 21 }, { 524, 25 }, { 579, 20 }, { 831, 20 },
            { 276, 26 }, { 294, 26 }, { 296, 26 }, { 383, 26 }, { 384, 26 }, { 416, 26 },
            { 515, 26 }, { 813, 26 }, { 851, 26 }, { 832, 26 }, { 834, 26 },
            { 394, 14 }, { 397, 14 }, { 402, 14 }, { 484, 14 }, { 585, 14 }, { 594, 14 }, { 835, 14 },
            { 186, 37 }, { 187, 37 }, { 264, 37 }, { 332, 37 }, { 351, 37 }, { 557, 37 }, { 838, 37 },
            { 78, 39 }, { 189, 39 }, { 325, 39 }, { 470, 39 }, { 783, 39 }, { 784, 39 }, { 846, 39 },
        };

        private static readonly Dictionary<int, int> AnsIdToCol = new Dictionary<int, int>
        {
            { 1, 10 }, { 2, 11 }, { 3, 12 }, { 4, 13 }, { 5, 14 },
            { 6, 15 }, { 7, 16 }, { 8, 17 }, { 9, 18 }, { 10, 19 }, { 11, 20 }, { 12, 21 },
            { 13, 22 }, { 14, 23 }, { 15, 24 }, { 16, 25 }, { 17, 26 }, { 18, 27 },
            { 19, 28 }, { 20, 29 }, { 21, 30 }, { 22, 31 }, { 23, 32 }, { 24, 33 },
            { 25, 34 }, { 26, 35 }, { 27, 36 },
            { 28, 37 }, { 29, 38 }, { 30, 39 }, { 31, 40 }, { 32, 41 }, { 33, 42 },
            { 34, 43 }, { 35, 44 }, { 36, 45 }, { 37, 46 }, { 38, 47 }, { 39, 48 }, { 40, 49 },
        };

        // =====================================================================
        // MAPARE STATICA: ID_Catedra -> DenumireCatedra
        // Sursa: agsis_dw.dbo.Post_Profesor_Materie (date istorice stabile)
        // ID-urile de catedra sunt stabile intre ani universitari
        // =====================================================================
        private static readonly Dictionary<long, string> MapareCatedra = new Dictionary<long, string>
        {
            { 583, "DEPARTAMENTUL AUTOMATICA SI TEHNOLOGIA INFORMATIEI" },
            { 554, "DEPARTAMENTUL AUTOMATICA SI TEHNOLOGIA INFORMATIEI" },
            { 524, "DEPARTAMENTUL AUTOMATICA SI TEHNOLOGIA INFORMATIEI" },
            { 572, "DEPARTAMENTUL AUTOVEHICULE SI TRANSPORTURI" },
            { 584, "DEPARTAMENTUL AUTOVEHICULE SI TRANSPORTURI" },
            { 542, "DEPARTAMENTUL AUTOVEHICULE SI TRANSPORTURI" },
            { 570, "DEPARTAMENTUL DESIGN DE PRODUS, MECATRONIC SI MEDIU" },
            { 585, "DEPARTAMENTUL DESIGN DE PRODUS, MECATRONIC SI MEDIU" },
            { 540, "DEPARTAMENTUL DESIGN DE PRODUS, MECATRONIC SI MEDIU" },
            { 566, "DEPARTAMENTUL DISCIPLINELOR FUNDAMENTALE, PROFILACTICE SI CLINICE" },
            { 586, "DEPARTAMENTUL DISCIPLINELOR FUNDAMENTALE, PROFILACTICE SI CLINICE" },
            { 536, "DEPARTAMENTUL DISCIPLINELOR FUNDAMENTALE, PROFILACTICE SI CLINICE" },
            { 532, "DEPARTAMENTUL DREPT" },
            { 562, "DEPARTAMENTUL DREPT" },
            { 587, "DEPARTAMENTUL DREPT" },
            { 577, "DEPARTAMENTUL EDUCATIE FIZICA SI MOTRICITATE SPECIALA" },
            { 588, "DEPARTAMENTUL EDUCATIE FIZICA SI MOTRICITATE SPECIALA" },
            { 547, "DEPARTAMENTUL EDUCATIE FIZICA SI MOTRICITATE SPECIALA" },
            { 575, "DEPARTAMENTUL ELECTRONICA SI CALCULATOARE" },
            { 545, "DEPARTAMENTUL ELECTRONICA SI CALCULATOARE" },
            { 589, "DEPARTAMENTUL ELECTRONICA SI CALCULATOARE" },
            { 557, "DEPARTAMENTUL EXPLOATARI FORESTIERE, AMENAJAREA PADURILOR SI MASURATORI TERESTRE" },
            { 527, "DEPARTAMENTUL EXPLOATARI FORESTIERE, AMENAJAREA PADURILOR SI MASURATORI TERESTRE" },
            { 590, "DEPARTAMENTUL EXPLOATARI FORESTIERE, AMENAJAREA PADURILOR SI MASURATORI TERESTRE" },
            { 531, "DEPARTAMENTUL FINANTE, CONTABILITATE SI TEORIE ECONOMICA" },
            { 591, "DEPARTAMENTUL FINANTE, CONTABILITATE SI TEORIE ECONOMICA" },
            { 561, "DEPARTAMENTUL FINANTE, CONTABILITATE SI TEORIE ECONOMICA" },
            { 573, "DEPARTAMENTUL INGINERIA FABRICATIEI" },
            { 592, "DEPARTAMENTUL INGINERIA FABRICATIEI" },
            { 543, "DEPARTAMENTUL INGINERIA FABRICATIEI" },
            { 522, "DEPARTAMENTUL INGINERIA MATERIALELOR SI SUDURA" },
            { 593, "DEPARTAMENTUL INGINERIA MATERIALELOR SI SUDURA" },
            { 552, "DEPARTAMENTUL INGINERIA MATERIALELOR SI SUDURA" },
            { 529, "DEPARTAMENTUL INGINERIA SI MANAGEMENTUL ALIMENTATIEI SI TURISMULUI" },
            { 594, "DEPARTAMENTUL INGINERIA SI MANAGEMENTUL ALIMENTATIEI SI TURISMULUI" },
            { 559, "DEPARTAMENTUL INGINERIA SI MANAGEMENTUL ALIMENTATIEI SI TURISMULUI" },
            { 558, "DEPARTAMENTUL INGINERIE CIVILA" },
            { 595, "DEPARTAMENTUL INGINERIE CIVILA" },
            { 528, "DEPARTAMENTUL INGINERIE CIVILA" },
            { 523, "DEPARTAMENTUL INGINERIE ELECTRICA SI FIZICA APLICATA" },
            { 553, "DEPARTAMENTUL INGINERIE ELECTRICA SI FIZICA APLICATA" },
            { 596, "DEPARTAMENTUL INGINERIE ELECTRICA SI FIZICA APLICATA" },
            { 597, "DEPARTAMENTUL INGINERIE MECANICA" },
            { 550, "DEPARTAMENTUL INGINERIE MECANICA" },
            { 520, "DEPARTAMENTUL INGINERIE MECANICA" },
            { 521, "DEPARTAMENTUL INGINERIE SI MANAGEMENT INDUSTRIAL" },
            { 598, "DEPARTAMENTUL INGINERIE SI MANAGEMENT INDUSTRIAL" },
            { 551, "DEPARTAMENTUL INGINERIE SI MANAGEMENT INDUSTRIAL" },
            { 576, "DEPARTAMENTUL INSTALATII PENTRU CONSTRUCTII" },
            { 546, "DEPARTAMENTUL INSTALATII PENTRU CONSTRUCTII" },
            { 599, "DEPARTAMENTUL INSTALATII PENTRU CONSTRUCTII" },
            { 568, "DEPARTAMENTUL INTERPRETARE SI PEDAGOGIE MUZICALA" },
            { 538, "DEPARTAMENTUL INTERPRETARE SI PEDAGOGIE MUZICALA" },
            { 600, "DEPARTAMENTUL INTERPRETARE SI PEDAGOGIE MUZICALA" },
            { 534, "DEPARTAMENTUL LINGVISTICA TEORETICA SI APLICATA" },
            { 601, "DEPARTAMENTUL LINGVISTICA TEORETICA SI APLICATA" },
            { 564, "DEPARTAMENTUL LINGVISTICA TEORETICA SI APLICATA" },
            { 602, "DEPARTAMENTUL LITERATURA SI STUDII CULTURALE" },
            { 533, "DEPARTAMENTUL LITERATURA SI STUDII CULTURALE" },
            { 563, "DEPARTAMENTUL LITERATURA SI STUDII CULTURALE" },
            { 548, "DEPARTAMENTUL MANAGEMENT SI INFORMATICA ECONOMICA" },
            { 603, "DEPARTAMENTUL MANAGEMENT SI INFORMATICA ECONOMICA" },
            { 578, "DEPARTAMENTUL MANAGEMENT SI INFORMATICA ECONOMICA" },
            { 579, "DEPARTAMENTUL MARKETING,TURISM-SERVICII SI AFACERI INTERNATIONALE" },
            { 604, "DEPARTAMENTUL MARKETING,TURISM-SERVICII SI AFACERI INTERNATIONALE" },
            { 549, "DEPARTAMENTUL MARKETING,TURISM-SERVICII SI AFACERI INTERNATIONALE" },
            { 605, "DEPARTAMENTUL MATEMATICA SI INFORMATICA" },
            { 560, "DEPARTAMENTUL MATEMATICA SI INFORMATICA" },
            { 530, "DEPARTAMENTUL MATEMATICA SI INFORMATICA" },
            { 565, "DEPARTAMENTUL PERFORMANTA MOTRICA" },
            { 535, "DEPARTAMENTUL PERFORMANTA MOTRICA" },
            { 606, "DEPARTAMENTUL PERFORMANTA MOTRICA" },
            { 555, "DEPARTAMENTUL PRELUCRAREA LEMNULUI SI DESIGNUL PRODUSELOR DIN LEMN" },
            { 525, "DEPARTAMENTUL PRELUCRAREA LEMNULUI SI DESIGNUL PRODUSELOR DIN LEMN" },
            { 607, "DEPARTAMENTUL PRELUCRAREA LEMNULUI SI DESIGNUL PRODUSELOR DIN LEMN" },
            { 539, "Departamentul Psihologie si Stiinte ale Educatie" },
            { 608, "DEPARTAMENTUL PSIHOLOGIE SI STIINTELE EDUCATIEI" },
            { 569, "DEPARTAMENTUL PSIHOLOGIE SI STIINTELE EDUCATIEI" },
            { 526, "DEPARTAMENTUL SILVICULTURA" },
            { 609, "DEPARTAMENTUL SILVICULTURA" },
            { 556, "DEPARTAMENTUL SILVICULTURA" },
            { 610, "DEPARTAMENTUL SPECIALITATILOR MEDICALE SI CHIRURGICALE" },
            { 567, "DEPARTAMENTUL SPECIALITATILOR MEDICALE SI CHIRURGICALE" },
            { 537, "DEPARTAMENTUL SPECIALITATILOR MEDICALE SI CHIRURGICALE" },
            { 574, "DEPARTAMENTUL STIINTA MATERIALELOR" },
            { 544, "DEPARTAMENTUL STIINTA MATERIALELOR" },
            { 611, "DEPARTAMENTUL STIINTA MATERIALELOR" },
            { 612, "DEPARTAMENTUL STIINTE SOCIALE SI ALE COMUNICARII" },
            { 541, "DEPARTAMENTUL STIINTE SOCIALE SI ALE COMUNICARII" },
            { 571, "DEPARTAMENTUL STIINTE SOCIALE SI ALE COMUNICARII" },
            { 862, "DPPD" },
            { 613, "DPPD" },
            { 863, "DPPD" },
            // ID-uri noi din anul 45 - vor fi rezolvate prin DenumireFacultate din view
            { 614, "DEPARTAMENTUL AUTOMATICA SI TEHNOLOGIA INFORMATIEI" },
            { 615, "DEPARTAMENTUL AUTOVEHICULE SI TRANSPORTURI" },
            { 616, "DEPARTAMENTUL DESIGN DE PRODUS, MECATRONIC SI MEDIU" },
            { 617, "DEPARTAMENTUL DISCIPLINELOR FUNDAMENTALE, PROFILACTICE SI CLINICE" },
            { 618, "DEPARTAMENTUL DREPT" },
            { 619, "DEPARTAMENTUL EDUCATIE FIZICA SI MOTRICITATE SPECIALA" },
            { 620, "DEPARTAMENTUL ELECTRONICA SI CALCULATOARE" },
            { 621, "DEPARTAMENTUL EXPLOATARI FORESTIERE, AMENAJAREA PADURILOR SI MASURATORI TERESTRE" },
            { 622, "DEPARTAMENTUL FINANTE, CONTABILITATE SI TEORIE ECONOMICA" },
            { 623, "DEPARTAMENTUL INGINERIA FABRICATIEI" },
            { 624, "DEPARTAMENTUL INGINERIA MATERIALELOR SI SUDURA" },
            { 625, "DEPARTAMENTUL INGINERIA SI MANAGEMENTUL ALIMENTATIEI SI TURISMULUI" },
            { 626, "DEPARTAMENTUL INGINERIE CIVILA" },
            { 627, "DEPARTAMENTUL INGINERIE ELECTRICA SI FIZICA APLICATA" },
            { 628, "DEPARTAMENTUL INGINERIE MECANICA" },
            { 629, "DEPARTAMENTUL INGINERIE SI MANAGEMENT INDUSTRIAL" },
            { 630, "DEPARTAMENTUL INSTALATII PENTRU CONSTRUCTII" },
            { 631, "DEPARTAMENTUL INTERPRETARE SI PEDAGOGIE MUZICALA" },
            { 632, "DEPARTAMENTUL LINGVISTICA TEORETICA SI APLICATA" },
            { 633, "DEPARTAMENTUL LITERATURA SI STUDII CULTURALE" },
            { 634, "DEPARTAMENTUL MANAGEMENT SI INFORMATICA ECONOMICA" },
            { 635, "DEPARTAMENTUL MARKETING,TURISM-SERVICII SI AFACERI INTERNATIONALE" },
            { 636, "DEPARTAMENTUL MATEMATICA SI INFORMATICA" },
            { 637, "DEPARTAMENTUL PERFORMANTA MOTRICA" },
            { 638, "DEPARTAMENTUL PRELUCRAREA LEMNULUI SI DESIGNUL PRODUSELOR DIN LEMN" },
            { 639, "Departamentul Psihologie si Stiinte ale Educatie" },
            { 640, "DEPARTAMENTUL SILVICULTURA" },
            { 641, "DEPARTAMENTUL SPECIALITATILOR MEDICALE SI CHIRURGICALE" },
            { 642, "DEPARTAMENTUL STIINTA MATERIALELOR" },
            { 643, "DEPARTAMENTUL STIINTE SOCIALE SI ALE COMUNICARII" },
            { 644, "DPPD" },
        };

        // =====================================================================
        // CORECȚII NUME CORUPTE (caractere ? din encoding greșit în AGSIS)
        // Sursa: interogare directă DB + corectare manuală
        // =====================================================================
        private static readonly Dictionary<int, string> NumeCorecte = new Dictionary<int, string>
        {
            { 6621, "ȘIREIU RAMONA DANIELA" },
            { 6698, "BENEA ALINA-PETRUȚA" },
            { 6631, "CĂBAȘ NICOLAE SERGIU" },
            { 4605, "CĂPRIȚĂ FLORIN" },
            { 4165, "CIORICEANU IONUȚ-HORIA" },
            { 2375, "ILAȘ MAGDALENA" },
            { 6616, "MÂNDRU MARIA SPERANȚA" },
            { 16893, "MILEA GIGUȘA-ROXANA" },
            { 16894, "NEGOIȚĂ ELENA" },
            { 6803, "STAREȘU CAMELIA MARIANA" },
            { 6800, "ȘERBAN AGURIȚA DORINELA" },
            { 5881, "TUCHEL IONUȚ-VLAD" },
            // Altii din DW cu ?
            { 2899, "BREZEANU ALIN IONUȚ" },
            { 4345, "DIACONU ȘTEFANIA-ROXANA" },
            { 4352, "MANIȘIU(VASILE) VIRGINIA IOANA" },
            { 4401, "FOLEA(VECERDI) CRISTINA AGNEȘ" },
            { 4821, "CHIRA CODRUȚA-ELENA" },
            { 5833, "MARCHIȘ(TOMA) MARIA-ALEXANDRA" },
            { 5884, "VEZETEU COSMIN-DĂNUȚ" },
            { 6716, "BĂȘEANU IONUȚ-CRISTIAN-COZMIN" },
            { 6721, "MÂNZĂȚANU DIANA" },
            { 6761, "CIOPLEIAȘ BOGDAN-NICOLAE" },
        };

        private string FixNume(string? nume, object? idProfesorObj)
        {
            if (idProfesorObj != null && idProfesorObj != DBNull.Value)
            {
                int id = Convert.ToInt32(idProfesorObj);
                if (NumeCorecte.TryGetValue(id, out var corect)) return corect;
            }
            return nume ?? "";
        }

        public RapoarteController(IConfiguration configuration, IMemoryCache cache)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection")!;
            _cache = cache;
        }

        private string GetDenumireCatedra(long idCatedra)
        {
            return MapareCatedra.TryGetValue(idCatedra, out var den) ? den : $"Catedra {idCatedra}";
        }

        #region ================= LISTE (DROPDOWNS) =================

        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAni()
        {
            return Ok(_cache.GetOrCreate("ListaAniUniv", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<object>();
                using var conn = new SqlConnection(_connectionString);
                string sql = @"
                    SELECT UPPER(LTRIM(RTRIM(Denumire))) COLLATE DATABASE_DEFAULT AS AnCurat 
                    FROM [AGSIS].[dbo].[AnUniversitar] 
                    WHERE Denumire IS NOT NULL 
                    ORDER BY Ordine DESC";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var an = reader["AnCurat"]?.ToString() ?? "";
                    lista.Add(new { id = an, nume = an });
                }
                return lista;
            }));
        }

        [HttpGet("liste/facultati")]
        public ActionResult GetFacultati()
        {
            return Ok(_cache.GetOrCreate("ListaFacultati_v3", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<string> { "Toti" };
                using var conn = new SqlConnection(_connectionString);
                // Returnam DenumireFacultate EXACT din DB (cu diacritice, fara UPPER)
                // pentru ca filtrarea la departamente foloseste COLLATE CI_AI
                string sql = @"
                    SELECT DISTINCT vcm.DenumireFacultate
                    FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                    WHERE vcm.DenumireFacultate IS NOT NULL
                      AND LTRIM(RTRIM(vcm.DenumireFacultate)) != ''
                    ORDER BY vcm.DenumireFacultate ASC";
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var val = reader["DenumireFacultate"]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(val)) lista.Add(val);
                }
                return lista;
            }));
        }

        [HttpGet("liste/departamente")]
        public ActionResult GetDepartamente(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            // IMPORTANT: ID_Catedra apare la MULTIPLE facultati in acelasi an.
            // Solutia corecta: filtram randurile din view dupa DenumireFacultate exact (COLLATE),
            // apoi returnam DenumireCatedra din maparea statica, deduplicate.
            // Compararea facultatii: ignoram case si diacritice prin COLLATE Latin1_General_CI_AI
            string sql = @"
                SELECT DISTINCT vcm.StatDeFunctiiID_Catedra
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                WHERE vcm.StatDeFunctiiID_Catedra IS NOT NULL
                  AND (
                       @fac = 'Toti'
                    OR vcm.DenumireFacultate COLLATE Latin1_General_CI_AI = @fac COLLATE Latin1_General_CI_AI
                  )
                ORDER BY vcm.StatDeFunctiiID_Catedra";
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            var idSet = new HashSet<string>();
            while (reader.Read())
            {
                if (reader["StatDeFunctiiID_Catedra"] != DBNull.Value)
                {
                    long idCat = Convert.ToInt64(reader["StatDeFunctiiID_Catedra"]);
                    var den = GetDenumireCatedra(idCat);
                    if (idSet.Add(den)) lista.Add(den);
                }
            }
            lista.Sort((a, b) => a == "Toti" ? -1 : b == "Toti" ? 1 : string.Compare(a, b));
            return Ok(lista);
        }

        [HttpGet("liste/specializari-per-facultate")]
        public ActionResult GetSpecializari(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            string sql = @"
                SELECT DISTINCT 
                    UPPER(LTRIM(RTRIM(
                        REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                            CASE WHEN CHARINDEX('+', vcm.DenumireSpecializare) > 0 
                                 THEN LEFT(vcm.DenumireSpecializare, CHARINDEX('+', vcm.DenumireSpecializare) - 1)
                                 ELSE vcm.DenumireSpecializare END,
                        ' - CORECT', ''), ' CORECT', ''), ' - COPIE', ''), 'S', 'S'), 'T', 'T')
                    ))) COLLATE DATABASE_DEFAULT AS SpecCurata
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                WHERE vcm.DenumireSpecializare IS NOT NULL
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(vcm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) = @fac)
                ORDER BY SpecCurata";
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var val = reader["SpecCurata"]?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(val) && !lista.Contains(val)) lista.Add(val);
            }
            return Ok(lista);
        }

        [HttpGet("liste/profesori-per-specializari")]
        public ActionResult GetProfesori(string? anUniv, string? facultate, string? specializari, string? departament)
        {
            var lista = new List<string> { "Toti" };
            bool toateSpec = string.IsNullOrEmpty(specializari) || specializari == "Toti";
            using var conn = new SqlConnection(_connectionString);
            string sql = @"
                SELECT DISTINCT vcm.NumeIntregProfesor
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv = au.ID_AnUniv
                WHERE vcm.NumeIntregProfesor IS NOT NULL
                  AND (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(vcm.DenumireFacultate, CHAR(9), ''), 'S', 'S'), 'T', 'T')))) = @fac)
                  AND (@allSpecs = 1 OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                        CASE WHEN CHARINDEX('+', vcm.DenumireSpecializare) > 0 
                             THEN LEFT(vcm.DenumireSpecializare, CHARINDEX('+', vcm.DenumireSpecializare) - 1)
                             ELSE vcm.DenumireSpecializare END,
                      'S', 'S'), 'T', 'T')))) IN (SELECT value FROM STRING_SPLIT(@listaSpecs, ',')))
                ORDER BY vcm.NumeIntregProfesor";
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@allSpecs", toateSpec ? 1 : 0);
            cmd.Parameters.AddWithValue("@listaSpecs", specializari ?? "");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                lista.Add(reader["NumeIntregProfesor"]?.ToString() ?? "");
            return Ok(lista);
        }

        #endregion

        #region ================= HELPER SQL COMUN =================
        // Blocul BaseData folosit in toate rapoartele 1-7
        // Sursa: View_CentralizareMateriiProfesor (are anul 45, NumeIntregProfesor, DenumireFacultate)
        // JOIN StatDeFunctiiPeSpecializare pentru DenTitularSauSuplinitor
        // DenumireCatedra rezolvata prin MapareCatedra (dictionar static)
        // =====================================================================
        // BaseData CTE - sursa comuna pentru rapoartele 1-7
        // Filtrul pe facultate foloseste COLLATE Latin1_General_CI_AI 
        // (DenumireFacultate din DB are diacritice, se compara case+accent insensitive)
        // Forma de invatamant NU mai e in cascada - se pastreaza ca filtru simplu pe specializare
        // =====================================================================
        private const string BaseDataSql = @"
            WITH BaseData AS (
                SELECT 
                    vcm.NumeIntregProfesor                                          AS NumeIntreg,
                    vcm.ID_Profesor                                                  AS ID_Profesor,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                        CASE WHEN CHARINDEX('+', vcm.DenumireSpecializare) > 0 
                             THEN LEFT(vcm.DenumireSpecializare, CHARINDEX('+', vcm.DenumireSpecializare) - 1)
                             ELSE vcm.DenumireSpecializare END,
                    'S', 'S'), 'T', 'T')))) COLLATE DATABASE_DEFAULT                AS SpecializareCurata,
                    vcm.DenumireSpecializare                                         AS NumeSpecOriginal,
                    ISNULL(vcm.DenumireMaterie, 'Nedefinit')                         AS DenumireMaterie,
                    ISNULL(sf.DenTitularSauSuplinitor, 'Nespecificat')               AS TipPost,
                    ISNULL(vcm.NrSemestruDinAn, 0)                                  AS Semestru,
                    CAST(ISNULL(vcm.NrOreConventionale, 0) AS DECIMAL(10,4))        AS OreConvLinie,
                    ISNULL(vcm.Nr_Ore_Curs, 0)                                      AS OreCursLinie,
                    ISNULL(vcm.Nr_Ore_Seminar, 0) + ISNULL(vcm.Nr_Ore_Laborator, 0) + ISNULL(vcm.Nr_Ore_Proiect, 0) AS OreAplicatiiLinie,
                    ISNULL(vcm.Nr_Ore_Seminar, 0)                                   AS OreSeminarLinie,
                    ISNULL(vcm.Nr_Ore_Laborator, 0)                                 AS OreLaboratorLinie,
                    ISNULL(vcm.Nr_Ore_Proiect, 0)                                   AS OreProiectLinie,
                    LTRIM(RTRIM(ISNULL(vcm.DenumireFacultate, '')))                 AS FacultateCurata,
                    vcm.StatDeFunctiiID_Catedra                                      AS ID_Catedra,
                    UPPER(LTRIM(RTRIM(au.Denumire))) COLLATE DATABASE_DEFAULT        AS AnCurat
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv = au.ID_AnUniv
                LEFT JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf 
                    ON sf.ID_StatDeFunctii = vcm.ID_StatDeFunctii
                    AND sf.ID_AnUniv = vcm.ID_AnUniv
                    AND sf.DenumireSpecializare = vcm.DenumireSpecializare
                    AND sf.DenumireMaterie = vcm.DenumireMaterie
                    AND sf.NrSemestruDinAn = vcm.NrSemestruDinAn
            )";
        #endregion

        #region ================= RAPORT 1: NORMA PROFESORI =================

        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string? anUniv, string? facultate, string? specializari,
            string? profesor, int semestru = 0, string tipPost = "Toti",
            string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            string sql = BaseDataSql + @",
            Filtrat AS (
                SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru, ID_Catedra,
                       SUM(OreConvLinie)       AS TotalOreConvItem,
                       SUM(OreCursLinie)       AS TotalOreCurs,
                       SUM(OreAplicatiiLinie)  AS TotalOreAplicatii
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@prof = 'Toti' OR NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru, ID_Catedra
            )
            SELECT f.NumeIntreg AS Profesor, f.SpecializareCurata AS Specializare,
                   f.DenumireMaterie AS Materie, f.TipPost, f.Semestru, f.ID_Catedra,
                   f.TotalOreCurs AS NrOreCurs, f.TotalOreAplicatii AS NrOreAplicatii,
                   f.TotalOreConvItem AS NrOreConventionale,
                   SUM(f.TotalOreConvItem) OVER(PARTITION BY f.NumeIntreg, f.TipPost) AS TotalTipPost,
                   SUM(f.TotalOreConvItem) OVER(PARTITION BY f.NumeIntreg) AS TotalPost
            FROM Filtrat f
            ORDER BY f.NumeIntreg, f.TipPost, f.DenumireMaterie";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                result.Add(new
                {
                    Profesor = reader["Profesor"]?.ToString() ?? "",
                    Specializare = reader["Specializare"]?.ToString() ?? "",
                    Materie = reader["Materie"]?.ToString() ?? "",
                    TipPost = reader["TipPost"]?.ToString() ?? "",
                    Semestru = reader["Semestru"],
                    Departament = GetDenumireCatedra(idCat),
                    NrOreCurs = Convert.ToDouble(reader["NrOreCurs"]),
                    NrOreAplicatii = Convert.ToDouble(reader["NrOreAplicatii"]),
                    NrOreConventionale = Convert.ToDouble(reader["NrOreConventionale"]),
                    TotalTipPost = Convert.ToDouble(reader["TotalTipPost"]),
                    TotalPost = Convert.ToDouble(reader["TotalPost"])
                });
            }
            return Ok(result);
        }

        [HttpGet("export/norma")]
        public IActionResult ExportNormaExcel(string? anUniv, string? facultate, string? specializari,
            string? profesor, int semestru = 0, string tipPost = "Toti",
            string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"), new DataColumn("Specializare"), new DataColumn("Materie"),
                new DataColumn("Departament"), new DataColumn("Tip Post"), new DataColumn("Semestru"),
                new DataColumn("Nr Ore Curs", typeof(double)), new DataColumn("Nr Ore Aplicatii", typeof(double)),
                new DataColumn("Nr Ore Conventionale", typeof(double))
            });

            using var conn = new SqlConnection(_connectionString);
            string sql = BaseDataSql + @",
            Filtrat AS (
                SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru, ID_Catedra,
                       SUM(OreConvLinie) AS TotalOreConvItem,
                       SUM(OreCursLinie) AS TotalOreCurs,
                       SUM(OreAplicatiiLinie) AS TotalOreAplicatii
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@prof = 'Toti' OR NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru, ID_Catedra
            )
            SELECT f.NumeIntreg, f.SpecializareCurata, f.DenumireMaterie, f.TipPost, f.Semestru,
                   f.ID_Catedra, f.TotalOreCurs, f.TotalOreAplicatii, f.TotalOreConvItem
            FROM Filtrat f
            ORDER BY f.NumeIntreg, f.TipPost, f.DenumireMaterie";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dt.Rows.Add(reader["NumeIntreg"], reader["SpecializareCurata"], reader["DenumireMaterie"],
                    GetDenumireCatedra(idCat), reader["TipPost"], reader["Semestru"],
                    Convert.ToDouble(reader["TotalOreCurs"]), Convert.ToDouble(reader["TotalOreAplicatii"]),
                    Convert.ToDouble(reader["TotalOreConvItem"]));
            }

            string fileName = string.IsNullOrEmpty(profesor) || profesor == "Toti"
                ? "NormaProfesori_General.xlsx"
                : $"NormaProfesori_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.xlsx";

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme");
            ws.Cell(1, 1).Value = "Filtre Aplicate"; ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
            ws.Cell(2, 1).Value = $"An: {anUniv} | Facultate: {facultate} | Departament: {departament}";
            ws.Cell(3, 1).Value = $"Profesor: {profesor} | Semestru: {(semestru == 0 ? "Toate" : semestru.ToString())} | Tip Post: {tipPost}";
            var tbl = ws.Cell(5, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Nr Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Aplicatii").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr Ore Conventionale").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(5, 1, 5, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        #endregion

        #region ================= RAPORT 2: ORE PE PROGRAM =================

        [HttpGet("ore-profesor-program")]
        public async Task<IActionResult> GetOreProfProgram(string? anUniv = "Toti", string? facultate = "Toti",
            string? specializari = "Toti", string? profesor = "Toti", int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new List<object>();
            string sql = BaseDataSql + @",
            Filtrat AS (
                SELECT NumeIntreg AS Profesor, SpecializareCurata AS ProgramStudiu,
                       SUM(OreConvLinie) AS OreConvProgram
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@prof = 'Toti' OR NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                GROUP BY NumeIntreg, SpecializareCurata
            ),
            TotalProfesor AS (SELECT Profesor, SUM(OreConvProgram) AS TotalPost FROM Filtrat GROUP BY Profesor)
            SELECT f.Profesor, ISNULL(f.ProgramStudiu, 'Nespecificat') AS ProgramStudiu,
                   f.OreConvProgram AS NrOreConv, t.TotalPost,
                   CAST(CASE WHEN t.TotalPost = 0 THEN 0 ELSE (f.OreConvProgram / t.TotalPost) * 100 END AS DECIMAL(10,2)) AS ProcentPost
            FROM Filtrat f INNER JOIN TotalProfesor t ON f.Profesor = t.Profesor
            ORDER BY f.Profesor, f.OreConvProgram DESC";

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                result.Add(new
                {
                    Profesor = reader["Profesor"]?.ToString() ?? "",
                    ProgramStudiu = reader["ProgramStudiu"]?.ToString() ?? "",
                    NrOreConv = Convert.ToDouble(reader["NrOreConv"]),
                    TotalPost = Convert.ToDouble(reader["TotalPost"]),
                    ProcentPost = Convert.ToDouble(reader["ProcentPost"])
                });
            }
            return Ok(result);
        }

        [HttpGet("export/ore-program")]
        public async Task<IActionResult> ExportOreProgramExcel(string? anUniv, string? facultate,
            string? specializari, string? profesor, int semestru = 0, string tipPost = "Toti",
            string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"), new DataColumn("Program Studiu"),
                new DataColumn("Nr Ore Conv", typeof(double)), new DataColumn("Procent Post", typeof(double)),
                new DataColumn("Total Post", typeof(double))
            });

            string sql = BaseDataSql + @",
            Filtrat AS (
                SELECT NumeIntreg, SpecializareCurata AS ProgramStudiu, SUM(OreConvLinie) AS OreConvProgram
                FROM BaseData
                WHERE (@an = 'Toti' OR AnCurat = @an)
                  AND (@fac = 'Toti' OR FacultateCurata = @fac)
                  AND (@prof = 'Toti' OR NumeIntreg = @prof)
                  AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
                  AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
                  AND (@semestru = 0 OR Semestru = @semestru)
                  AND (@tipPost = 'Toti' OR TipPost = @tipPost)
                GROUP BY NumeIntreg, SpecializareCurata
            ),
            TotalProfesor AS (SELECT NumeIntreg, SUM(OreConvProgram) AS TotalPost FROM Filtrat GROUP BY NumeIntreg)
            SELECT f.NumeIntreg, ISNULL(f.ProgramStudiu,'Nespecificat') AS ProgramStudiu,
                   f.OreConvProgram, t.TotalPost,
                   CAST(CASE WHEN t.TotalPost=0 THEN 0 ELSE (f.OreConvProgram/t.TotalPost)*100 END AS DECIMAL(10,2)) AS ProcentPost
            FROM Filtrat f INNER JOIN TotalProfesor t ON f.NumeIntreg=t.NumeIntreg
            ORDER BY f.NumeIntreg, f.OreConvProgram DESC";

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                dt.Rows.Add(reader["NumeIntreg"], reader["ProgramStudiu"],
                    Convert.ToDouble(reader["OreConvProgram"]), Convert.ToDouble(reader["ProcentPost"]),
                    Convert.ToDouble(reader["TotalPost"]));

            string fileName = string.IsNullOrEmpty(profesor) || profesor == "Toti"
                ? "StatisticaOre_General.xlsx"
                : $"StatisticaOre_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.xlsx";

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Distributie Ore");
            ws.Cell(1, 1).Value = $"An: {anUniv} | Facultate: {facultate} | Profesor: {profesor}";
            var tbl = ws.Cell(3, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(3, 1, 3, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        #endregion

        #region ================= RAPORT 3: NORME TOTALURI =================

        [HttpGet("norma-totaluri")]
        public ActionResult GetNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            string sql = @"
            WITH DateProfesor AS (
                SELECT 
                    vcm.NumeIntregProfesor                                                      AS NumeComplet,
                    vcm.StatDeFunctiiID_Catedra                                                 AS ID_Catedra,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(vcm.DenumireFacultate,''), CHAR(9),''),'S','S'),'T','T')))) AS Facultate,
                    CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4))                    AS OreConv,
                    UPPER(LTRIM(RTRIM(au.Denumire)))                                            AS AnCurat
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv = au.ID_AnUniv
                WHERE (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(vcm.DenumireFacultate,CHAR(9),''),'S','S'),'T','T')))) = @fac)
                  AND (@prof = 'Toti' OR vcm.NumeIntregProfesor = @prof)
            ),
            Totalizat AS (
                SELECT NumeComplet, ID_Catedra, Facultate,
                       SUM(OreConv) AS TotalOreConv,
                       SUM(OreConv) * 14 AS TotalAnualOreConv,
                       ROW_NUMBER() OVER(PARTITION BY NumeComplet ORDER BY SUM(OreConv) DESC) AS Rang
                FROM DateProfesor
                GROUP BY NumeComplet, ID_Catedra, Facultate
            )
            SELECT NumeComplet, ID_Catedra, Facultate, TotalOreConv, TotalAnualOreConv
            FROM Totalizat WHERE Rang = 1
            ORDER BY NumeComplet";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                result.Add(new
                {
                    Profesor = reader["NumeComplet"]?.ToString() ?? "",
                    Departament = GetDenumireCatedra(idCat),
                    Facultate = reader["Facultate"]?.ToString() ?? "",
                    TotalOreConv = Math.Round(Convert.ToDecimal(reader["TotalOreConv"]), 2),
                    TotalAnualOreConv = Math.Round(Convert.ToDecimal(reader["TotalAnualOreConv"]), 2)
                });
            }
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public IActionResult ExportNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume Profesor"), new DataColumn("Departament"), new DataColumn("Facultate"),
                new DataColumn("Total Ore Conv.", typeof(decimal)), new DataColumn("Total tot anul ore conv.", typeof(decimal))
            });

            using var conn = new SqlConnection(_connectionString);
            string sql = @"
            WITH DateProfesor AS (
                SELECT vcm.NumeIntregProfesor AS NumeComplet, vcm.StatDeFunctiiID_Catedra AS ID_Catedra,
                       UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(vcm.DenumireFacultate,''),CHAR(9),''),'S','S'),'T','T')))) AS Facultate,
                       CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4)) AS OreConv
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv = au.ID_AnUniv
                WHERE (@an = 'Toti' OR UPPER(LTRIM(RTRIM(au.Denumire))) = @an)
                  AND (@fac = 'Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(vcm.DenumireFacultate,CHAR(9),''),'S','S'),'T','T')))) = @fac)
                  AND (@prof = 'Toti' OR vcm.NumeIntregProfesor = @prof)
            ),
            Totalizat AS (
                SELECT NumeComplet, ID_Catedra, Facultate,
                       SUM(OreConv) AS TotalOreConv, SUM(OreConv)*14 AS TotalAnualOreConv,
                       ROW_NUMBER() OVER(PARTITION BY NumeComplet ORDER BY SUM(OreConv) DESC) AS Rang
                FROM DateProfesor GROUP BY NumeComplet, ID_Catedra, Facultate
            )
            SELECT NumeComplet, ID_Catedra, Facultate, TotalOreConv, TotalAnualOreConv
            FROM Totalizat WHERE Rang=1 ORDER BY NumeComplet";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dt.Rows.Add(reader["NumeComplet"], GetDenumireCatedra(idCat), reader["Facultate"],
                    Math.Round(Convert.ToDecimal(reader["TotalOreConv"]), 2),
                    Math.Round(Convert.ToDecimal(reader["TotalAnualOreConv"]), 2));
            }

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Totaluri Norme");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Total Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total tot anul ore conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nume Profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Totaluri_Norme.xlsx");
        }

        #endregion

        #region ================= RAPORT 4: LIMBI STRAINE =================

        [HttpGet("limbi-straine")]
        public ActionResult GetLimbiStraine(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            string sql = BaseDataSql + @"
            SELECT 
                NumeIntreg AS NumeComplet,
                SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConvLinie ELSE 0 END) * 14 AS Sem1,
                SUM(CASE WHEN Semestru IN (2,4,6,8,10,12) THEN OreConvLinie ELSE 0 END) * 14 AS Sem2,
                SUM(OreConvLinie) * 14 AS Total
            FROM BaseData
            WHERE (@an = 'Toti' OR AnCurat = @an)
              AND (@fac = 'Toti' OR FacultateCurata = @fac)
              AND (@formaInv = 'Toti' OR NumeSpecOriginal LIKE '% ' + @formaInv + '%' OR NumeSpecOriginal LIKE '%-' + @formaInv + '%')
              AND (@prof = 'Toti' OR NumeIntreg = @prof)
              AND (@specs = 'Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs, ',')))
              AND (@semestru = 0 OR Semestru = @semestru)
              AND (@tipPost = 'Toti' OR TipPost = @tipPost)
              AND (
                   NumeSpecOriginal LIKE '%englez%' OR NumeSpecOriginal LIKE '%francez%' 
                   OR NumeSpecOriginal LIKE '%german%' OR NumeSpecOriginal LIKE '%american%'
                   OR NumeSpecOriginal LIKE '%(EN)%' OR NumeSpecOriginal LIKE '%(FR)%' OR NumeSpecOriginal LIKE '%(G)%'
                   OR NumeSpecOriginal IN (
                       'Inginerie virtuala in proiectarea autovehiculelor',
                       'Metode practice integrate in ingineria sistemelor de propulsie',
                       'Ingineria proceselor de fabricatie avansate',
                       'Managementul afacerilor industriale si antreprenoriat',
                       'Inginerie electrica si calculatoare','Sisteme electrice avansate',
                       'Securitate cibernetica','Informatica aplicata','Tehnologii Internet',
                       'Cultura si discurs in spatiul anglo american',
                       'Studii de limba si de cultura franceza',
                       'Studii de limba si literatura germana din perspectiva interculturala',
                       'Studii lingvistice pentru comunicare interculturala',
                       'Traducere si interpretariat din limba franceza in limba romana',
                       'Studii americane','Performanta umana in antrenamentul sportiv',
                       'Administrarea afacerilor','Managementul resurselor umane',
                       'Dezvoltarea afacerilor turistice','Medicina traditionala chineza'
                   )
              )
            GROUP BY NumeIntreg
            HAVING SUM(OreConvLinie) > 0
            ORDER BY NumeIntreg";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            int nr = 1;
            while (reader.Read())
            {
                result.Add(new
                {
                    NrCrt = nr++,
                    NumeProfesor = reader["NumeComplet"]?.ToString() ?? "",
                    Sem1 = Convert.ToDouble(reader["Sem1"]),
                    Sem2 = Convert.ToDouble(reader["Sem2"]),
                    Total = Convert.ToDouble(reader["Total"])
                });
            }
            return Ok(result);
        }

        [HttpGet("export/limbi-straine")]
        public IActionResult ExportLimbiStraine(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nr. Crt.", typeof(int)), new DataColumn("Nume si prenume profesor"),
                new DataColumn("Total Sem 1", typeof(decimal)), new DataColumn("Total Sem 2", typeof(decimal)),
                new DataColumn("Total", typeof(decimal))
            });

            using var conn = new SqlConnection(_connectionString);
            string sql = BaseDataSql + @"
            SELECT NumeIntreg AS NumeComplet,
                   SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConvLinie ELSE 0 END)*14 AS Sem1,
                   SUM(CASE WHEN Semestru IN (2,4,6,8,10,12) THEN OreConvLinie ELSE 0 END)*14 AS Sem2,
                   SUM(OreConvLinie)*14 AS Total
            FROM BaseData
            WHERE (@an='Toti' OR AnCurat=@an) AND (@fac='Toti' OR FacultateCurata=@fac)
              AND (@formaInv='Toti' OR NumeSpecOriginal LIKE '% '+@formaInv+'%' OR NumeSpecOriginal LIKE '%-'+@formaInv+'%')
              AND (@prof='Toti' OR NumeIntreg=@prof)
              AND (@specs='Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
              AND (@semestru=0 OR Semestru=@semestru) AND (@tipPost='Toti' OR TipPost=@tipPost)
              AND (NumeSpecOriginal LIKE '%englez%' OR NumeSpecOriginal LIKE '%francez%'
                   OR NumeSpecOriginal LIKE '%german%' OR NumeSpecOriginal LIKE '%american%'
                   OR NumeSpecOriginal LIKE '%(EN)%' OR NumeSpecOriginal LIKE '%(FR)%' OR NumeSpecOriginal LIKE '%(G)%'
                   OR NumeSpecOriginal IN ('Inginerie virtuala in proiectarea autovehiculelor',
                       'Metode practice integrate in ingineria sistemelor de propulsie',
                       'Ingineria proceselor de fabricatie avansate',
                       'Managementul afacerilor industriale si antreprenoriat',
                       'Inginerie electrica si calculatoare','Sisteme electrice avansate',
                       'Securitate cibernetica','Informatica aplicata','Tehnologii Internet',
                       'Cultura si discurs in spatiul anglo american','Studii de limba si de cultura franceza',
                       'Studii de limba si literatura germana din perspectiva interculturala',
                       'Studii lingvistice pentru comunicare interculturala',
                       'Traducere si interpretariat din limba franceza in limba romana',
                       'Studii americane','Performanta umana in antrenamentul sportiv',
                       'Administrarea afacerilor','Managementul resurselor umane',
                       'Dezvoltarea afacerilor turistice','Medicina traditionala chineza'))
            GROUP BY NumeIntreg HAVING SUM(OreConvLinie)>0 ORDER BY NumeIntreg";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            int nr = 1;
            while (reader.Read())
                dt.Rows.Add(nr++, reader["NumeComplet"],
                    Convert.ToDecimal(reader["Sem1"]), Convert.ToDecimal(reader["Sem2"]),
                    Convert.ToDecimal(reader["Total"]));

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Limbi Straine");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Total Sem 1").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total Sem 2").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nume si prenume profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Raport_Limbi_Straine.xlsx");
        }

        #endregion

        #region ================= RAPORT 5: DISCIPLINE PREDATE =================

        [HttpGet("discipline-predate")]
        public ActionResult GetDisciplinePredate(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            string sql = BaseDataSql + @"
            SELECT DISTINCT 
                NumeIntreg AS NumeComplet, ID_Catedra, DenumireMaterie,
                STUFF(
                    CASE WHEN OreCursLinie > 0 THEN ', Curs' ELSE '' END +
                    CASE WHEN OreSeminarLinie > 0 THEN ', Seminar' ELSE '' END +
                    CASE WHEN OreLaboratorLinie > 0 THEN ', Laborator' ELSE '' END +
                    CASE WHEN OreProiectLinie > 0 THEN ', Proiect' ELSE '' END
                ,1,2,'') AS TipActivitate
            FROM BaseData
            WHERE (@an='Toti' OR AnCurat=@an) AND (@fac='Toti' OR FacultateCurata=@fac)
              AND (@formaInv='Toti' OR NumeSpecOriginal LIKE '% '+@formaInv+'%' OR NumeSpecOriginal LIKE '%-'+@formaInv+'%')
              AND (@prof='Toti' OR NumeIntreg=@prof)
              AND (@specs='Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
              AND (@semestru=0 OR Semestru=@semestru) AND (@tipPost='Toti' OR TipPost=@tipPost)
            ORDER BY NumeIntreg, DenumireMaterie";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                result.Add(new
                {
                    Profesor = reader["NumeComplet"]?.ToString() ?? "",
                    Departament = GetDenumireCatedra(idCat),
                    Materie = reader["DenumireMaterie"]?.ToString() ?? "",
                    TipActivitate = reader["TipActivitate"]?.ToString() ?? ""
                });
            }
            return Ok(result);
        }

        [HttpGet("export/discipline-predate")]
        public IActionResult ExportDisciplinePredate(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"), new DataColumn("Departament"),
                new DataColumn("Disciplina"), new DataColumn("Tip Activitate")
            });

            using var conn = new SqlConnection(_connectionString);
            string sql = BaseDataSql + @"
            SELECT DISTINCT NumeIntreg, ID_Catedra, DenumireMaterie,
                   STUFF(CASE WHEN OreCursLinie>0 THEN ', Curs' ELSE '' END+
                         CASE WHEN OreSeminarLinie>0 THEN ', Seminar' ELSE '' END+
                         CASE WHEN OreLaboratorLinie>0 THEN ', Laborator' ELSE '' END+
                         CASE WHEN OreProiectLinie>0 THEN ', Proiect' ELSE '' END,1,2,'') AS TipActivitate
            FROM BaseData
            WHERE (@an='Toti' OR AnCurat=@an) AND (@fac='Toti' OR FacultateCurata=@fac)
              AND (@formaInv='Toti' OR NumeSpecOriginal LIKE '% '+@formaInv+'%' OR NumeSpecOriginal LIKE '%-'+@formaInv+'%')
              AND (@prof='Toti' OR NumeIntreg=@prof)
              AND (@specs='Toti' OR SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
              AND (@semestru=0 OR Semestru=@semestru) AND (@tipPost='Toti' OR TipPost=@tipPost)
            ORDER BY NumeIntreg, DenumireMaterie";

            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dt.Rows.Add(reader["NumeIntreg"], GetDenumireCatedra(idCat),
                    reader["DenumireMaterie"], reader["TipActivitate"]?.ToString() ?? "");
            }

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Discipline Predate");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Discipline_Predate.xlsx");
        }

        #endregion

        #region ================= RAPORT 6: TITULARI =================

        // TITULARI = profesori cu cel putin un post cu TitularSauSuplinitor=1
        // in View_PostProfesorMaterie (norma de baza la UTBv).
        // NU se foloseste DenTitularSauSuplinitor='Tit' din StatDeFunctiiPeSpecializare
        // care include si asociati cu titulatura pe materie specifica.
        private const string TitulariSubquery = @"
            EXISTS (
                SELECT 1 FROM [AGSIS].[pi].[View_PostProfesorMaterie] vp2
                WHERE vp2.ID_Profesor = vcm.ID_Profesor
                  AND vp2.ID_AnUniv = vcm.ID_AnUniv
                  AND vp2.TitularSauSuplinitor = 1
            )";

        [HttpGet("titulari")]
        public ActionResult GetTitulari(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            string sql = $@"
            SELECT DISTINCT 
                vcm.NumeIntregProfesor                      AS NumeComplet,
                vcm.ID_Profesor                             AS ID_Profesor,
                vcm.StatDeFunctiiID_Catedra                 AS ID_Catedra,
                ISNULL(vcm.DenumireFacultate,'Nespecificat') AS Facultate
            FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
            INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv = au.ID_AnUniv
            WHERE {TitulariSubquery}
              AND (@an='Toti' OR UPPER(LTRIM(RTRIM(au.Denumire)))=@an)
              AND (@fac='Toti' OR vcm.DenumireFacultate COLLATE Latin1_General_CI_AI = @fac COLLATE Latin1_General_CI_AI)
            ORDER BY vcm.NumeIntregProfesor";
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                result.Add(new
                {
                    Profesor = FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    Departament = GetDenumireCatedra(idCat),
                    Facultate = reader["Facultate"]?.ToString() ?? ""
                });
            }
            return Ok(result);
        }

        [HttpGet("export/titulari")]
        public IActionResult ExportTitulari(string? anUniv, string? facultate, string? departament)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"), new DataColumn("Departament"), new DataColumn("Facultate")
            });
            using var conn = new SqlConnection(_connectionString);
            string sql = $@"
            SELECT DISTINCT vcm.NumeIntregProfesor AS NumeComplet, vcm.ID_Profesor,
                   vcm.StatDeFunctiiID_Catedra AS ID_Catedra,
                   ISNULL(vcm.DenumireFacultate,'Nespecificat') AS Facultate
            FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
            INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv=au.ID_AnUniv
            WHERE {TitulariSubquery}
              AND (@an='Toti' OR UPPER(LTRIM(RTRIM(au.Denumire)))=@an)
              AND (@fac='Toti' OR vcm.DenumireFacultate COLLATE Latin1_General_CI_AI = @fac COLLATE Latin1_General_CI_AI)
            ORDER BY vcm.NumeIntregProfesor";
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dt.Rows.Add(
                    FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    GetDenumireCatedra(idCat),
                    reader["Facultate"]);
            }
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Titulari");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Cadre_Didactice_Titulare.xlsx");
        }

        #endregion

        #region ================= RAPORT 7: COLABORATORI =================

        // COLABORATORI = profesori care NU au niciun post cu TitularSauSuplinitor=1
        // in View_PostProfesorMaterie pentru anul respectiv (norma de baza NU e la UTBv)

        [HttpGet("colaboratori")]
        public ActionResult GetColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            string sql = $@"
            SELECT DISTINCT vcm.NumeIntregProfesor AS NumeComplet, vcm.ID_Profesor,
                   vcm.StatDeFunctiiID_Catedra AS ID_Catedra,
                   ISNULL(vcm.DenumireFacultate,'Nespecificat') AS Facultate
            FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
            INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv=au.ID_AnUniv
            WHERE NOT {TitulariSubquery}
              AND (@an='Toti' OR UPPER(LTRIM(RTRIM(au.Denumire)))=@an)
              AND (@fac='Toti' OR vcm.DenumireFacultate COLLATE Latin1_General_CI_AI = @fac COLLATE Latin1_General_CI_AI)
            ORDER BY vcm.NumeIntregProfesor";
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                result.Add(new
                {
                    Profesor = FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    Departament = GetDenumireCatedra(idCat),
                    Facultate = reader["Facultate"]?.ToString() ?? ""
                });
            }
            return Ok(result);
        }

        [HttpGet("export/colaboratori")]
        public IActionResult ExportColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Nume si prenume"), new DataColumn("Departament"), new DataColumn("Facultate")
            });
            using var conn = new SqlConnection(_connectionString);
            string sql = $@"
            SELECT DISTINCT vcm.NumeIntregProfesor AS NumeComplet, vcm.ID_Profesor,
                   vcm.StatDeFunctiiID_Catedra AS ID_Catedra,
                   ISNULL(vcm.DenumireFacultate,'Nespecificat') AS Facultate
            FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
            INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv=au.ID_AnUniv
            WHERE NOT {TitulariSubquery}
              AND (@an='Toti' OR UPPER(LTRIM(RTRIM(au.Denumire)))=@an)
              AND (@fac='Toti' OR vcm.DenumireFacultate COLLATE Latin1_General_CI_AI = @fac COLLATE Latin1_General_CI_AI)
            ORDER BY vcm.NumeIntregProfesor";
            conn.Open();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dt.Rows.Add(
                    FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    GetDenumireCatedra(idCat),
                    reader["Facultate"]);
            }
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Colaboratori");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Cadre_Didactice_Colaboratori.xlsx");
        }

        #endregion

        #region ================= RAPORT 8: ANS =================

        [HttpGet("date-ans")]
        public IActionResult GetDateANS([FromQuery] int idAnUniv = 45)
        {
            var dateBrute = new List<RandSqlANS>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string query = $@"
                SELECT 
                    vcm.NumeIntregProfesor                                          AS NumeComplet,
                    vcm.ID_Profesor                                                  AS ID_Profesor,
                    ppm.DenumireGradDidacticPost                                    AS GradFunctie,
                    CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4))        AS OreConventionale,
                    ISNULL(vcm.DenumireFacultate,'Nespecificat')                   AS Facultate,
                    vcm.StatDeFunctiiID_Catedra                                     AS ID_Catedra,
                    sf.id_metaspecializare                                          AS IdMetaspec
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    ON sf.ID_StatDeFunctii=vcm.ID_StatDeFunctii AND sf.ID_AnUniv=vcm.ID_AnUniv
                    AND sf.DenumireSpecializare=vcm.DenumireSpecializare AND sf.DenumireMaterie=vcm.DenumireMaterie
                    AND sf.NrSemestruDinAn=vcm.NrSemestruDinAn
                -- GradFunctie: din DW cel mai recent an disponibil per profesor
                LEFT JOIN (
                    SELECT ID_Profesor, DenumireGradDidacticPost,
                           ROW_NUMBER() OVER(PARTITION BY ID_Profesor ORDER BY ID_AnUniv DESC) AS Rn
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie]
                    WHERE DenumireGradDidacticPost IS NOT NULL
                ) ppm ON ppm.ID_Profesor = vcm.ID_Profesor AND ppm.Rn = 1
                WHERE vcm.ID_AnUniv = @ID_AnUniv
                  AND {TitulariSubquery}";

            using var cmd = new SqlCommand(query, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                int idMeta = reader["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(reader["IdMetaspec"]) : 0;
                if (!MappingMetaspec.TryGetValue(idMeta, out int idAns)) continue;
                if (!AnsIdToCol.ContainsKey(idAns)) continue;
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dateBrute.Add(new RandSqlANS
                {
                    NumeComplet = FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    Facultate = reader["Facultate"]?.ToString() ?? "",
                    Departament = GetDenumireCatedra(idCat),
                    GradFunctie = reader["GradFunctie"]?.ToString() ?? "",
                    OreConventionale = Convert.ToDecimal(reader["OreConventionale"]),
                    IdANS = idAns,
                });
            }

            var profesori = dateBrute
                .GroupBy(x => x.NumeComplet)
                .Select(g =>
                {
                    var grupDept = g.GroupBy(x => new { x.Departament, x.Facultate })
                        .Select(dg => new { dg.Key.Departament, dg.Key.Facultate, TotalOre = dg.Sum(x => x.OreConventionale), Grad = dg.OrderByDescending(x => x.OreConventionale).First().GradFunctie })
                        .OrderByDescending(d => d.TotalOre).First();
                    var orePerAns = g.GroupBy(x => x.IdANS).ToDictionary(ag => ag.Key, ag => ag.Sum(x => x.OreConventionale));
                    decimal totalOre = orePerAns.Values.Sum();
                    var fractiuni = new Dictionary<string, decimal>();
                    if (totalOre > 0)
                    {
                        int maxKey = orePerAns.OrderByDescending(x => x.Value).First().Key;
                        decimal sum = 0;
                        foreach (var kv in orePerAns)
                        {
                            if (kv.Key == maxKey) continue;
                            decimal frac = Math.Round(kv.Value / totalOre, 2);
                            fractiuni[DomeniiExcel[AnsIdToCol[kv.Key] - 10]] = frac;
                            sum += frac;
                        }
                        fractiuni[DomeniiExcel[AnsIdToCol[maxKey] - 10]] = Math.Round(1m - sum, 2);
                    }
                    return new { NumeComplet = g.Key, grupDept.Facultate, grupDept.Departament, GradFunctie = MapareFunctieANS(grupDept.Grad), DomeniiMapate = fractiuni };
                })
                .OrderBy(p => p.NumeComplet).ToList();

            return Ok(profesori);
        }

        [HttpGet("export/raport-ans")]
        public IActionResult ExportRaportANS([FromQuery] int idAnUniv = 45)
        {
            var dateBrute = new List<RandSqlANS>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            string query = @"
                SELECT vcm.NumeIntregProfesor AS NumeComplet, vcm.ID_Profesor,
                       ppm.DenumireGradDidacticPost AS GradFunctie,
                       vcm.StatDeFunctiiID_Catedra AS ID_Catedra,
                       ISNULL(vcm.DenumireFacultate,'') AS Facultate,
                       CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4)) AS OreConventionale,
                       sf.id_metaspecializare AS IdMetaspec
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[pi].[StatDeFunctiiPeSpecializare] sf
                    ON sf.ID_StatDeFunctii=vcm.ID_StatDeFunctii AND sf.ID_AnUniv=vcm.ID_AnUniv
                    AND sf.DenumireSpecializare=vcm.DenumireSpecializare AND sf.DenumireMaterie=vcm.DenumireMaterie
                    AND sf.NrSemestruDinAn=vcm.NrSemestruDinAn
                LEFT JOIN (
                    SELECT ID_Profesor, DenumireGradDidacticPost,
                           ROW_NUMBER() OVER(PARTITION BY ID_Profesor ORDER BY ID_AnUniv DESC) AS Rn
                    FROM [agsis_dw].[dbo].[Post_Profesor_Materie]
                    WHERE DenumireGradDidacticPost IS NOT NULL
                ) ppm ON ppm.ID_Profesor=vcm.ID_Profesor AND ppm.Rn=1
                WHERE vcm.ID_AnUniv=@ID_AnUniv
                  AND {TitulariSubquery}";

            using var cmd = new SqlCommand(query, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                int idMeta = reader["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(reader["IdMetaspec"]) : 0;
                if (!MappingMetaspec.TryGetValue(idMeta, out int idAns)) continue;
                if (!AnsIdToCol.ContainsKey(idAns)) continue;
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dateBrute.Add(new RandSqlANS
                {
                    NumeComplet = FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    Departament = GetDenumireCatedra(idCat),
                    Facultate = reader["Facultate"]?.ToString() ?? "",
                    GradFunctie = reader["GradFunctie"]?.ToString() ?? "",
                    OreConventionale = Convert.ToDecimal(reader["OreConventionale"]),
                    IdANS = idAns,
                });
            }

            var profesori = new List<ProfANS>();
            foreach (var grp in dateBrute.GroupBy(x => x.NumeComplet))
            {
                var grupDept = grp.GroupBy(x => new { x.Departament, x.Facultate })
                    .Select(dg => new { dg.Key.Departament, dg.Key.Facultate, TotalOre = dg.Sum(x => x.OreConventionale), Grad = dg.OrderByDescending(x => x.OreConventionale).First().GradFunctie })
                    .OrderByDescending(d => d.TotalOre).First();
                var orePerCol = new Dictionary<int, decimal>();
                foreach (var rand in grp)
                {
                    int col = AnsIdToCol[rand.IdANS];
                    if (!orePerCol.ContainsKey(col)) orePerCol[col] = 0m;
                    orePerCol[col] += rand.OreConventionale;
                }
                decimal totalOre = orePerCol.Values.Sum();
                var fractiuni = new Dictionary<int, decimal>();
                if (totalOre > 0)
                {
                    int maxKey = orePerCol.OrderByDescending(x => x.Value).First().Key;
                    decimal sum = 0;
                    foreach (var kv in orePerCol)
                    {
                        if (kv.Key == maxKey) continue;
                        decimal frac = Math.Round(kv.Value / totalOre, 2);
                        fractiuni[kv.Key] = frac; sum += frac;
                    }
                    fractiuni[maxKey] = Math.Round(1m - sum, 2);
                }
                profesori.Add(new ProfANS { NumeComplet = grp.Key, Departament = grupDept.Departament, Facultate = grupDept.Facultate, GradFunctie = MapareFunctieANS(grupDept.Grad), Fractiuni = fractiuni });
            }

            var overrides = new Dictionary<string, Dictionary<int, decimal>>
            {
                ["VOLMER MARIUS"] = new Dictionary<int, decimal> { { AnsIdToCol[7], 0.83m }, { AnsIdToCol[12], 0.17m } },
                ["ZAHARIA SEBASTIAN MARIAN"] = new Dictionary<int, decimal> { { AnsIdToCol[12], 0.74m }, { AnsIdToCol[9], 0.27m } },
            };
            foreach (var prof in profesori)
                if (overrides.TryGetValue(prof.NumeComplet, out var ov)) prof.Fractiuni = ov;

            profesori = profesori.OrderBy(p => p.NumeComplet).ToList();
            var wb = BuildANSWorkbookFromScratch();
            var ws = wb.Worksheets.First();
            int dataStartRow = 9;
            if (profesori.Count > 0) ws.Row(dataStartRow).InsertRowsBelow(profesori.Count - 1);
            for (int i = 0; i < profesori.Count; i++)
            {
                var prof = profesori[i];
                int r = dataStartRow + i;
                ws.Cell(r, 1).Value = i + 1; ws.Cell(r, 2).Value = prof.NumeComplet;
                ws.Cell(r, 3).Value = ""; ws.Cell(r, 4).Value = prof.GradFunctie;
                ws.Cell(r, 5).Value = 1; ws.Cell(r, 6).Value = 0; ws.Cell(r, 7).Value = "";
                ws.Cell(r, 8).Value = prof.Facultate; ws.Cell(r, 9).Value = prof.Departament;
                foreach (var kv in prof.Fractiuni) ws.Cell(r, kv.Key).Value = kv.Value;
                ws.Cell(r, 50).FormulaA1 = $"=SUM(J{r}:AW{r})";
                if (i % 2 != 0)
                    for (int c = 1; c <= 50; c++)
                        ws.Cell(r, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5f5f5");
            }
            int totalRow = dataStartRow + profesori.Count;
            ws.Cell(totalRow, 1).Value = "Total general:"; ws.Cell(totalRow, 1).Style.Font.Bold = true;
            for (int c = 10; c <= 49; c++)
            {
                string cl = ColumnLetter(c);
                ws.Cell(totalRow, c).FormulaA1 = $"=SUM({cl}{dataStartRow}:{cl}{totalRow - 1})";
                ws.Cell(totalRow, c).Style.Font.Bold = true;
            }
            ws.Cell(totalRow, 50).FormulaA1 = $"=SUM(J{totalRow}:AW{totalRow})";
            ws.Cell(totalRow, 50).Style.Font.Bold = true;
            using var stream = new MemoryStream(); wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Date_ANS_{idAnUniv}.xlsx");
        }

        #endregion

        #region ================= HELPERS =================

        private void AddBaseParams(SqlCommand cmd, string? anUniv, string? facultate, string? departament,
            string? formaInvatamant, string? profesor, string? specializari, int semestru, string? tipPost)
        {
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            // facultate vine exact din DB (cu diacritice), compararea se face cu COLLATE CI_AI in SQL
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@dept", departament ?? "Toti");
            cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
            cmd.Parameters.AddWithValue("@semestru", semestru);
            cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");
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

        private XLWorkbook BuildANSWorkbookFromScratch()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("CD DRU");
            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, 50).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov";
            ws.Range(3, 1, 3, 6).Merge();
            ws.Cell(4, 1).Value = "NOTA: Se includ in tabel toate cadrele didactice si de cercetare titulare (cu norma de baza in universitate), indiferent de forma de angajare.";
            ws.Range(4, 1, 4, 9).Merge();
            ws.Cell(4, 10).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta. Va rugam sa completati numai spatiile marcate cu culoarea verde.";
            ws.Range(4, 10, 4, 27).Merge();
            ws.Cell(4, 28).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta. Va rugam sa completati numai spatiile marcate cu culoarea verde.";
            ws.Range(4, 28, 4, 50).Merge();
            ws.Cell(5, 1).Value = "Nr. \nCrt."; ws.Cell(5, 2).Value = "Nume si prenume cadru didactic";
            ws.Cell(5, 3).Value = "CNP"; ws.Cell(5, 4).Value = "Functie cadru didactic sau cercetare";
            ws.Cell(5, 5).Value = "Forma de angajare"; ws.Cell(5, 6).Value = "Calitate conducator doctorat";
            ws.Cell(5, 7).Value = "Varsta"; ws.Cell(5, 8).Value = "Facultate"; ws.Cell(5, 9).Value = "Departament";
            ws.Cell(5, 10).Value = "Matematica si stiinte ale naturii"; ws.Cell(5, 15).Value = "Stiinte ingineresti";
            ws.Cell(5, 22).Value = "Stiinte biologice si biomedicale"; ws.Cell(5, 28).Value = "Stiinte sociale";
            ws.Cell(5, 37).Value = "Stiinte umaniste si arte"; ws.Cell(5, 50).Value = "Total";
            ws.Range(5, 1, 7, 1).Merge(); ws.Range(5, 2, 7, 2).Merge(); ws.Range(5, 3, 7, 3).Merge();
            ws.Range(5, 4, 7, 4).Merge(); ws.Range(5, 5, 7, 5).Merge(); ws.Range(5, 6, 7, 6).Merge();
            ws.Range(5, 7, 7, 7).Merge(); ws.Range(5, 8, 7, 8).Merge(); ws.Range(5, 9, 7, 9).Merge();
            ws.Range(5, 10, 5, 14).Merge(); ws.Range(5, 15, 5, 21).Merge();
            ws.Range(5, 22, 5, 27).Merge(); ws.Range(5, 28, 5, 36).Merge();
            ws.Range(5, 37, 5, 49).Merge(); ws.Range(5, 50, 7, 50).Merge();
            string[] subdomenii = { "Matematica", "Informatica", "Fizica", "Chimie si inginerie chimica", "Stiintele pamantului si atmosferei", "Inginerie civila", "Inginerie electrica, electronica si telecomunicatii", "Inginerie geologica, mine, petrol si gaze", "Ingineria transporturilor", "Ingineria resurselor vegetale si animale", "Ingineria sistemelor, calculatoare si tehnologia informatiei", "Inginerie mecanica, mecatronica, inginerie industriala si management", "Biologie", "Biochimie", "Medicina", "Medicina veterinara", "Medicina dentara", "Farmacie", "Stiinte juridice", "Stiinte administrative", "Stiinte ale comunicarii", "Sociologie", "Stiinte politice", "Stiinte militare, informatii si ordine publica", "Stiinte economice (doar Cibernetica, statistica si informatica economica)", "Stiinte economice (fara Cibernetica, statistica si informatica economica)", "Psihologie si stiinte comportamentale", "Filologie", "Filosofie", "Istorie", "Teologie", "Studii culturale", "Arhitectura si urbanism", "Arte vizuale (fara Istoria si teoria artei)", "Arte vizuale (doar Istoria si teoria artei)", "Teatru si artele spectacolului", "Cinematografie si media", "Muzica (doar Interpretare muzicala)", "Muzica (fara Interpretare muzicala)", "Stiintele Sportului si Educatiei Fizice" };
            for (int i = 0; i < subdomenii.Length; i++) { ws.Cell(6, 10 + i).Value = subdomenii[i]; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }
            for (int i = 0; i < 9; i++) ws.Cell(8, i + 1).Value = ((char)('A' + i)).ToString();
            for (int i = 0; i < 41; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, 50).Value = "40";
            var headerFill = XLColor.FromHtml(BrandColorHex);
            for (int r = 5; r <= 8; r++) for (int c = 1; c <= 50; c++) { ws.Cell(r, c).Style.Font.Bold = true; ws.Cell(r, c).Style.Font.FontColor = XLColor.Black; ws.Cell(r, c).Style.Fill.BackgroundColor = XLColor.White; ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; ws.Cell(r, c).Style.Alignment.WrapText = true; ws.Cell(r, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ws.Cell(r, c).Style.Border.InsideBorder = XLBorderStyleValues.Thin; }
            ws.Range(5, 1, 5, 50).Style.Fill.BackgroundColor = headerFill;
            ws.Range(5, 1, 5, 50).Style.Font.FontColor = XLColor.White;
            ws.Column(1).Width = 5; ws.Column(2).Width = 30; ws.Column(3).Width = 14; ws.Column(4).Width = 22; ws.Column(5).Width = 10; ws.Column(6).Width = 12; ws.Column(7).Width = 8; ws.Column(8).Width = 28; ws.Column(9).Width = 28;
            for (int c = 10; c <= 50; c++) ws.Column(c).Width = 12;
            return wb;
        }

        private string MapareFunctieANS(string grad)
        {
            if (string.IsNullOrWhiteSpace(grad)) return "Asistent";
            string g = grad.ToLower();
            if (g.Contains("profesor")) return "Profesor";
            if (g.Contains("conferentiar")) return "Conferentiar";
            if (g.Contains("lector") || g.Contains("sef lucrari") || g.Contains("sl")) return "Lector/Sef de lucrari (SL)";
            if (g.Contains("asistent de cercetare")) return "Asistent de cercetare";
            if (g.Contains("asistent")) return "Asistent";
            if (g.Contains("preparator")) return "Preparator";
            if (g.Contains("cercetator stiintific i") || g.Contains("cs i")) return "Cercetator stiintific I (CS I)";
            if (g.Contains("cercetator stiintific ii") || g.Contains("cs ii")) return "Cercetator stiintific II (CS II)";
            if (g.Contains("cercetator stiintific iii") || g.Contains("cs iii")) return "Cercetator stiintific III (CS III)";
            if (g.Contains("cercetator")) return "Cercetator";
            return "Asistent";
        }

        public class ProfANS { public string NumeComplet { get; set; } = ""; public string Departament { get; set; } = ""; public string Facultate { get; set; } = ""; public string GradFunctie { get; set; } = ""; public Dictionary<int, decimal> Fractiuni { get; set; } = new(); }
        public class RandSqlANS { public string NumeComplet { get; set; } = ""; public string Facultate { get; set; } = ""; public string Departament { get; set; } = ""; public string GradFunctie { get; set; } = ""; public decimal OreConventionale { get; set; } public int IdANS { get; set; } }

        #endregion
    }
}