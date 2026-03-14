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
            { 862, "DPPD" }, { 613, "DPPD" }, { 863, "DPPD" },
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

        private static readonly Dictionary<int, string> NumeCorecte = new Dictionary<int, string>
        {
            { 6621, "ȘIREIU RAMONA DANIELA" }, { 6698, "BENEA ALINA-PETRUȚA" },
            { 6631, "CĂBAȘ NICOLAE SERGIU" },  { 4605, "CĂPRIȚĂ FLORIN" },
            { 4165, "CIORICEANU IONUȚ-HORIA" },{ 2375, "ILAȘ MAGDALENA" },
            { 6616, "MÂNDRU MARIA SPERANȚA" }, { 16893, "MILEA GIGUȘA-ROXANA" },
            { 16894, "NEGOIȚĂ ELENA" },         { 6803, "STAREȘU CAMELIA MARIANA" },
            { 6800, "ȘERBAN AGURIȚA DORINELA" },{ 5881, "TUCHEL IONUȚ-VLAD" },
            { 2899, "BREZEANU ALIN IONUȚ" },    { 4345, "DIACONU ȘTEFANIA-ROXANA" },
            { 4352, "MANIȘIU(VASILE) VIRGINIA IOANA" },
            { 4401, "FOLEA(VECERDI) CRISTINA AGNEȘ" },
            { 4821, "CHIRA CODRUȚA-ELENA" },
            { 5833, "MARCHIȘ(TOMA) MARIA-ALEXANDRA" },
            { 5884, "VEZETEU COSMIN-DĂNUȚ" },
            { 6716, "BĂȘEANU IONUȚ-CRISTIAN-COZMIN" },
            { 6721, "MÂNZĂȚANU DIANA" }, { 6761, "CIOPLEIAȘ BOGDAN-NICOLAE" },
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

        private string GetDenumireCatedra(long idCatedra) =>
            MapareCatedra.TryGetValue(idCatedra, out var den) ? den : $"Catedra {idCatedra}";

        #region ================= LISTE (DROPDOWNS) =================

        [HttpGet("liste/ani-universitari")]
        public ActionResult GetAni()
        {
            return Ok(_cache.GetOrCreate("ListaAniUniv", entry =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(60);
                var lista = new List<object>();
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = new SqlCommand(
                    "SELECT UPPER(LTRIM(RTRIM(Denumire))) COLLATE DATABASE_DEFAULT AS AnCurat FROM [AGSIS].[dbo].[AnUniversitar] WHERE Denumire IS NOT NULL ORDER BY Ordine DESC", conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read()) { var an = reader["AnCurat"]?.ToString() ?? ""; lista.Add(new { id = an, nume = an }); }
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
                conn.Open();
                using var cmd = new SqlCommand(
                    "SELECT DISTINCT vcm.DenumireFacultate FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm WHERE vcm.DenumireFacultate IS NOT NULL AND LTRIM(RTRIM(vcm.DenumireFacultate)) != '' ORDER BY vcm.DenumireFacultate ASC", conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read()) { var val = reader[0]?.ToString()?.Trim() ?? ""; if (!string.IsNullOrWhiteSpace(val)) lista.Add(val); }
                return lista;
            }));
        }

        [HttpGet("liste/departamente")]
        public ActionResult GetDepartamente(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(
                "SELECT DISTINCT vcm.StatDeFunctiiID_Catedra FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm WHERE vcm.StatDeFunctiiID_Catedra IS NOT NULL AND (@fac='Toti' OR vcm.DenumireFacultate COLLATE Latin1_General_CI_AI = @fac COLLATE Latin1_General_CI_AI) ORDER BY vcm.StatDeFunctiiID_Catedra", conn);
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            var idSet = new HashSet<string>();
            while (reader.Read())
                if (reader[0] != DBNull.Value) { var den = GetDenumireCatedra(Convert.ToInt64(reader[0])); if (idSet.Add(den)) lista.Add(den); }
            lista.Sort((a, b) => a == "Toti" ? -1 : b == "Toti" ? 1 : string.Compare(a, b));
            return Ok(lista);
        }

        [HttpGet("liste/specializari-per-facultate")]
        public ActionResult GetSpecializari(string? anUniv, string? numeFacultate)
        {
            var lista = new List<string> { "Toti" };
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"SELECT DISTINCT UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                    CASE WHEN CHARINDEX('+',vcm.DenumireSpecializare)>0 THEN LEFT(vcm.DenumireSpecializare,CHARINDEX('+',vcm.DenumireSpecializare)-1)
                         ELSE vcm.DenumireSpecializare END,' - CORECT',''),' CORECT',''),' - COPIE',''),'S','S'),'T','T')))) COLLATE DATABASE_DEFAULT AS SpecCurata
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                WHERE vcm.DenumireSpecializare IS NOT NULL
                  AND (@fac='Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(vcm.DenumireFacultate,CHAR(9),''),'S','S'),'T','T'))))=@fac)
                ORDER BY SpecCurata", conn);
            cmd.Parameters.AddWithValue("@fac", numeFacultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read()) { var val = reader[0]?.ToString() ?? ""; if (!string.IsNullOrWhiteSpace(val) && !lista.Contains(val)) lista.Add(val); }
            return Ok(lista);
        }

        [HttpGet("liste/profesori-per-specializari")]
        public ActionResult GetProfesori(string? anUniv, string? facultate, string? specializari, string? departament)
        {
            var lista = new List<string> { "Toti" };
            bool toateSpec = string.IsNullOrEmpty(specializari) || specializari == "Toti";
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(@"SELECT DISTINCT vcm.NumeIntregProfesor
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv=au.ID_AnUniv
                WHERE vcm.NumeIntregProfesor IS NOT NULL
                  AND (@an='Toti' OR UPPER(LTRIM(RTRIM(au.Denumire)))=@an)
                  AND (@fac='Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(vcm.DenumireFacultate,CHAR(9),''),'S','S'),'T','T'))))=@fac)
                  AND (@allSpecs=1 OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(CASE WHEN CHARINDEX('+',vcm.DenumireSpecializare)>0 THEN LEFT(vcm.DenumireSpecializare,CHARINDEX('+',vcm.DenumireSpecializare)-1) ELSE vcm.DenumireSpecializare END,'S','S'),'T','T')))) IN (SELECT value FROM STRING_SPLIT(@listaSpecs,',')))
                ORDER BY vcm.NumeIntregProfesor", conn);
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@allSpecs", toateSpec ? 1 : 0);
            cmd.Parameters.AddWithValue("@listaSpecs", specializari ?? "");
            using var reader = cmd.ExecuteReader();
            while (reader.Read()) lista.Add(reader[0]?.ToString() ?? "");
            return Ok(lista);
        }

        #endregion

        #region ================= HELPER SQL COMUN =================

        // =====================================================================
        // BaseData CTE - sursa comuna rapoartele 1-5
        //
        // REGULA CRITICA TipPost:
        // - TipPost vine EXCLUSIV din DenTitularSauSuplinitor din StatDeFunctiiPeSpecializare
        // - Valorile 'Tit'/'Titular'/'Titulara' -> 'Titular'
        // - Valorile 'Sup'/'Suplinitor'/'SupTit' si orice altceva -> 'Suplinitor'
        // - NU se amesteca niciodata randurile de Titular cu cele de Suplinitor!
        // - Gruparea cuplajelor se face OBLIGATORIU pe (Profesor, Materie, TipPost, Semestru)
        //
        // Filtrul @dept se aplica pe DenumireCatedra a profesorului din nomenclatorul HR,
        // NU pe catedra specializarii studentilor.
        // =====================================================================
        private const string BaseDataSql = @"
            WITH BaseData AS (
                SELECT
                    vcm.NumeIntregProfesor                                          AS NumeIntreg,
                    vcm.ID_Profesor                                                  AS ID_Profesor,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                        CASE WHEN CHARINDEX('+',vcm.DenumireSpecializare)>0
                             THEN LEFT(vcm.DenumireSpecializare,CHARINDEX('+',vcm.DenumireSpecializare)-1)
                             ELSE vcm.DenumireSpecializare END,
                    'S','S'),'T','T')))) COLLATE DATABASE_DEFAULT                   AS SpecializareCurata,
                    vcm.DenumireSpecializare                                         AS NumeSpecOriginal,
                    ISNULL(vcm.DenumireMaterie,'Nedefinit')                          AS DenumireMaterie,
                    -- TipPost: pastram valoarea originala normalizata - NU suprascriem Titular cu Suplinitor!
                    -- 'TIT'/'Tit'/'Titular'/'Titulara' -> 'Titular'
                    -- 'SUP'/'Sup'/'Suplinitor'/'SupTit'/NULL/gol -> 'Suplinitor'
                    CASE
                        WHEN UPPER(LTRIM(RTRIM(ISNULL(sf.DenTitularSauSuplinitor,'')))) IN ('TIT','TITULAR','TITULARA')
                             THEN 'Titular'
                        ELSE 'Suplinitor'
                    END                                                              AS TipPost,
                    ISNULL(vcm.NrSemestruDinAn,0)                                   AS Semestru,
                    CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4))         AS OreConvLinie,
                    ISNULL(vcm.Nr_Ore_Curs,0)                                       AS OreCursLinie,
                    ISNULL(vcm.Nr_Ore_Seminar,0)+ISNULL(vcm.Nr_Ore_Laborator,0)+ISNULL(vcm.Nr_Ore_Proiect,0) AS OreAplicatiiLinie,
                    LTRIM(RTRIM(ISNULL(vcm.DenumireFacultate,'')))                  AS FacultateCurata,
                    vcm.StatDeFunctiiID_Catedra                                      AS ID_Catedra,
                    UPPER(LTRIM(RTRIM(au.Denumire))) COLLATE DATABASE_DEFAULT        AS AnCurat,
                    CASE
                        WHEN vcm.DenumireSpecializare LIKE '%-IFR%' OR vcm.DenumireSpecializare LIKE '% IFR%' OR vcm.DenumireSpecializare LIKE '%IFR' THEN 'IFR'
                        WHEN vcm.DenumireSpecializare LIKE '%-ID%'  OR vcm.DenumireSpecializare LIKE '% ID%'  OR vcm.DenumireSpecializare LIKE '%- ID' THEN 'ID'
                        ELSE 'IF'
                    END                                                              AS FormaInv,
                    vcm.ID_StatDeFunctii                                             AS ID_StatDeFunctii
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv=au.ID_AnUniv
                -- JOIN pe StatDeFunctiiPeSpecializare FARA MAX pe DenTitularSauSuplinitor
                -- Fiecare rand Tit si Sup raman separate => nu se colapseaza niciodata
                -- Daca acelasi ID_StatDeFunctii are atat Tit cat si Sup, ambele raman
                LEFT JOIN (
                    SELECT ID_StatDeFunctii, ID_AnUniv, DenumireSpecializare, DenumireMaterie,
                           NrSemestruDinAn, DenTitularSauSuplinitor
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare]
                ) sf
                    ON sf.ID_StatDeFunctii=vcm.ID_StatDeFunctii AND sf.ID_AnUniv=vcm.ID_AnUniv
                    AND sf.DenumireSpecializare=vcm.DenumireSpecializare
                    AND sf.DenumireMaterie=vcm.DenumireMaterie AND sf.NrSemestruDinAn=vcm.NrSemestruDinAn
                LEFT JOIN (
                    SELECT vp.ID_Profesor, vp.DenumireCatedra AS DeptProfesor
                    FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] vp
                    WHERE vp.ID_AnUnivCatedra = 45
                      AND vp.Ordine = (
                          SELECT MAX(vp2.Ordine) FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] vp2
                          WHERE vp2.ID_Profesor=vp.ID_Profesor AND vp2.ID_AnUnivCatedra=45
                      )
                ) profDept ON profDept.ID_Profesor=vcm.ID_Profesor
                WHERE (@dept='Toti' OR profDept.DeptProfesor COLLATE Latin1_General_CI_AI = @dept COLLATE Latin1_General_CI_AI)
            )";

        private const string FiltreStandard = @"(@an='Toti' OR bd.AnCurat=@an)
                             AND (@fac='Toti' OR bd.FacultateCurata=@fac)
                             AND (@prof='Toti' OR bd.NumeIntreg=@prof)
                             AND (@formaInv='Toti' OR bd.NumeSpecOriginal LIKE '% '+@formaInv+'%' OR bd.NumeSpecOriginal LIKE '%-'+@formaInv+'%')
                             AND (@specs='Toti' OR bd.SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
                             AND (@semestru=0 OR bd.Semestru=@semestru)
                             AND (@tipPost='Toti' OR bd.TipPost=@tipPost)";

        #endregion

        #region ================= RAPORT 1: NORMA PROFESORI =================

        // =====================================================================
        // ARHITECTURA FINALA - 3 CTE-uri:
        // 1. Filtrat: Filtrare date brute. TipPost vine din BaseData (normalizat corect).
        //    IMPORTANT: NU se mai face nicio conversie suplimentara a TipPost aici!
        //    Titular ramane Titular, Suplinitor ramane Suplinitor.
        // 2. Agregat: GROUP BY (Profesor, Spec, Materie, TipPost, Sem)
        //    => TipPost ESTE in GROUP BY => randurile Titular si Suplinitor raman SEPARATE
        // 3. Cuplaje: GROUP BY (Profesor, Materie, TipPost, Sem)
        //    => TipPost ESTE in GROUP BY => cuplajele se calculeaza SEPARAT per tip post
        // =====================================================================
        private string BuildNormaSql()
        {
            return BaseDataSql + @",
            Filtrat AS (
                SELECT bd.NumeIntreg, bd.SpecializareCurata, bd.NumeSpecOriginal,
                       bd.DenumireMaterie,
                       bd.TipPost,  -- NU modificam TipPost! Vine corect din BaseData.
                       bd.Semestru, bd.ID_Catedra, bd.ID_StatDeFunctii,
                       bd.OreCursLinie, bd.OreAplicatiiLinie, bd.OreConvLinie
                FROM BaseData bd
                WHERE (@an='Toti' OR bd.AnCurat=@an)
                  AND (@fac='Toti' OR bd.FacultateCurata=@fac)
                  AND (@prof='Toti' OR bd.NumeIntreg=@prof)
                  AND (@formaInv='Toti' OR bd.NumeSpecOriginal LIKE '% '+@formaInv+'%' OR bd.NumeSpecOriginal LIKE '%-'+@formaInv+'%')
                  AND (@specs='Toti' OR bd.SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
                  AND (@semestru=0 OR bd.Semestru=@semestru)
                  AND (@tipPost='Toti' OR bd.TipPost=@tipPost)
            ),
            Agregat AS (
                -- TipPost ESTE in GROUP BY: randurile Titular si Suplinitor sunt complet separate
                -- OreCurs: MAX (cursul comun nu se inmulteste pe grupe)
                -- OreAplicatii: SUM (laborator pe grupe diferite se aduna)
                -- OreConv: MAX (conventionalele nu se inmultesc)
                SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru,
                       MAX(ID_Catedra)        AS ID_Catedra,
                       MAX(ID_StatDeFunctii)  AS ID_StatDeFunctii,
                       MAX(OreCursLinie)      AS OreCurs,
                       SUM(OreAplicatiiLinie) AS OreAplicatii,
                       MAX(OreConvLinie)      AS OreConv
                FROM Filtrat
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru
            ),
            Cuplaje AS (
                -- Cuplajele se calculeaza SEPARAT per TipPost
                -- O materie predata ca Titular la spec A+B = 1 rand Titular cu SpecPrimara=A
                -- Aceeasi materie predata ca Suplinitor la spec C = 1 rand Suplinitor separat
                SELECT a.NumeIntreg, a.DenumireMaterie, a.TipPost, a.Semestru,
                       COUNT(DISTINCT a.SpecializareCurata) AS NrSpec,
                       MIN(a.SpecializareCurata)            AS SpecPrimara,
                       SUM(a.OreAplicatii)                  AS OreAplicatiiMax,
                       MAX(a.OreConv)                       AS OreConvMax
                FROM Agregat a
                GROUP BY a.NumeIntreg, a.DenumireMaterie, a.TipPost, a.Semestru
            )
            SELECT c.NumeIntreg AS Profesor, c.SpecPrimara AS Specializare,
                   c.DenumireMaterie AS Materie, c.TipPost, c.Semestru,
                   ag.ID_Catedra,
                   ag.OreCurs          AS NrOreCurs,
                   c.OreAplicatiiMax   AS NrOreAplicatii,
                   c.OreConvMax        AS NrOreConventionale,
                   c.NrSpec,
                   CASE WHEN c.NrSpec > 1 THEN (
                       SELECT TOP 1 STUFF((
                           SELECT ', ' + a2.SpecializareCurata
                           FROM Agregat a2
                           WHERE a2.NumeIntreg=c.NumeIntreg AND a2.DenumireMaterie=c.DenumireMaterie
                             AND a2.TipPost=c.TipPost AND a2.Semestru=c.Semestru
                             AND a2.SpecializareCurata <> c.SpecPrimara
                           ORDER BY a2.SpecializareCurata
                           FOR XML PATH(''),TYPE
                       ).value('.','NVARCHAR(MAX)'),1,2,'')
                       FROM Agregat WHERE NumeIntreg=c.NumeIntreg
                   ) ELSE '' END AS SpecCuplate,
                   SUM(c.OreConvMax) OVER(PARTITION BY c.NumeIntreg, c.TipPost) AS TotalTipPost,
                   SUM(c.OreConvMax) OVER(PARTITION BY c.NumeIntreg)            AS TotalPost
            FROM Cuplaje c
            INNER JOIN Agregat ag
                ON ag.NumeIntreg=c.NumeIntreg AND ag.DenumireMaterie=c.DenumireMaterie
                AND ag.TipPost=c.TipPost AND ag.Semestru=c.Semestru
                AND ag.SpecializareCurata=c.SpecPrimara
            ORDER BY c.NumeIntreg, c.TipPost DESC, c.DenumireMaterie
            OPTION (RECOMPILE)";
        }

        [HttpGet("norma-profesori")]
        public ActionResult GetNormaProfesori(string? anUniv, string? facultate, string? specializari,
            string? profesor, int semestru = 0, string tipPost = "Toti",
            string? formaInvatamant = "Toti", string? departament = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(BuildNormaSql(), conn);
            cmd.CommandTimeout = 300;
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
                    NrOreCurs = reader["NrOreCurs"] != DBNull.Value ? Convert.ToDouble(reader["NrOreCurs"]) : 0.0,
                    NrOreAplicatii = reader["NrOreAplicatii"] != DBNull.Value ? Convert.ToDouble(reader["NrOreAplicatii"]) : 0.0,
                    NrOreConventionale = reader["NrOreConventionale"] != DBNull.Value ? Convert.ToDouble(reader["NrOreConventionale"]) : 0.0,
                    Mentiuni = reader["SpecCuplate"] != DBNull.Value && !string.IsNullOrEmpty(reader["SpecCuplate"]?.ToString())
                                         ? "Cuplaj cu: " + reader["SpecCuplate"]?.ToString() : "",
                    TotalTipPost = reader["TotalTipPost"] != DBNull.Value ? Convert.ToDouble(reader["TotalTipPost"]) : 0.0,
                    TotalPost = reader["TotalPost"] != DBNull.Value ? Convert.ToDouble(reader["TotalPost"]) : 0.0
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
                new DataColumn("Nr Ore Conventionale", typeof(double)), new DataColumn("Mentiuni")
            });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(BuildNormaSql(), conn);
            cmd.CommandTimeout = 300;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dt.Rows.Add(reader["Profesor"], reader["Specializare"], reader["Materie"],
                    GetDenumireCatedra(idCat), reader["TipPost"], reader["Semestru"],
                    reader["NrOreCurs"] != DBNull.Value ? Convert.ToDouble(reader["NrOreCurs"]) : 0.0,
                    reader["NrOreAplicatii"] != DBNull.Value ? Convert.ToDouble(reader["NrOreAplicatii"]) : 0.0,
                    reader["NrOreConventionale"] != DBNull.Value ? Convert.ToDouble(reader["NrOreConventionale"]) : 0.0,
                    reader["SpecCuplate"] != DBNull.Value && !string.IsNullOrEmpty(reader["SpecCuplate"]?.ToString())
                        ? "Cuplaj cu: " + reader["SpecCuplate"]?.ToString() : "");
            }
            string fileName = string.IsNullOrEmpty(profesor) || profesor == "Toti"
                ? "NormaProfesori_General.xlsx"
                : $"NormaProfesori_{string.Join("_", profesor.Split(Path.GetInvalidFileNameChars()))}.xlsx";
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme");
            ws.Cell(1, 1).Value = "Filtre Aplicate"; ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
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
            FiltratRaw AS (
                SELECT bd.NumeIntreg, bd.SpecializareCurata, bd.DenumireMaterie,
                       bd.TipPost,  -- TipPost pastrat din BaseData, nu modificat
                       bd.Semestru, bd.ID_StatDeFunctii,
                       bd.OreCursLinie, bd.OreAplicatiiLinie, bd.OreConvLinie
                FROM BaseData bd
                WHERE (@an='Toti' OR bd.AnCurat=@an) AND (@fac='Toti' OR bd.FacultateCurata=@fac)
                  AND (@prof='Toti' OR bd.NumeIntreg=@prof)
                  AND (@formaInv='Toti' OR bd.NumeSpecOriginal LIKE '% '+@formaInv+'%' OR bd.NumeSpecOriginal LIKE '%-'+@formaInv+'%')
                  AND (@specs='Toti' OR bd.SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
                  AND (@semestru=0 OR bd.Semestru=@semestru)
                  AND (@tipPost='Toti' OR bd.TipPost=@tipPost)
            ),
            Agregat2 AS (
                -- TipPost in GROUP BY: Titular si Suplinitor raman separate
                SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru,
                       MAX(OreCursLinie)      AS OreCurs,
                       SUM(OreAplicatiiLinie) AS OreAplicatii,
                       MAX(OreConvLinie)      AS OreConv
                FROM FiltratRaw
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru
            ),
            CuplajeProg AS (
                -- Cuplaje SEPARATE pe TipPost
                SELECT NumeIntreg, DenumireMaterie, TipPost, Semestru,
                       MIN(SpecializareCurata) AS SpecPrimara,
                       SUM(OreAplicatii)  AS OreAplicatiiMax,
                       MAX(OreConv)       AS OreConvMax
                FROM Agregat2
                GROUP BY NumeIntreg, DenumireMaterie, TipPost, Semestru
            ),
            Dedup AS (
                SELECT cp.NumeIntreg, cp.SpecPrimara AS SpecializareCurata,
                       ag.OreCurs AS OreCurs,
                       cp.OreAplicatiiMax AS OreAplicatii,
                       cp.OreConvMax AS OreConv
                FROM CuplajeProg cp
                INNER JOIN Agregat2 ag
                    ON ag.NumeIntreg=cp.NumeIntreg AND ag.DenumireMaterie=cp.DenumireMaterie
                    AND ag.TipPost=cp.TipPost AND ag.Semestru=cp.Semestru
                    AND ag.SpecializareCurata=cp.SpecPrimara
            ),
            Filtrat AS (
                SELECT NumeIntreg AS Profesor, SpecializareCurata AS ProgramStudiu,
                       SUM(OreConv)       AS OreConvProgram,
                       SUM(OreCurs)       AS OreCursProgram,
                       SUM(OreAplicatii)  AS OreAplicatiiProgram
                FROM Dedup
                GROUP BY NumeIntreg, SpecializareCurata
                HAVING SUM(OreConv)>0
            ),
            TotalProfesor AS (SELECT Profesor, SUM(OreConvProgram) AS TotalPost FROM Filtrat GROUP BY Profesor)
            SELECT f.Profesor, ISNULL(f.ProgramStudiu,'Nespecificat') AS ProgramStudiu,
                   ISNULL(f.OreConvProgram,0) AS NrOreConv,
                   ISNULL(f.OreCursProgram,0) AS NrOreCurs,
                   ISNULL(f.OreAplicatiiProgram,0) AS NrOreAplicatii,
                   ISNULL(t.TotalPost,0) AS TotalPost,
                   CAST(CASE WHEN ISNULL(t.TotalPost,0)=0 THEN 0 ELSE (ISNULL(f.OreConvProgram,0)/t.TotalPost)*100 END AS DECIMAL(10,2)) AS ProcentPost
            FROM Filtrat f INNER JOIN TotalProfesor t ON f.Profesor=t.Profesor
            ORDER BY f.Profesor, f.OreConvProgram DESC";

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 120;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                result.Add(new
                {
                    Profesor = reader["Profesor"]?.ToString() ?? "",
                    ProgramStudiu = reader["ProgramStudiu"]?.ToString() ?? "",
                    NrOreConv = reader["NrOreConv"] != DBNull.Value ? Convert.ToDouble(reader["NrOreConv"]) : 0.0,
                    NrOreCurs = reader["NrOreCurs"] != DBNull.Value ? Convert.ToDouble(reader["NrOreCurs"]) : 0.0,
                    NrOreAplicatii = reader["NrOreAplicatii"] != DBNull.Value ? Convert.ToDouble(reader["NrOreAplicatii"]) : 0.0,
                    TotalPost = reader["TotalPost"] != DBNull.Value ? Convert.ToDouble(reader["TotalPost"]) : 0.0,
                    ProcentPost = reader["ProcentPost"] != DBNull.Value ? Convert.ToDouble(reader["ProcentPost"]) : 0.0
                });
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
                new DataColumn("Nr Ore Curs", typeof(double)), new DataColumn("Nr Ore Aplicatii", typeof(double)),
                new DataColumn("Nr Ore Conv", typeof(double)), new DataColumn("Procent Post", typeof(double))
            });
            string sql = BaseDataSql + @",
            FiltratRaw AS (
                SELECT bd.NumeIntreg, bd.SpecializareCurata, bd.DenumireMaterie,
                       bd.TipPost, bd.Semestru, bd.ID_StatDeFunctii,
                       bd.OreCursLinie, bd.OreAplicatiiLinie, bd.OreConvLinie
                FROM BaseData bd
                WHERE (@an='Toti' OR bd.AnCurat=@an) AND (@fac='Toti' OR bd.FacultateCurata=@fac)
                  AND (@prof='Toti' OR bd.NumeIntreg=@prof)
                  AND (@formaInv='Toti' OR bd.NumeSpecOriginal LIKE '% '+@formaInv+'%' OR bd.NumeSpecOriginal LIKE '%-'+@formaInv+'%')
                  AND (@specs='Toti' OR bd.SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
                  AND (@semestru=0 OR bd.Semestru=@semestru)
                  AND (@tipPost='Toti' OR bd.TipPost=@tipPost)
            ),
            Agregat2 AS (
                SELECT NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru,
                       MAX(OreCursLinie) AS OreCurs, SUM(OreAplicatiiLinie) AS OreAplicatii, MAX(OreConvLinie) AS OreConv
                FROM FiltratRaw
                GROUP BY NumeIntreg, SpecializareCurata, DenumireMaterie, TipPost, Semestru
            ),
            CuplajeProg AS (
                SELECT NumeIntreg, DenumireMaterie, TipPost, Semestru,
                       MIN(SpecializareCurata) AS SpecPrimara,
                       SUM(OreAplicatii) AS OreAplicatiiMax, MAX(OreConv) AS OreConvMax
                FROM Agregat2 GROUP BY NumeIntreg, DenumireMaterie, TipPost, Semestru
            ),
            Dedup AS (
                SELECT cp.NumeIntreg, cp.SpecPrimara AS SpecializareCurata,
                       ag.OreCurs AS OreCurs, cp.OreAplicatiiMax AS OreAplicatii, cp.OreConvMax AS OreConv
                FROM CuplajeProg cp
                INNER JOIN Agregat2 ag
                    ON ag.NumeIntreg=cp.NumeIntreg AND ag.DenumireMaterie=cp.DenumireMaterie
                    AND ag.TipPost=cp.TipPost AND ag.Semestru=cp.Semestru
                    AND ag.SpecializareCurata=cp.SpecPrimara
            ),
            Filtrat AS (
                SELECT NumeIntreg, SpecializareCurata AS ProgramStudiu,
                       SUM(OreConv) AS OreConvProgram, SUM(OreCurs) AS OreCursProgram,
                       SUM(OreAplicatii) AS OreAplicatiiProgram
                FROM Dedup GROUP BY NumeIntreg, SpecializareCurata HAVING SUM(OreConv)>0
            ),
            TotalProfesor AS (SELECT NumeIntreg, SUM(OreConvProgram) AS TotalPost FROM Filtrat GROUP BY NumeIntreg)
            SELECT f.NumeIntreg, ISNULL(f.ProgramStudiu,'Nespecificat') AS ProgramStudiu,
                   ISNULL(f.OreCursProgram,0) AS NrOreCurs, ISNULL(f.OreAplicatiiProgram,0) AS NrOreAplicatii,
                   ISNULL(f.OreConvProgram,0) AS OreConvProgram,
                   CAST(CASE WHEN ISNULL(t.TotalPost,0)=0 THEN 0 ELSE (ISNULL(f.OreConvProgram,0)/t.TotalPost)*100 END AS DECIMAL(10,2)) AS ProcentPost
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
                    reader["NrOreCurs"] != DBNull.Value ? Convert.ToDouble(reader["NrOreCurs"]) : 0.0,
                    reader["NrOreAplicatii"] != DBNull.Value ? Convert.ToDouble(reader["NrOreAplicatii"]) : 0.0,
                    reader["OreConvProgram"] != DBNull.Value ? Convert.ToDouble(reader["OreConvProgram"]) : 0.0,
                    reader["ProcentPost"] != DBNull.Value ? Convert.ToDouble(reader["ProcentPost"]) : 0.0);

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

        // =====================================================================
        // FIX COMPLET R3:
        // - TipPost vine din DenTitularSauSuplinitor (nu din bit TitularSauSuplinitor)
        // - 'TIT'/'Titular'/'Titulara' -> 'Titular', orice altceva -> 'Suplinitor'
        // - Titular si Suplinitor raman pe randuri SEPARATE (TipPost in GROUP BY)
        // - Dedup elimina grupe multiple per (Profesor, TipPost, FormaInv, Materie, Sem)
        // - Agregare finala: max 2 randuri per profesor (Titular + Suplinitor)
        // =====================================================================
        private string BuildNormaTotaluriSql()
        {
            return @"
            WITH DateBrute AS (
                SELECT
                    vcm.NumeIntregProfesor                                              AS NumeComplet,
                    vcm.ID_Profesor,
                    vcm.StatDeFunctiiID_Catedra                                         AS ID_Catedra,
                    LTRIM(RTRIM(ISNULL(vcm.DenumireFacultate,'')))                      AS Facultate,
                    CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4))             AS OreConv,
                    ISNULL(vcm.DenumireMaterie,'Nedefinit')                              AS DenumireMaterie,
                    ISNULL(vcm.NrSemestruDinAn,0)                                       AS Semestru,
                    vcm.ID_StatDeFunctii,
                    -- TipPost din StatDeFunctiiPeSpecializare (sursa corecta pentru norma)
                    -- Titular si Suplinitor RAMAN SEPARATE
                    CASE
                        WHEN UPPER(LTRIM(RTRIM(ISNULL(sf.DenTitularSauSuplinitor,'')))) IN ('TIT','TITULAR','TITULARA')
                             THEN 'Titular'
                        ELSE 'Suplinitor'
                    END AS TipPost,
                    CASE
                        WHEN vcm.DenumireSpecializare LIKE '%-IFR%' OR vcm.DenumireSpecializare LIKE '% IFR%' THEN 'IFR'
                        WHEN vcm.DenumireSpecializare LIKE '%-ID%'  OR vcm.DenumireSpecializare LIKE '% ID%'  THEN 'ID'
                        ELSE 'IF'
                    END AS FormaInv
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv=au.ID_AnUniv
                LEFT JOIN (
                    SELECT ID_StatDeFunctii, ID_AnUniv, DenumireSpecializare, DenumireMaterie,
                           NrSemestruDinAn, DenTitularSauSuplinitor
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare]
                ) sf
                    ON sf.ID_StatDeFunctii=vcm.ID_StatDeFunctii AND sf.ID_AnUniv=vcm.ID_AnUniv
                    AND sf.DenumireSpecializare=vcm.DenumireSpecializare
                    AND sf.DenumireMaterie=vcm.DenumireMaterie AND sf.NrSemestruDinAn=vcm.NrSemestruDinAn
                WHERE (@an='Toti' OR UPPER(LTRIM(RTRIM(au.Denumire)))=@an)
                  AND (@fac='Toti' OR LTRIM(RTRIM(vcm.DenumireFacultate)) COLLATE Latin1_General_CI_AI=@fac COLLATE Latin1_General_CI_AI)
                  AND (@prof='Toti' OR vcm.NumeIntregProfesor=@prof)
            ),
            -- Dedup: elimina grupe multiple per materie
            -- TipPost in GROUP BY: Titular si Suplinitor raman separate!
            Dedup AS (
                SELECT NumeComplet, ID_Profesor, TipPost, FormaInv,
                       DenumireMaterie, Semestru,
                       MAX(OreConv)    AS OreConvDedup,
                       MAX(ID_Catedra) AS ID_Catedra,
                       MAX(Facultate)  AS Facultate
                FROM DateBrute
                GROUP BY NumeComplet, ID_Profesor, TipPost, FormaInv,
                         DenumireMaterie, Semestru
            ),
            -- Agregare finala: max 2 randuri per profesor (Titular + Suplinitor)
            Agreg AS (
                SELECT NumeComplet, ID_Profesor,
                       MAX(ID_Catedra) AS ID_Catedra,
                       MAX(Facultate)  AS Facultate,
                       TipPost,
                       CAST(ISNULL(SUM(CASE WHEN FormaInv='IF'  THEN OreConvDedup ELSE 0 END),0) AS DECIMAL(10,2)) AS OreIF,
                       CAST(ISNULL(SUM(CASE WHEN FormaInv='ID'  THEN OreConvDedup ELSE 0 END),0) AS DECIMAL(10,2)) AS OreID,
                       CAST(ISNULL(SUM(CASE WHEN FormaInv='IFR' THEN OreConvDedup ELSE 0 END),0) AS DECIMAL(10,2)) AS OreIFR,
                       CAST(ISNULL(SUM(OreConvDedup),0) AS DECIMAL(10,2))                                           AS TotalOreConv
                FROM Dedup
                GROUP BY NumeComplet, ID_Profesor, TipPost
            )
            SELECT NumeComplet, ID_Profesor, ID_Catedra, Facultate, TipPost,
                   OreIF, OreID, OreIFR, TotalOreConv,
                   CAST(TotalOreConv*14 AS DECIMAL(10,2)) AS TotalAnual
            FROM Agreg
            ORDER BY NumeComplet, TipPost DESC";
        }

        [HttpGet("norma-totaluri")]
        public ActionResult GetNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(BuildNormaTotaluriSql(), conn);
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
                    Profesor = FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    Departament = GetDenumireCatedra(idCat),
                    Facultate = reader["Facultate"]?.ToString() ?? "",
                    TipNorma = reader["TipPost"]?.ToString() ?? "",
                    OreIF = reader["OreIF"] != DBNull.Value ? Convert.ToDecimal(reader["OreIF"]) : 0m,
                    OreID = reader["OreID"] != DBNull.Value ? Convert.ToDecimal(reader["OreID"]) : 0m,
                    OreIFR = reader["OreIFR"] != DBNull.Value ? Convert.ToDecimal(reader["OreIFR"]) : 0m,
                    TotalOreConv = reader["TotalOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalOreConv"]) : 0m,
                    TotalAnualOreConv = reader["TotalAnual"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAnual"]) : 0m
                });
            }
            return Ok(result);
        }

        [HttpGet("export/norma-totaluri")]
        public IActionResult ExportNormaTotaluri(string? anUniv, string? facultate, string? departament, string? profesor)
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[] {
                new DataColumn("Profesor"), new DataColumn("Departament"), new DataColumn("Facultate"),
                new DataColumn("Tip Norma"),
                new DataColumn("Ore IF",  typeof(decimal)), new DataColumn("Ore ID",  typeof(decimal)),
                new DataColumn("Ore IFR", typeof(decimal)),
                new DataColumn("Total Ore Conv.", typeof(decimal)),
                new DataColumn("Total Anual (x14)", typeof(decimal))
            });
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(BuildNormaTotaluriSql(), conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                dt.Rows.Add(
                    FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    GetDenumireCatedra(idCat), reader["Facultate"], reader["TipPost"],
                    reader["OreIF"] != DBNull.Value ? Convert.ToDecimal(reader["OreIF"]) : 0m,
                    reader["OreID"] != DBNull.Value ? Convert.ToDecimal(reader["OreID"]) : 0m,
                    reader["OreIFR"] != DBNull.Value ? Convert.ToDecimal(reader["OreIFR"]) : 0m,
                    reader["TotalOreConv"] != DBNull.Value ? Convert.ToDecimal(reader["TotalOreConv"]) : 0m,
                    reader["TotalAnual"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAnual"]) : 0m);
            }
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Totaluri Norme");
            var tbl = ws.Cell(1, 1).InsertTable(dt); tbl.Theme = XLTableTheme.None;
            tbl.ShowTotalsRow = true;
            tbl.Field("Ore IF").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore ID").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore IFR").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Total Anual (x14)").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Profesor").TotalsRowLabel = "TOTAL GENERAL";
            ws.Columns().AdjustToContents();
            StyleHeader(ws.Range(1, 1, 1, dt.Columns.Count));
            using var stream = new MemoryStream(); wb.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Totaluri_Norme.xlsx");
        }

        #endregion

        #region ================= RAPORT 4: LIMBI STRAINE =================

        // =====================================================================
        // FIX LIMBI STRAINE:
        // - Datele brute includ AMBELE tipuri de norme (Titular + Suplinitor)
        // - Nu se mai elimina nicio forma de incadrare
        // - Deduplicare corecta a cuplajelor: per (Profesor, Materie, Sem, TipPost) -> MAX
        // - Totalul = suma TUTUROR orelor predate, indiferent de tipul normei
        // =====================================================================
        private string BuildLimbiSql()
        {
            return @"
            WITH DateLimbi AS (
                SELECT
                    vcm.NumeIntregProfesor                                             AS NumeComplet,
                    ISNULL(vcm.DenumireMaterie,'Nedefinit')                            AS DenumireMaterie,
                    ISNULL(vcm.NrSemestruDinAn,0)                                      AS Semestru,
                    vcm.ID_StatDeFunctii,
                    -- TipPost normalizat: 'TIT'/'Titular'/'Titulara' -> 'Titular', rest -> 'Suplinitor'
                    CASE
                        WHEN UPPER(LTRIM(RTRIM(ISNULL(sf.DenTitularSauSuplinitor,'')))) IN ('TIT','TITULAR','TITULARA')
                             THEN 'Titular'
                        ELSE 'Suplinitor'
                    END AS TipPost,
                    CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4))            AS OreConv,
                    UPPER(LTRIM(RTRIM(au.Denumire)))                                   AS AnCurat,
                    LTRIM(RTRIM(ISNULL(vcm.DenumireFacultate,'')))                     AS FacultateCurata,
                    vcm.DenumireSpecializare                                            AS NumeSpecOriginal,
                    UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                        CASE WHEN CHARINDEX('+',vcm.DenumireSpecializare)>0
                             THEN LEFT(vcm.DenumireSpecializare,CHARINDEX('+',vcm.DenumireSpecializare)-1)
                             ELSE vcm.DenumireSpecializare END,'S','S'),'T','T')))) AS SpecializareCurata
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN [AGSIS].[dbo].[AnUniversitar] au ON vcm.ID_AnUniv=au.ID_AnUniv
                -- JOIN fara MAX: fiecare rand Tit/Sup raman separate
                LEFT JOIN (
                    SELECT ID_StatDeFunctii, ID_AnUniv, DenumireSpecializare, DenumireMaterie,
                           NrSemestruDinAn, DenTitularSauSuplinitor
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare]
                ) sf
                    ON sf.ID_StatDeFunctii=vcm.ID_StatDeFunctii AND sf.ID_AnUniv=vcm.ID_AnUniv
                    AND sf.DenumireSpecializare=vcm.DenumireSpecializare
                    AND sf.DenumireMaterie=vcm.DenumireMaterie AND sf.NrSemestruDinAn=vcm.NrSemestruDinAn
                WHERE vcm.ID_AnUniv = 45  -- FILTRU OBLIGATORIU: doar anul universitar 2025-2026
                  AND (@fac='Toti' OR LTRIM(RTRIM(vcm.DenumireFacultate)) COLLATE Latin1_General_CI_AI=@fac COLLATE Latin1_General_CI_AI)
                  AND (@prof='Toti' OR vcm.NumeIntregProfesor=@prof)
                  AND (@formaInv='Toti' OR vcm.DenumireSpecializare LIKE '% '+@formaInv+'%' OR vcm.DenumireSpecializare LIKE '%-'+@formaInv+'%')
                  AND (@specs='Toti' OR UPPER(LTRIM(RTRIM(REPLACE(REPLACE(
                        CASE WHEN CHARINDEX('+',vcm.DenumireSpecializare)>0
                             THEN LEFT(vcm.DenumireSpecializare,CHARINDEX('+',vcm.DenumireSpecializare)-1)
                             ELSE vcm.DenumireSpecializare END,'S','S'),'T','T'))))
                       IN (SELECT value FROM STRING_SPLIT(@specs,',')))
                  AND (@semestru=0 OR vcm.NrSemestruDinAn=@semestru)
                  AND (@tipPost='Toti' OR
                       CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(sf.DenTitularSauSuplinitor,'')))) IN ('TIT','TITULAR','TITULARA')
                            THEN 'Titular' ELSE 'Suplinitor' END = @tipPost)
                  AND (vcm.DenumireSpecializare LIKE '%englez%' OR vcm.DenumireSpecializare LIKE '%francez%'
                    OR vcm.DenumireSpecializare LIKE '%german%' OR vcm.DenumireSpecializare LIKE '%american%'
                    OR vcm.DenumireSpecializare LIKE '%(EN)%'   OR vcm.DenumireSpecializare LIKE '%(FR)%'
                    OR vcm.DenumireSpecializare LIKE '%(G)%'
                    OR vcm.DenumireSpecializare IN (
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
                        'Dezvoltarea afacerilor turistice','Medicina traditionala chineza'))
            ),
            -- Pre-agregare per (Profesor, Spec, Materie, TipPost, Sem)
            -- TipPost in GROUP BY: Titular si Suplinitor raman separate
            PreAgreg AS (
                SELECT NumeComplet, SpecializareCurata, DenumireMaterie, Semestru, TipPost,
                       MAX(OreConv) AS OreConvSpec
                FROM DateLimbi
                GROUP BY NumeComplet, SpecializareCurata, DenumireMaterie, Semestru, TipPost
            ),
            -- Dedup cuplaje: per (Profesor, Materie, Sem, TipPost) -> MAX ore per cuplaj
            Dedup AS (
                SELECT NumeComplet, DenumireMaterie, Semestru, TipPost,
                       MAX(OreConvSpec) AS OreConvDedup
                FROM PreAgreg
                GROUP BY NumeComplet, DenumireMaterie, Semestru, TipPost
            )
            -- Total = SUMA tuturor orelor (Titular + Suplinitor) per profesor
            SELECT NumeComplet,
                   CAST(ROUND(SUM(CASE WHEN Semestru IN (1,3,5,7,9,11) THEN OreConvDedup ELSE 0 END)*14,2) AS DECIMAL(10,2)) AS Sem1,
                   CAST(ROUND(SUM(CASE WHEN Semestru IN (2,4,6,8,10,12) THEN OreConvDedup ELSE 0 END)*14,2) AS DECIMAL(10,2)) AS Sem2,
                   CAST(ROUND(SUM(OreConvDedup)*14,2) AS DECIMAL(10,2)) AS Total
            FROM Dedup
            GROUP BY NumeComplet
            HAVING SUM(OreConvDedup)>0
            ORDER BY NumeComplet";
        }

        [HttpGet("limbi-straine")]
        public ActionResult GetLimbiStraine(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(BuildLimbiSql(), conn);
            cmd.CommandTimeout = 120;
            AddLimbiParams(cmd, anUniv, facultate, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            int nr = 1;
            while (reader.Read())
                result.Add(new
                {
                    NrCrt = nr++,
                    Profesor = reader["NumeComplet"]?.ToString() ?? "",
                    TotalSem1 = reader["Sem1"] != DBNull.Value ? Convert.ToDecimal(reader["Sem1"]) : 0m,
                    TotalSem2 = reader["Sem2"] != DBNull.Value ? Convert.ToDecimal(reader["Sem2"]) : 0m,
                    Total = reader["Total"] != DBNull.Value ? Convert.ToDecimal(reader["Total"]) : 0m
                });
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
            conn.Open();
            using var cmd = new SqlCommand(BuildLimbiSql(), conn);
            cmd.CommandTimeout = 120;
            AddLimbiParams(cmd, anUniv, facultate, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            int nr = 1;
            while (reader.Read())
                dt.Rows.Add(nr++, reader["NumeComplet"],
                    reader["Sem1"] != DBNull.Value ? Convert.ToDecimal(reader["Sem1"]) : 0m,
                    reader["Sem2"] != DBNull.Value ? Convert.ToDecimal(reader["Sem2"]) : 0m,
                    reader["Total"] != DBNull.Value ? Convert.ToDecimal(reader["Total"]) : 0m);

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

        private void AddLimbiParams(SqlCommand cmd, string? anUniv, string? facultate,
            string? formaInvatamant, string? profesor, string? specializari, int semestru, string? tipPost)
        {
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            cmd.Parameters.AddWithValue("@formaInv", formaInvatamant ?? "Toti");
            cmd.Parameters.AddWithValue("@prof", profesor ?? "Toti");
            cmd.Parameters.AddWithValue("@specs", string.IsNullOrWhiteSpace(specializari) ? "Toti" : specializari);
            cmd.Parameters.AddWithValue("@semestru", semestru);
            cmd.Parameters.AddWithValue("@tipPost", tipPost ?? "Toti");
        }

        #endregion

        #region ================= RAPORT 5: DISCIPLINE PREDATE =================

        private string BuildDisciplineSql()
        {
            return BaseDataSql + @",
            DU AS (
                SELECT DISTINCT
                    bd.NumeIntreg, bd.ID_Catedra, bd.ID_Profesor, bd.FormaInv, bd.DenumireMaterie
                FROM BaseData bd
                WHERE (@an='Toti' OR bd.AnCurat=@an)
                  AND (@fac='Toti' OR bd.FacultateCurata=@fac)
                  AND (@prof='Toti' OR bd.NumeIntreg=@prof)
                  AND (@formaInv='Toti' OR bd.NumeSpecOriginal LIKE '% '+@formaInv+'%' OR bd.NumeSpecOriginal LIKE '%-'+@formaInv+'%')
                  AND (@specs='Toti' OR bd.SpecializareCurata IN (SELECT value FROM STRING_SPLIT(@specs,',')))
                  AND (@semestru=0 OR bd.Semestru=@semestru)
                  AND (@tipPost='Toti' OR bd.TipPost=@tipPost)
            ),
            Prof AS (
                SELECT NumeIntreg, MIN(ID_Catedra) AS ID_Catedra,
                       MAX(ID_Profesor) AS ID_Profesor, FormaInv
                FROM DU
                GROUP BY NumeIntreg, FormaInv
            )
            SELECT p.NumeIntreg, p.ID_Catedra, p.ID_Profesor, p.FormaInv,
                   STUFF((
                       SELECT DISTINCT ', ' + d2.DenumireMaterie
                       FROM DU d2
                       WHERE d2.NumeIntreg=p.NumeIntreg AND d2.FormaInv=p.FormaInv
                       ORDER BY ', ' + d2.DenumireMaterie
                       FOR XML PATH(''),TYPE
                   ).value('.','NVARCHAR(MAX)'),1,2,'') AS Discipline
            FROM Prof p
            ORDER BY p.FormaInv, p.NumeIntreg";
        }

        [HttpGet("discipline-predate")]
        public ActionResult GetDisciplinePredate(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(BuildDisciplineSql(), conn);
            cmd.CommandTimeout = 180;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                result.Add(new
                {
                    Profesor = FixNume(reader["NumeIntreg"]?.ToString(), reader["ID_Profesor"]),
                    Departament = GetDenumireCatedra(idCat),
                    FormaInvatamant = reader["FormaInv"]?.ToString() ?? "",
                    Discipline = reader["Discipline"]?.ToString() ?? ""
                });
            }
            return Ok(result);
        }

        [HttpGet("export/discipline-predate")]
        public IActionResult ExportDisciplinePredate(string? anUniv, string? facultate, string? departament,
            string? profesor, string? specializari, int semestru = 0,
            string tipPost = "Toti", string? formaInvatamant = "Toti")
        {
            var datePerForma = new Dictionary<string, DataTable>();
            foreach (var forma in new[] { "IF", "ID", "IFR" })
            {
                var dt = new DataTable();
                dt.Columns.AddRange(new[] {
                    new DataColumn("Nr.Crt.", typeof(int)), new DataColumn("Nume si prenume"),
                    new DataColumn("Departament"), new DataColumn("Discipline Predate")
                });
                datePerForma[forma] = dt;
            }
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(BuildDisciplineSql(), conn);
            cmd.CommandTimeout = 180;
            AddBaseParams(cmd, anUniv, facultate, departament, formaInvatamant, profesor, specializari, semestru, tipPost);
            using var reader = cmd.ExecuteReader();
            var nrCrt = new Dictionary<string, int> { ["IF"] = 1, ["ID"] = 1, ["IFR"] = 1 };
            while (reader.Read())
            {
                string forma = reader["FormaInv"]?.ToString() ?? "IF";
                if (!datePerForma.ContainsKey(forma)) forma = "IF";
                long idCat = reader["ID_Catedra"] != DBNull.Value ? Convert.ToInt64(reader["ID_Catedra"]) : 0;
                datePerForma[forma].Rows.Add(
                    nrCrt[forma]++,
                    FixNume(reader["NumeIntreg"]?.ToString(), reader["ID_Profesor"]),
                    GetDenumireCatedra(idCat),
                    reader["Discipline"]?.ToString() ?? "");
            }
            using var memZip = new MemoryStream();
            using (var archive = new ZipArchive(memZip, ZipArchiveMode.Create, true))
            {
                foreach (var kvp in datePerForma)
                {
                    if (kvp.Value.Rows.Count == 0) continue;
                    var entry = archive.CreateEntry($"Discipline_Predate_{kvp.Key}.xlsx");
                    using var entryStream = entry.Open();
                    using var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add($"Discipline {kvp.Key}");
                    ws.Cell(1, 1).Value = $"Discipline Predate - {kvp.Key} | An: {anUniv} | Facultate: {facultate}";
                    ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColorHex);
                    var tbl = ws.Cell(3, 1).InsertTable(kvp.Value); tbl.Theme = XLTableTheme.None;
                    ws.Columns().AdjustToContents();
                    ws.Column(4).Style.Alignment.WrapText = true; ws.Column(4).Width = 80;
                    StyleHeader(ws.Range(3, 1, 3, kvp.Value.Columns.Count));
                    using var wbStream = new MemoryStream(); wb.SaveAs(wbStream); wbStream.Position = 0; wbStream.CopyTo(entryStream);
                }
            }
            memZip.Position = 0;
            return File(memZip.ToArray(), "application/zip", "Discipline_Predate_IF_ID_IFR.zip");
        }

        #endregion

        #region ================= RAPORT 6: TITULARI =================

        // =====================================================================
        // SURSA CORECTA TITULARI:
        // View_Profesori_CF_AnUniv cu:
        //   - TitularAnUniv = 1
        //   - ID_AnUnivCatedra = 45
        //   - DidCerc = 'D' (doar cadre didactice, NU cercetatori)
        //   - NumeIntreg NOT LIKE '--%' (elimina randuri reziduale tip '--Cercetator--')
        //   - LTRIM(RTRIM(NumeIntreg)) != '' (elimina randuri cu nume gol)
        // =====================================================================
        private const string TitulariSql = @"
            WITH TitulariBase AS (
                SELECT v.ID_Profesor, v.NumeIntreg, v.DenumireCatedra, v.DenumireFacultate,
                       ROW_NUMBER() OVER(PARTITION BY v.ID_Profesor ORDER BY v.Ordine DESC) AS Rn
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] v
                WHERE v.TitularAnUniv = 1
                  AND v.ID_AnUnivCatedra = 45
                  AND v.DidCerc = 'D'
                  AND LTRIM(RTRIM(ISNULL(v.NumeIntreg,''))) != ''
                  AND v.NumeIntreg NOT LIKE '--%'
            )
            SELECT ID_Profesor, NumeIntreg AS NumeComplet, DenumireCatedra, DenumireFacultate AS Facultate
            FROM TitulariBase
            WHERE Rn = 1
              AND (@fac='Toti' OR DenumireFacultate COLLATE Latin1_General_CI_AI=@fac COLLATE Latin1_General_CI_AI)
            ORDER BY NumeIntreg";

        [HttpGet("titulari")]
        public ActionResult GetTitulari(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(TitulariSql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(new
                {
                    Profesor = FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    Departament = reader["DenumireCatedra"]?.ToString() ?? "",
                    Facultate = reader["Facultate"]?.ToString() ?? ""
                });
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
            conn.Open();
            using var cmd = new SqlCommand(TitulariSql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                dt.Rows.Add(
                    FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    reader["DenumireCatedra"]?.ToString() ?? "",
                    reader["Facultate"]?.ToString() ?? "");
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

        // =====================================================================
        // SURSA CORECTA COLABORATORI:
        // Profesori care predau in an 45 (sunt in View_CentralizareMateriiProfesor)
        // DAR nu sunt titulari didactici (nu au TitularAnUniv=1 AND DidCerc='D' AND ID_AnUnivCatedra=45)
        // Deduplicare finala garanteaza un singur rand per ID_Profesor
        // =====================================================================
        private const string ColaboratoriSql = @"
            WITH TitulariDidactici AS (
                -- Setul exact de titulari didactici (acelasi criteriu ca R6 Titulari)
                SELECT DISTINCT ID_Profesor
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv]
                WHERE TitularAnUniv = 1
                  AND ID_AnUnivCatedra = 45
                  AND DidCerc = 'D'
                  AND LTRIM(RTRIM(ISNULL(NumeIntreg,''))) != ''
                  AND NumeIntreg NOT LIKE '--%'
            ),
            ColabBase AS (
                -- Toti profesorii activi in an 45 care NU sunt titulari didactici
                SELECT DISTINCT vcm.ID_Profesor
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                WHERE vcm.ID_AnUniv = 45
                  AND NOT EXISTS (SELECT 1 FROM TitulariDidactici td WHERE td.ID_Profesor = vcm.ID_Profesor)
            ),
            NomenclatorColab AS (
                -- Cel mai recent rand din nomenclator per colaborator
                SELECT v.ID_Profesor, v.NumeIntreg, v.DenumireCatedra, v.DenumireFacultate,
                       ROW_NUMBER() OVER(PARTITION BY v.ID_Profesor ORDER BY v.Ordine DESC) AS Rn
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] v
                INNER JOIN ColabBase cb ON cb.ID_Profesor = v.ID_Profesor
                WHERE v.ID_AnUnivCatedra = 45
            ),
            FaraEntry AS (
                -- Colaboratori care nu apar deloc in nomenclator
                SELECT DISTINCT vcm.ID_Profesor, vcm.NumeIntregProfesor AS NumeIntreg,
                       ISNULL(vcm.DenumireFacultate,'Nespecificat') AS DenumireFacultate,
                       vcm.StatDeFunctiiID_Catedra,
                       ROW_NUMBER() OVER(PARTITION BY vcm.ID_Profesor ORDER BY vcm.ID_Profesor) AS Rn
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN ColabBase cb ON cb.ID_Profesor = vcm.ID_Profesor
                WHERE vcm.ID_AnUniv = 45
                  AND NOT EXISTS (SELECT 1 FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] vx WHERE vx.ID_Profesor = vcm.ID_Profesor)
            ),
            RezultatBrut AS (
                SELECT nc.ID_Profesor, nc.NumeIntreg AS NumeComplet,
                       nc.DenumireCatedra, nc.DenumireFacultate AS Facultate
                FROM NomenclatorColab nc WHERE nc.Rn = 1
                UNION ALL
                SELECT fe.ID_Profesor, fe.NumeIntreg AS NumeComplet,
                       CAST(fe.StatDeFunctiiID_Catedra AS VARCHAR(50)) AS DenumireCatedra,
                       fe.DenumireFacultate AS Facultate
                FROM FaraEntry fe WHERE fe.Rn = 1
            ),
            Final AS (
                SELECT ID_Profesor, NumeComplet, DenumireCatedra, Facultate,
                       ROW_NUMBER() OVER(PARTITION BY ID_Profesor ORDER BY NumeComplet) AS Rn2
                FROM RezultatBrut
            )
            SELECT ID_Profesor, NumeComplet, DenumireCatedra, Facultate
            FROM Final
            WHERE Rn2 = 1
              AND LTRIM(RTRIM(ISNULL(NumeComplet,''))) != ''
              AND NumeComplet NOT LIKE '--%'
              AND (@fac='Toti' OR Facultate COLLATE Latin1_General_CI_AI=@fac COLLATE Latin1_General_CI_AI)
            ORDER BY NumeComplet";

        [HttpGet("colaboratori")]
        public ActionResult GetColaboratori(string? anUniv, string? facultate, string? departament)
        {
            var result = new List<object>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(ColaboratoriSql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string dept = reader["DenumireCatedra"]?.ToString() ?? "";
                if (long.TryParse(dept, out long idCat)) dept = GetDenumireCatedra(idCat);
                result.Add(new
                {
                    Profesor = FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    Departament = dept,
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
            conn.Open();
            using var cmd = new SqlCommand(ColaboratoriSql, conn);
            cmd.CommandTimeout = 120;
            cmd.Parameters.AddWithValue("@fac", facultate ?? "Toti");
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string dept = reader["DenumireCatedra"]?.ToString() ?? "";
                if (long.TryParse(dept, out long idCat)) dept = GetDenumireCatedra(idCat);
                dt.Rows.Add(
                    FixNume(reader["NumeComplet"]?.ToString(), reader["ID_Profesor"]),
                    dept, reader["Facultate"]?.ToString() ?? "");
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

        // =====================================================================
        // FIX ANS:
        // - Lista titulari = acelasi criteriu ca R6: TitularAnUniv=1, ID_AnUnivCatedra=45, DidCerc='D'
        // - Elimina cercetatori, randuri reziduale, nume goale
        // - Rezultat: lista identica cu Cadre_Didactice_Titulare
        // =====================================================================

        // Interogare titulari didactici pentru ANS (identica cu TitulariSql ca logica)
        private const string QueryTitulariANS = @"
            WITH T AS (
                SELECT v.ID_Profesor, v.NumeIntreg, v.DenumireCatedra,
                       v.DenumireFacultate, v.DenumireGradDidactic,
                       ROW_NUMBER() OVER(PARTITION BY v.ID_Profesor ORDER BY v.Ordine DESC) AS Rn
                FROM [AGSIS].[dbo].[View_Profesori_CF_AnUniv] v
                WHERE v.TitularAnUniv = 1
                  AND v.ID_AnUnivCatedra = @ID_AnUniv
                  AND v.DidCerc = 'D'
                  AND LTRIM(RTRIM(ISNULL(v.NumeIntreg,''))) != ''
                  AND v.NumeIntreg NOT LIKE '--%'
            )
            SELECT ID_Profesor, NumeIntreg, DenumireCatedra, DenumireFacultate, DenumireGradDidactic
            FROM T
            WHERE Rn = 1
            ORDER BY NumeIntreg";

        [HttpGet("date-ans")]
        public IActionResult GetDateANS([FromQuery] int idAnUniv = 45)
        {
            var totiTitularii = new List<(int Id, string Nume, string Facultate, string Dept, string Grad)>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var cmd1 = new SqlCommand(QueryTitulariANS, conn))
            {
                cmd1.CommandTimeout = 60;
                cmd1.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                using var r1 = cmd1.ExecuteReader();
                while (r1.Read())
                    totiTitularii.Add((
                        Convert.ToInt32(r1["ID_Profesor"]),
                        FixNume(r1["NumeIntreg"]?.ToString(), r1["ID_Profesor"]),
                        r1["DenumireFacultate"]?.ToString() ?? "",
                        r1["DenumireCatedra"]?.ToString() ?? "",
                        r1["DenumireGradDidactic"]?.ToString() ?? ""
                    ));
            }

            string queryOre = @"
                SELECT vcm.ID_Profesor,
                       CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4)) AS OreConventionale,
                       sf.id_metaspecializare AS IdMetaspec
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN (
                    SELECT ID_StatDeFunctii, ID_AnUniv, DenumireSpecializare, DenumireMaterie, NrSemestruDinAn,
                           MAX(id_metaspecializare) AS id_metaspecializare,
                           MAX(DenTitularSauSuplinitor) AS DenTitularSauSuplinitor
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare]
                    GROUP BY ID_StatDeFunctii, ID_AnUniv, DenumireSpecializare, DenumireMaterie, NrSemestruDinAn
                ) sf
                    ON sf.ID_StatDeFunctii=vcm.ID_StatDeFunctii AND sf.ID_AnUniv=vcm.ID_AnUniv
                    AND sf.DenumireSpecializare=vcm.DenumireSpecializare
                    AND sf.DenumireMaterie=vcm.DenumireMaterie AND sf.NrSemestruDinAn=vcm.NrSemestruDinAn
                WHERE vcm.ID_AnUniv=@ID_AnUniv";
            var orePerProf = new Dictionary<int, List<(int IdMeta, decimal Ore)>>();
            using (var cmd2 = new SqlCommand(queryOre, conn))
            {
                cmd2.CommandTimeout = 120;
                cmd2.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                using var r2 = cmd2.ExecuteReader();
                while (r2.Read())
                {
                    int idProf = Convert.ToInt32(r2["ID_Profesor"]);
                    int idMeta = r2["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(r2["IdMetaspec"]) : 0;
                    decimal ore = Convert.ToDecimal(r2["OreConventionale"]);
                    if (!orePerProf.ContainsKey(idProf)) orePerProf[idProf] = new List<(int, decimal)>();
                    orePerProf[idProf].Add((idMeta, ore));
                }
            }

            var profesori = new List<object>();
            int nrCrt = 1;
            foreach (var (id, nume, fac, dept, grad) in totiTitularii.OrderBy(t => t.Nume))
            {
                var fractiuni = new Dictionary<string, decimal>();
                if (orePerProf.TryGetValue(id, out var listaOre))
                {
                    var orePerAns = new Dictionary<int, decimal>();
                    foreach (var (idMeta, ore) in listaOre)
                    {
                        if (!MappingMetaspec.TryGetValue(idMeta, out int idAns)) continue;
                        if (!AnsIdToCol.ContainsKey(idAns)) continue;
                        if (!orePerAns.ContainsKey(idAns)) orePerAns[idAns] = 0m;
                        orePerAns[idAns] += ore;
                    }
                    decimal totalOre = orePerAns.Values.Sum();
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
                }
                profesori.Add(new
                {
                    NrCrt = nrCrt++,
                    NumeComplet = nume,
                    Facultate = fac,
                    Departament = dept,
                    GradFunctie = MapareGradANS(grad),
                    DomeniiMapate = fractiuni
                });
            }
            return Ok(profesori);
        }

        [HttpGet("export/raport-ans")]
        public IActionResult ExportRaportANS([FromQuery] int idAnUniv = 45)
        {
            var totiTitulariiExp = new List<(int Id, string Nume, string Facultate, string Dept, string Grad)>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var cmd1 = new SqlCommand(QueryTitulariANS, conn))
            {
                cmd1.CommandTimeout = 60;
                cmd1.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                using var r1 = cmd1.ExecuteReader();
                while (r1.Read())
                    totiTitulariiExp.Add((
                        Convert.ToInt32(r1["ID_Profesor"]),
                        FixNume(r1["NumeIntreg"]?.ToString(), r1["ID_Profesor"]),
                        r1["DenumireFacultate"]?.ToString() ?? "",
                        r1["DenumireCatedra"]?.ToString() ?? "",
                        r1["DenumireGradDidactic"]?.ToString() ?? ""
                    ));
            }

            string queryOreExp = @"
                SELECT vcm.ID_Profesor,
                       CAST(ISNULL(vcm.NrOreConventionale,0) AS DECIMAL(10,4)) AS OreConventionale,
                       sf.id_metaspecializare AS IdMetaspec
                FROM [AGSIS].[pi].[View_CentralizareMateriiProfesor] vcm
                INNER JOIN (
                    SELECT ID_StatDeFunctii, ID_AnUniv, DenumireSpecializare, DenumireMaterie, NrSemestruDinAn,
                           MAX(id_metaspecializare) AS id_metaspecializare,
                           MAX(DenTitularSauSuplinitor) AS DenTitularSauSuplinitor
                    FROM [AGSIS].[pi].[StatDeFunctiiPeSpecializare]
                    GROUP BY ID_StatDeFunctii, ID_AnUniv, DenumireSpecializare, DenumireMaterie, NrSemestruDinAn
                ) sf
                    ON sf.ID_StatDeFunctii=vcm.ID_StatDeFunctii AND sf.ID_AnUniv=vcm.ID_AnUniv
                    AND sf.DenumireSpecializare=vcm.DenumireSpecializare
                    AND sf.DenumireMaterie=vcm.DenumireMaterie AND sf.NrSemestruDinAn=vcm.NrSemestruDinAn
                WHERE vcm.ID_AnUniv=@ID_AnUniv";
            var orePerProfExp = new Dictionary<int, List<(int IdMeta, decimal Ore)>>();
            using (var cmd2 = new SqlCommand(queryOreExp, conn))
            {
                cmd2.CommandTimeout = 120;
                cmd2.Parameters.AddWithValue("@ID_AnUniv", idAnUniv);
                using var r2 = cmd2.ExecuteReader();
                while (r2.Read())
                {
                    int idProf = Convert.ToInt32(r2["ID_Profesor"]);
                    int idMeta = r2["IdMetaspec"] != DBNull.Value ? Convert.ToInt32(r2["IdMetaspec"]) : 0;
                    decimal ore = Convert.ToDecimal(r2["OreConventionale"]);
                    if (!orePerProfExp.ContainsKey(idProf)) orePerProfExp[idProf] = new List<(int, decimal)>();
                    orePerProfExp[idProf].Add((idMeta, ore));
                }
            }

            var profesori = new List<ProfANS>();
            foreach (var (id, nume, fac, dept, grad) in totiTitulariiExp.OrderBy(t => t.Nume))
            {
                var fractiuni = new Dictionary<int, decimal>();
                if (orePerProfExp.TryGetValue(id, out var listaOreE))
                {
                    var orePerCol = new Dictionary<int, decimal>();
                    foreach (var (idMeta, ore) in listaOreE)
                    {
                        if (!MappingMetaspec.TryGetValue(idMeta, out int idAns)) continue;
                        if (!AnsIdToCol.ContainsKey(idAns)) continue;
                        int col = AnsIdToCol[idAns];
                        if (!orePerCol.ContainsKey(col)) orePerCol[col] = 0m;
                        orePerCol[col] += ore;
                    }
                    decimal totalOre = orePerCol.Values.Sum();
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
                }
                profesori.Add(new ProfANS
                {
                    NumeComplet = nume,
                    Departament = dept,
                    Facultate = fac,
                    GradFunctie = MapareGradANS(grad),
                    Fractiuni = fractiuni
                });
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
                var prof = profesori[i]; int r = dataStartRow + i;
                ws.Cell(r, 1).Value = i + 1; ws.Cell(r, 2).Value = prof.NumeComplet;
                ws.Cell(r, 3).Value = ""; ws.Cell(r, 4).Value = prof.GradFunctie;
                ws.Cell(r, 5).Value = 1; ws.Cell(r, 6).Value = 0; ws.Cell(r, 7).Value = "";
                ws.Cell(r, 8).Value = prof.Facultate; ws.Cell(r, 9).Value = prof.Departament;
                foreach (var kv in prof.Fractiuni) ws.Cell(r, kv.Key).Value = kv.Value;
                ws.Cell(r, 50).FormulaA1 = $"=SUM(J{r}:AW{r})";
                if (i % 2 != 0) for (int c = 1; c <= 50; c++) ws.Cell(r, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5f5f5");
            }
            int totalRow = dataStartRow + profesori.Count;
            ws.Cell(totalRow, 1).Value = "Total general:"; ws.Cell(totalRow, 1).Style.Font.Bold = true;
            for (int c = 10; c <= 49; c++) { string cl = ColumnLetter(c); ws.Cell(totalRow, c).FormulaA1 = $"=SUM({cl}{dataStartRow}:{cl}{totalRow - 1})"; ws.Cell(totalRow, c).Style.Font.Bold = true; }
            ws.Cell(totalRow, 50).FormulaA1 = $"=SUM(J{totalRow}:AW{totalRow})"; ws.Cell(totalRow, 50).Style.Font.Bold = true;
            using var stream = new MemoryStream(); wb.SaveAs(stream); wb.Dispose();
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Date_ANS_{idAnUniv}.xlsx");
        }

        #endregion

        #region ================= HELPERS =================

        private void AddBaseParams(SqlCommand cmd, string? anUniv, string? facultate, string? departament,
            string? formaInvatamant, string? profesor, string? specializari, int semestru, string? tipPost)
        {
            cmd.Parameters.AddWithValue("@an", anUniv ?? "Toti");
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

        private string MapareGradANS(string? gradDidactic)
        {
            if (string.IsNullOrWhiteSpace(gradDidactic)) return "Asist. dr.";
            string g = gradDidactic.ToUpperInvariant().Trim();
            if (g.Contains("PROFESOR UNIVERSITAR") || (g.Contains("PROFESOR") && !g.Contains("CONSULTANT"))) return "Prof. dr.";
            if (g.Contains("CONFERENTIAR") || g.Contains("CONFERENȚIAR")) return "Conf. dr.";
            if (g.Contains("LECTOR") || g.Contains("SEF LUCR") || g.Contains("ȘEF LUCR")
                || g.Contains("SEFUL LUCR") || (g.Contains("SEF") && g.Contains("LUCRARI"))
                || (g.Contains("ȘEF") && g.Contains("LUCRĂRI"))) return "Șef lucr. dr.";
            if (g.Contains("ASISTENT")) return "Asist. dr.";
            if (g.Contains("PREPARATOR")) return "Preparator";
            if (g.Contains("CERCET") && g.Contains(" I ")) return "CS I";
            if (g.Contains("CERCET") && g.Contains(" II")) return "CS II";
            if (g.Contains("CERCET") && g.Contains(" III")) return "CS III";
            if (g.Contains("CERCET")) return "CS";
            return "Asist. dr.";
        }

        private XLWorkbook BuildANSWorkbookFromScratch()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("CD DRU");
            ws.Cell(2, 1).Value = "Anexa 1. Tabel institutional privind normarea si activitatea cadrelor didactice si de cercetare";
            ws.Range(2, 1, 2, 50).Merge();
            ws.Cell(3, 1).Value = "Universitatea Transilvania din Brasov"; ws.Range(3, 1, 3, 6).Merge();
            ws.Cell(4, 1).Value = "NOTA: Se includ in tabel toate cadrele didactice si de cercetare titulare (cu norma de baza in universitate), indiferent de forma de angajare.";
            ws.Range(4, 1, 4, 9).Merge();
            ws.Cell(4, 10).Value = "NOTA: IMPORTANT! Va rugam sa completati in prima faza, in sectiunile aferente, fractiunile de norma pentru fiecare domeniu de stiinta.";
            ws.Range(4, 10, 4, 27).Merge(); ws.Range(4, 28, 4, 50).Merge();
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
            string[] subdomenii = {
                "Matematica","Informatica","Fizica","Chimie si inginerie chimica","Stiintele pamantului si atmosferei",
                "Inginerie civila","Inginerie electrica, electronica si telecomunicatii",
                "Inginerie geologica, mine, petrol si gaze","Ingineria transporturilor",
                "Ingineria resurselor vegetale si animale",
                "Ingineria sistemelor, calculatoare si tehnologia informatiei",
                "Inginerie mecanica, mecatronica, inginerie industriala si management",
                "Biologie","Biochimie","Medicina","Medicina veterinara","Medicina dentara","Farmacie",
                "Stiinte juridice","Stiinte administrative","Stiinte ale comunicarii","Sociologie",
                "Stiinte politice","Stiinte militare, informatii si ordine publica",
                "Stiinte economice (doar Cibernetica, statistica si informatica economica)",
                "Stiinte economice (fara Cibernetica, statistica si informatica economica)",
                "Psihologie si stiinte comportamentale","Filologie","Filosofie","Istorie","Teologie",
                "Studii culturale","Arhitectura si urbanism","Arte vizuale (fara Istoria si teoria artei)",
                "Arte vizuale (doar Istoria si teoria artei)","Teatru si artele spectacolului",
                "Cinematografie si media","Muzica (doar Interpretare muzicala)",
                "Muzica (fara Interpretare muzicala)","Stiintele Sportului si Educatiei Fizice"
            };
            for (int i = 0; i < subdomenii.Length; i++) { ws.Cell(6, 10 + i).Value = subdomenii[i]; ws.Range(6, 10 + i, 7, 10 + i).Merge(); }
            for (int i = 0; i < 9; i++) ws.Cell(8, i + 1).Value = ((char)('A' + i)).ToString();
            for (int i = 0; i < 41; i++) ws.Cell(8, 10 + i).Value = i + 1;
            ws.Cell(8, 50).Value = "40";
            var headerFill = XLColor.FromHtml(BrandColorHex);
            for (int r = 5; r <= 8; r++) for (int c = 1; c <= 50; c++)
            {
                ws.Cell(r, c).Style.Font.Bold = true; ws.Cell(r, c).Style.Font.FontColor = XLColor.Black;
                ws.Cell(r, c).Style.Fill.BackgroundColor = XLColor.White;
                ws.Cell(r, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(r, c).Style.Alignment.WrapText = true;
                ws.Cell(r, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }
            ws.Range(5, 1, 5, 50).Style.Fill.BackgroundColor = headerFill;
            ws.Range(5, 1, 5, 50).Style.Font.FontColor = XLColor.White;
            ws.Column(1).Width = 5; ws.Column(2).Width = 30; ws.Column(3).Width = 14;
            ws.Column(4).Width = 22; ws.Column(5).Width = 10; ws.Column(6).Width = 12;
            ws.Column(7).Width = 8; ws.Column(8).Width = 28; ws.Column(9).Width = 28;
            for (int c = 10; c <= 50; c++) ws.Column(c).Width = 12;
            return wb;
        }

        private string MapareFunctieANSbyId(int idTipGrad) => idTipGrad switch
        {
            1 => "Prof. dr.",
            2 => "Conf. dr.",
            3 => "Șef lucr. dr.",
            4 => "Asist. dr.",
            7 => "Preparator",
            9 => "Asist. dr.",
            10 => "CS III",
            11 => "Șef lucr. dr.",
            18 => "CS I",
            _ => "Asist. dr."
        };

        public class ProfANS
        {
            public string NumeComplet { get; set; } = "";
            public string Departament { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string GradFunctie { get; set; } = "";
            public Dictionary<int, decimal> Fractiuni { get; set; } = new();
        }

        public class RandSqlANS
        {
            public string NumeComplet { get; set; } = "";
            public string Facultate { get; set; } = "";
            public string Departament { get; set; } = "";
            public string GradDidactic { get; set; } = "";
            public decimal OreConventionale { get; set; }
            public int IdANS { get; set; }
        }

        #endregion
    }
}