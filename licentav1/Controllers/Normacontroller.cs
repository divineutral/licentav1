using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Data;

namespace LicentaV1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NormaController : ControllerBase
    {
        private readonly string _connectionString;
        private const string BrandColor = "#56723e";

        public NormaController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        // Coloane sincronizate cu Excel:
        // Nr.Crt, Profesor, Departament, Specializare, Materie, TipPost, Sem,
        // OreCurs, OreAplic(Sem+Lab+Prj), OreConv
        // JOIN pe ID_PlanMaterie_Prestator pentru cuplaje
        private const string SqlNorma = @"
            SELECT
                ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20)))   AS IdentificatorUnic,
                MIN(p.NumeIntreg) OVER
                    (PARTITION BY ISNULL(p.CNP, CAST(p.ID_Profesor AS VARCHAR(20))))
                                                                     AS NumeIntreg,
                p.DenumireCatedra                                    AS Departament,
                ppm.DenumireSpecializare,
                ppm.Denumire                                         AS Materie,
                ppm.TitularSauSuplinitor,
                ppm.NrSemestruDinAn                                  AS Semestru,
                ppm.Nr_Ore_Curs,
                ppm.Nr_Ore_Seminar + ppm.Nr_Ore_Laborator + ppm.Nr_Ore_Proiect
                                                                     AS Nr_Ore_Aplicatii,
                ppm.NrOreConventionale,
                ppm.DenumireFormaInv,
                dc.ID_Cuplaj
            FROM [AGSIS_DW].[dbo].[Post_Profesor_Materie] ppm
            INNER JOIN [AGSIS].[dbo].[View_Profesori_CF_AnUniv] p
                ON ppm.ID_Profesor = p.ID_Profesor
                AND p.ID_AnUnivCatedra = @ID_AnUniv
            LEFT JOIN [AGSIS].[pi].[View_DetaliereCuplaje] dc
                ON ppm.ID_PlanMaterie_Prestator = dc.ID_PlanMaterie_Prestator
                AND dc.ID_AnUniv = @ID_AnUniv
            WHERE ppm.ID_AnUniv = @ID_AnUniv
                AND (@fac = N'Toti'
                     OR p.ID_Facultate = TRY_CAST(@fac AS INT)
                     OR ppm.DenumireFacultate COLLATE DATABASE_DEFAULT
                      = @fac COLLATE DATABASE_DEFAULT)
                AND (@dept = N'Toti'
                     OR p.ID_Catedra = TRY_CAST(@dept AS INT)
                     OR ppm.DenumireCatedra COLLATE DATABASE_DEFAULT
                      = @dept COLLATE DATABASE_DEFAULT)
                AND (@specs = N'Toti' OR EXISTS (
                    SELECT 1 FROM STRING_SPLIT(@specs, ',') s 
                    WHERE ppm.DenumireSpecializare LIKE s.value + '%' 
                       OR ppm.DenumireSpecializare LIKE '%' + s.value + '%'
                ))
                AND (@prof = N'Toti'
                     OR p.NumeIntreg COLLATE DATABASE_DEFAULT
                      = @prof COLLATE DATABASE_DEFAULT)
                AND (@formaInv = N'Toti'
                     OR ppm.DenumireFormaInv COLLATE DATABASE_DEFAULT
                      = @formaInv COLLATE DATABASE_DEFAULT)
                AND (@sem = 0 OR ppm.NrSemestruDinAn = @sem)
                AND (@tipPost = N'Toti'
                     OR (@tipPost = N'Titular'    AND ppm.TitularSauSuplinitor = 1)
                     OR (@tipPost = N'Suplinitor' AND ppm.TitularSauSuplinitor = 0))

            ORDER BY NumeIntreg, ppm.DenumireSpecializare, ppm.Denumire, ppm.NrSemestruDinAn";

        private void AddParams(SqlCommand cmd, int id, string fac, string dept,
            string prof, string forma, int sem, string tip)
        {
            cmd.Parameters.AddWithValue("@ID_AnUniv", id);
            cmd.Parameters.AddWithValue("@fac", fac);
            cmd.Parameters.AddWithValue("@dept", dept);
            cmd.Parameters.AddWithValue("@prof", prof);
            cmd.Parameters.AddWithValue("@formaInv", forma);
            cmd.Parameters.AddWithValue("@sem", sem);
            cmd.Parameters.AddWithValue("@tipPost", tip);
        }

        private (string fac, string dept, string prof, string forma, string tip)
            ParseFilters(string? f, string? d, string? p, string? fi, string? t) =>
            (
                string.IsNullOrWhiteSpace(f) ? "Toti" : f.Trim(),
                string.IsNullOrWhiteSpace(d) ? "Toti" : d.Trim(),
                string.IsNullOrWhiteSpace(p) ? "Toti" : p.Trim(),
                string.IsNullOrWhiteSpace(fi) ? "Toti" : fi.Trim(),
                string.IsNullOrWhiteSpace(t) ? "Toti" : t.Trim()
            );

        [HttpGet]
        public async Task<IActionResult> GetNorma(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null, [FromQuery] string? departament = null,
            [FromQuery] string? profesor = null, [FromQuery] string? formaInvatamant = null,
            [FromQuery] int semestru = 0, [FromQuery] string? tipPost = null)
        {
            var (fac, dept, prof, forma, tip) = ParseFilters(facultate, departament, profesor, formaInvatamant, tipPost);
            var result = new List<object>();
            var cuplajeVaz = new HashSet<long>();

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma, conn);
            cmd.CommandTimeout = 120;
            AddParams(cmd, idAnUniv, fac, dept, prof, forma, semestru, tip);

            int nrCrt = 1;
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                long? cup = r["ID_Cuplaj"] == DBNull.Value ? null : (long?)Convert.ToInt64(r["ID_Cuplaj"]);
                // Deduplicare curs la comun: daca acelasi cuplaj, nu adaugam de doua ori
                if (cup.HasValue && !cuplajeVaz.Add(cup.Value)) continue;

                result.Add(new
                {
                    NrCrt = nrCrt++,
                    Profesor = r["NumeIntreg"].ToString(),
                    Departament = r["Departament"].ToString(),
                    Specializare = r["DenumireSpecializare"].ToString(),
                    Materie = r["Materie"].ToString(),
                    TipPost = Convert.ToInt32(r["TitularSauSuplinitor"]) == 1 ? "Titular" : "Suplinitor",
                    Semestru = r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    NrOreCurs = r["Nr_Ore_Curs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Nr_Ore_Curs"]),
                    NrOreAplicatii = r["Nr_Ore_Aplicatii"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Nr_Ore_Aplicatii"]),
                    NrOreConventionale = r["NrOreConventionale"] == DBNull.Value ? 0m : Convert.ToDecimal(r["NrOreConventionale"]),
                    FormaInv = r["DenumireFormaInv"].ToString()
                });
            }
            return Ok(result);
        }

        [HttpGet("export")]
        public async Task<IActionResult> ExportNorma(
            [FromQuery] int idAnUniv = 45,
            [FromQuery] string? facultate = null, [FromQuery] string? departament = null,
            [FromQuery] string? profesor = null, [FromQuery] string? formaInvatamant = null,
            [FromQuery] int semestru = 0, [FromQuery] string? tipPost = null)
        {
            var (fac, dept, prof, forma, tip) = ParseFilters(facultate, departament, profesor, formaInvatamant, tipPost);
            var dt = new DataTable();
            dt.Columns.AddRange(new[]
            {
                new DataColumn("Nr. Crt.", typeof(int)),
                new DataColumn("Profesor"),
                new DataColumn("Departament"),
                new DataColumn("Specializare"),
                new DataColumn("Materie"),
                new DataColumn("Tip Post"),
                new DataColumn("Sem.", typeof(int)),
                new DataColumn("Ore Curs",   typeof(decimal)),
                new DataColumn("Ore Aplic.", typeof(decimal)),
                new DataColumn("Ore Conv.",  typeof(decimal)),
                new DataColumn("Forma Inv.")
            });

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var cmd = new SqlCommand(SqlNorma, conn);
            cmd.CommandTimeout = 120;
            AddParams(cmd, idAnUniv, fac, dept, prof, forma, semestru, tip);

            int nrCrt = 1;
            var cuplajeVaz = new HashSet<long>();
            using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                long? cup = r["ID_Cuplaj"] == DBNull.Value ? null : (long?)Convert.ToInt64(r["ID_Cuplaj"]);
                if (cup.HasValue && !cuplajeVaz.Add(cup.Value)) continue;
                dt.Rows.Add(
                    nrCrt++,
                    r["NumeIntreg"].ToString(),
                    r["Departament"].ToString(),
                    r["DenumireSpecializare"].ToString(),
                    r["Materie"].ToString(),
                    Convert.ToInt32(r["TitularSauSuplinitor"]) == 1 ? "Titular" : "Suplinitor",
                    r["Semestru"] == DBNull.Value ? 0 : Convert.ToInt32(r["Semestru"]),
                    r["Nr_Ore_Curs"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Nr_Ore_Curs"]),
                    r["Nr_Ore_Aplicatii"] == DBNull.Value ? 0m : Convert.ToDecimal(r["Nr_Ore_Aplicatii"]),
                    r["NrOreConventionale"] == DBNull.Value ? 0m : Convert.ToDecimal(r["NrOreConventionale"]),
                    r["DenumireFormaInv"].ToString());
            }

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Norme");
            ws.Cell(1, 1).Value = $"Detaliere norme | An: {idAnUniv} | Facultate: {fac}";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml(BrandColor);
            ws.Range(1, 1, 1, dt.Columns.Count).Merge();
            ws.Cell(2, 1).Value = "NOTA: Ore Conv. = norma saptamanala reala (include coeficienti din plan)";
            ws.Cell(2, 1).Style.Font.Italic = true; ws.Cell(2, 1).Style.Font.FontColor = XLColor.Gray;
            ws.Range(2, 1, 2, dt.Columns.Count).Merge();

            var tbl = ws.Cell(4, 1).InsertTable(dt);
            tbl.Theme = XLTableTheme.None; tbl.ShowTotalsRow = true;
            tbl.Field("Ore Curs").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Aplic.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Ore Conv.").TotalsRowFunction = XLTotalsRowFunction.Sum;
            tbl.Field("Nr. Crt.").TotalsRowLabel = "TOTAL";
            ws.Range(4, 1, 4, dt.Columns.Count).Style.Fill.BackgroundColor = XLColor.FromHtml(BrandColor);
            ws.Range(4, 1, 4, dt.Columns.Count).Style.Font.FontColor = XLColor.White;
            ws.Range(4, 1, 4, dt.Columns.Count).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"NormaProfesori_{idAnUniv}.xlsx");
        }
    }
}