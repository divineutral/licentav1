# Rapoarte Cadre Didactice
### Universitatea Transilvania din Brașov
**Lucrare de licență Ioniță Diana-Andreea, 2026**

Aplicație web de tip Business Intelligence pentru vizualizarea și exportul datelor privind normele didactice, dezvoltată cu ASP.NET Core 10 și Microsoft SQL Server.

---

## Cerințe de sistem

- [.NET 10 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/10.0)
- Microsoft SQL Server (acces la baza de date AGSIS a Universității Transilvania)
- Visual Studio 2022 (recomandat) sau orice editor cu suport C#
- Browser modern (Chrome, Edge, Firefox)

---

## Instalare și rulare

### 1. Clonează repository-ul

```bash
git clone https://github.com/divineutral/licentav1.git
cd licentav1
```
> Sau alternativ, descarca arhiva ZIP si extrage-o. 

### 2. Configurează conexiunea la baza de date

Deschide fișierul `licentav1/appsettings.json` și înlocuiește valorile placeholder cu credențialele tale:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=YOUR_SERVER;Database=YOUR_DB;User Id=YOUR_USER;Password=YOUR_PASSWORD;TrustServerCertificate=True;Max Pool Size=500;Connect Timeout=60;Command Timeout=240;"
  }
}
```

> Aplicația necesită acces **read-only** la baza de date AGSIS și la vizualizările:
> `View_PostProfesorMaterie`, `View_Profesori_CF_AnUniv`, `View_FDS`, `AnUniversitar`

### 3. Restaurează pachetele NuGet

```bash
cd licentav1
dotnet restore
```

Pachetele instalate automat:

| Pachet | Versiune | Rol |
|--------|----------|-----|
| ClosedXML | 0.105.0 | Generare fișiere Excel |
| Microsoft.Data.SqlClient | 6.1.4 | Conexiune SQL Server |
| Microsoft.AspNetCore.OpenApi | 10.0.2 | Documentație API |
| Swashbuckle.AspNetCore | 10.1.0 | Interfață Swagger |
| Microsoft.EntityFrameworkCore.SqlServer | 10.0.2 | Acces date |

### 4. Rulează aplicația

```bash
dotnet run licentav1
```

sau din Visual Studio 2022: deschide `licentav1.slnx` → F5.

### 5. Accesează interfața

| URL | Conținut |
|-----|----------|
| `http://localhost:5069/index.html` | Dashboard principal |
| `http://localhost:5069/swagger` | Documentație API (Swagger UI) |

---

## Structura proiectului

```
licentav1/
├── Controllers/
│   └── RapoarteController.cs     # toate endpointurile API + logica de business
├── Services/
│   └── Program.cs                # configurare ASP.NET Core
├── Properties/
│   └── launchSettings.json       # porturi și profil de rulare
├── wwwroot/
│   ├── index.html                # interfața utilizator (SPA)
│   └── images/
│       └── logo.png
├── appsettings.json              # configurare conexiune BD (de completat)
└── licentav1.csproj              # dependențe NuGet
```

---

## Rapoarte disponibile

| # | Endpoint | Descriere |
|---|----------|-----------|
| 1 | `GET /api/Rapoarte/norma` | Detaliere norme profesori |
| 2 | `GET /api/Rapoarte/norma-totaluri` | Totaluri ore IF / ID / IFR |
| 3 | `GET /api/Rapoarte/distributie-ore` | Distribuție ore pe programe de studiu |
| 4 | `GET /api/Rapoarte/limbi-straine` | Predare în limbi străine |
| 5 | `GET /api/Rapoarte/discipline` | Discipline predate per profesor |
| 6 | `GET /api/Rapoarte/titulari` | Cadre didactice titulare |
| 7 | `GET /api/Rapoarte/colaboratori` | Asociați și colaboratori |
| 8 | `GET /api/Rapoarte/raport-ans` | Raport ANS —> fracții de normă pe ramuri științifice |

Fiecare raport are un endpoint de export Excel corespunzător: `GET /api/Rapoarte/export/{nume-raport}`

Raportul „Discipline" include și export per profesor cu arhivă ZIP: `GET /api/Rapoarte/export/discipline-zip`
---

## Tehnologii utilizate

- **Backend:** ASP.NET Core 10 Web API, C#, ADO.NET
- **Bază de date:** Microsoft SQL Server (interogări parametrizate, CTE, window functions)
- **Export:** ClosedXML pentru generare .xlsx
- **Frontend:** HTML5, CSS3, JavaScript, Bootstrap 5, jQuery, Select2
- **Versionare:** Git / GitHub
