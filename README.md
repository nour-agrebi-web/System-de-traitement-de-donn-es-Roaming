# Roaming Data Processing System

<!-- Badges: not applicable. -->

Internship project carried out at **Orange Tunisie**. Two Excel VBA macros process roaming data exports (VLR and SGSN numbers): they count how many times each number appears, identify the country from the international dialling code, list the mobile operators of that country, and produce a result workbook sorted by number of occurrences.

## Main features

- **`ProcessVLRData`** ([FoncionVLR_Solution .txt](./FoncionVLR_Solution%20.txt)): processes a VLR export.
- **`ProcessSGSNData`** ([FonctionResultSGSN.txt](./FonctionResultSGSN.txt)): processes an SGSN export.
- File pickers to select the input text file, the country reference workbook and the output file.
- Duplicate counting per VLR / SGSN number.
- Country detection by **longest matching dialling-code prefix**.
- Operator list per country from the reference workbook.
- Result sheet `Résultat` with the columns `VLR_NUMBER` / `SGSN_NUMBER`, `nombre_repetitions`, `Pays`, `Opérateurs disponibles`, sorted by number of occurrences (descending). Unknown prefixes are labelled `Inconnu`.
- Reference data: [indicateur_pays.xlsx](./indicateur_pays.xlsx) — 188 countries with their dialling code and main operators.

## Screenshots / demo

> To be completed: add a screenshot of an input file and of the generated result sheet (with anonymised numbers), for example in `./docs/screenshots/`.

## Prerequisites

- **Microsoft Excel for Windows** with macros enabled (the code uses `Scripting.Dictionary` through `CreateObject`, which is available on Windows only).
- The **Developer** tab (or `Alt` + `F11`) to open the VBA editor.

## Installation

```bash
git clone https://github.com/nour-agrebi-web/System-de-traitement-de-donn-es-Roaming.git
```

1. Open Excel and a new or existing workbook.
2. Press `Alt` + `F11` to open the VBA editor, then **Insert > Module**.
3. Paste the content of `FoncionVLR_Solution .txt` (and / or `FonctionResultSGSN.txt`) into the module.
4. Save the workbook as a macro-enabled workbook (`.xlsm`).

## Configuration / environment variables

None. All paths are chosen at run time through dialog boxes.

### Expected input formats

**Input text file** (VLR or SGSN export):

- comma-separated values;
- first line = header (ignored);
- the VLR / SGSN number is in the **3rd column**.

> To be completed: the exact structure of the export (columns 1 and 2) is not documented, and no sample input file is provided.

**Reference workbook** (first sheet, first line = header):

| Column A | Column B | Column C |
| --- | --- | --- |
| Country (`Pays`) | Dialling code (`Indicatif Téléphonique`) | Operators, comma-separated (`Opérateurs Téléphoniques`) |

`indicateur_pays.xlsx` follows this format.

## Running the project

1. In Excel, press `Alt` + `F8`.
2. Run `ProcessVLRData` or `ProcessSGSNData`.
3. Select, in order: the input `.txt` file, the reference workbook (`indicateur_pays.xlsx`), and the output file name (default `Résultat.xlsx`).
4. Information boxes show the number of countries loaded and the number of distinct numbers found.

## Usage

- `ProcessVLRData` displays the result workbook and asks whether to save it to the chosen output path.
- `ProcessSGSNData` builds and sorts the result workbook but does not save it (see notes): save it manually from Excel.

## Tests

> To be completed: no tests or sample data are provided.

## Project structure

```text
System-de-traitement-de-donn-es-Roaming/
├── FoncionVLR_Solution .txt   # VBA macro ProcessVLRData
├── FonctionResultSGSN.txt     # VBA macro ProcessSGSNData
└── indicateur_pays.xlsx       # country / dialling code / operators reference
```

## Technologies

- VBA (Visual Basic for Applications)
- Microsoft Excel (QueryTables text import, sorting, `Scripting.Dictionary`)

## Deployment

Not applicable: the macros are run manually from Excel.

## Contributing

> To be completed: no contribution guidelines exist.

## License

> To be completed: no `LICENSE` file is present. Check with Orange Tunisie whether the internship code may be published and under which terms.

## Authors / contacts

- Repository owner: [@nour-agrebi-web](https://github.com/nour-agrebi-web)

> To be completed: author's full name, internship supervisor, internship period.

## Useful resources

- [Getting started with VBA in Office](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)
- [Excel VBA reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel)

## Notes to complete

- `ProcessSGSNData` asks for an output path but never saves the result workbook; `ProcessVLRData` does.
- The file name `FoncionVLR_Solution .txt` contains a typo and a trailing space before the extension; consider renaming it (e.g. `FonctionVLR_Solution.bas`).
- The macros are stored as `.txt` files; exporting them as `.bas` modules would allow direct import into the VBA editor.
- Unused variables and unused error labels (`FileError`, `OutputError`) remain in the VBA code.
- The input export format, sample data, screenshots and author / internship details are missing.
- The reference workbook reflects operators at the time of the internship and may be outdated.
