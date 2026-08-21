# PDF → Excel Automatizācija ar AI

Rīks, kas automātiski nolasa PDF dokumentus (rēķinus, līgumus, atskaites u.c.), izmantojot Claude AI izvelk no tiem svarīgāko informāciju un sakārto to Excel failā — katram dokumenta tipam atsevišķā lapā.

## Ko tas dara

- Nolasa visus PDF failus norādītajā mapē, neatkarīgi no to struktūras vai formāta.
- Izmanto AI (Claude), lai noteiktu dokumenta tipu (Rēķins, Līgums, Atskaite u.c.) un izvilktu galvenos datus.
- Izvelk arī sarakstu ar 3–7 svarīgākajiem punktiem/nosacījumiem no katra dokumenta (piem., termiņus, saistības, līgumsodus).
- Automātiski ieraksta visu Excel failā — katram dokumenta tipam sava lapa, dati uzkrājas laika gaitā.

## Kāpēc AI, nevis fiksēti noteikumi

PDF dokumentu struktūra bieži ir dažāda un iepriekš nezināma, tāpēc klasiska pieeja (meklēt datus pēc fiksētas rindas/kolonnas) nedarbotos. AI modelis saprot dokumenta *saturu*, nevis tikai izkārtojumu — tāpēc rīks strādā ar praktiski jebkuru dokumenta formātu bez papildu pielāgošanas.

## Izmantotās tehnoloģijas

| Komponente | Tehnoloģija |
|---|---|
| Valoda | Python |
| PDF teksta izvilkšana | `pdfplumber` |
| AI analīze | Anthropic Claude API |
| Excel apstrāde | `openpyxl` |

## Uzstādīšana

```bash
pip install pdfplumber openpyxl anthropic
```

Iestati savu Anthropic API atslēgu kā vides mainīgo:

```bash
export ANTHROPIC_API_KEY="tava-atslega"
```

## Lietošana

1. Ievieto PDF failus mapē `pdf_faili`.
2. Palaid skriptu:

```bash
python pdf_uz_excel.py
```

3. Rezultāts tiks saglabāts failā `izvilktie_dati.xlsx` ar kolonnām: *Fails, Datums, Summa, Puses, Kopsavilkums, Svarīgie punkti, Apstrādāts*.

## Ierobežojumi

- Paredzēts teksta PDF apstrādei — skenētiem attēlu PDF nepieciešama papildu OCR apstrāde (nav iekļauta šajā versijā).
- AI izvilktā informācija jāpārbauda cilvēkam kritiskos gadījumos (piem., finanšu datos).

## Iespējamie uzlabojumi

- OCR atbalsts skenētiem dokumentiem.
- Grafiskā saskarne bez komandrindas lietošanas.
- Automātisks e-pasta paziņojums pēc apstrādes pabeigšanas.

---

**Autors:** Katrīna Saučuka
