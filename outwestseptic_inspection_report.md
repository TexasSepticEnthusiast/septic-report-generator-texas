# Out West Septic — Septic Inspection Report Workbook Spec

**Out West Septic — outwestseptic.com**
**TCEQ OSSF Installer I | OS0041298 | Lone Camp, TX**
**Owner: James "Tanner" Coupland | 940.452.0598**
**Serving Palo Pinto County & Rural North Texas**

This is the build spec for `outwestseptic_inspection_report.xlsm`. Tanner or the user builds the actual `.xlsm` in Excel by importing the CSV templates below. Open in Excel, then save as `.xlsm` (macro-enabled) if any VBA is added.

## Sheet 1 — Cover

| Cell | Content |
|------|---------|
| B2 | OUT WEST SEPTIC |
| B3 | Septic System Inspection Report |
| B4 | TCEQ OSSF Installer I — License OS0041298 |
| B5 | Lone Camp, TX |
| B6 | outwestseptic.com  •  940.452.0598 |
| B8 | Job site: |
| D8 | (input) |
| B9 | Date of inspection: |
| D9 | (input — `=TODAY()` by default) |
| B10 | Inspector: |
| D10 | Tanner Coupland, TCEQ OS0041298 |
| B11 | Reason for inspection: |
| D11 | (dropdown — see Validations below) |
| B12 | Property address: |
| D12 | (input) |
| B13 | County: |
| D13 | (dropdown — see Validations) |

**Conditional formatting (B14):** if overall score from Sheet 3 < 60 → "FAIL — see Notes"; if 60–79 → "CONDITIONAL PASS — repair recommended"; if 80+ → "PASS".

## Sheet 2 — Tank

| Cell | Label | Input type |
|------|-------|------------|
| B2 | Tank size (gallons) | integer (dropdown: 750, 1000, 1250, 1500, 2000, 2500) |
| B3 | Tank material | dropdown: Concrete, Fiberglass, Polyethylene, Steel |
| B4 | Tank age (years) | integer |
| B5 | Last pumped (date) | date |
| B6 | Cracks visible | Yes/No |
| B7 | Baffles intact | Yes/No |
| B8 | Inlet/outlet intact | Yes/No |
| B9 | Risers present | Yes |
| B10 | Tank watertight | Pass/Fail (with comment field) |
| B11 | Sludge depth (in) | number |
| B12 | Scum depth (in) | number |
| B13 | Liquid level OK | Pass/Fail |

**Conditional formatting:** B6/B7/B8 = "No" → red. B10/B13 = "Fail" → red.

## Sheet 3 — Drainfield

| Cell | Label | Input type |
|------|-------|------------|
| B2 | Drainfield type | dropdown: Gravel trench, Chamber, Pipe and gravel, Pressurized |
| B3 | Approximate age (years) | integer |
| B4 | Surfacing wet | Pass/Fail |
| B5 | Odors | Yes/No |
| B6 | Grass unusually green over field | Yes/No |
| B7 | Backup into house | Yes |
| B8 | Backup to surface | Yes |
| B9 | Lateral length (ft) | integer |
| B10 | Lateral count | integer |
| B11 | Trenches inspected (count) | integer |
| B12 | Distribution box OK | Pass/Fail |
| B13 | Inspection ports present | Yes/No |
| B14 | Perc test on file | Yes/No |

**Score (B15):** `=100 - (B4="Fail")*30 - (B5="Yes")*10 - (B6="Yes")*15 - (B7="Yes")*30 - (B8="Yes")*40 - (B12="Fail")*20`

## Sheet 4 — Site

| Cell | Label | Input type |
|------|-------|------------|
| B2 | Surface water visible within 50 ft | Yes/No |
| B3 | Septic smell outdoors | Yes/No |
| B4 | Pool of water over drainfield | Yes/No |
| B5 | Recent heavy use (large gathering, etc.) | Yes/No |
| B6 | Recent weather (heavy rain) | Yes/No |

## Sheet 5 — Summary (auto-calculated)

| Cell | Calculation |
|------|-------------|
| B2 | Tank score (formula references Sheet 2) |
| B3 | Drainfield score (formula references Sheet 3) |
| B4 | Site score (formula references Sheet 4) |
| B5 | Overall score = average of B2:B4 |
| B7 | Pass/Conditional/Fail band (per Cover B14 rule) |
| B9 | Inspector signature line |
| B11 | Date of report |

## Sheet 6 — Notes (free-form)

| Cell | Label |
|------|-------|
| B2 | Tank notes |
| B3 | Drainfield notes |
| B4 | Site notes |
| B5 | Recommended next steps |
| B6 | Customer initials |

## Validations

- **Reason for inspection:** Real Estate Sale, Routine Maintenance, Pre-Purchase, Lender Required, Suspected Problem, Other
- **County:** Palo Pinto County TX, Hood County TX, Erath County TX, Stephens County TX, Young County TX, Jack County TX, Eastland County TX, Mineral Wells TX, Graford TX, Possum Kingdom Lake TX

## Branding

- **Header on every sheet:** Out West Septic • TCEQ OSSF Installer I OS0041298 • outwestseptic.com • 940.452.0598
- **Footer on every sheet:** Serving Palo Pinto County & Rural North Texas • Lone Camp, TX
- **Color scheme:** Out West Septic teal `#007787`, accent orange `#cc5500`, neutral grays.

## Macro stubs (optional)

If Tanner wants to add VBA, suggested entry points:

```vba
Sub GeneratePDF()
    ' Export the active sheet as a PDF named for the property address.
    Dim path As String
    path = ThisWorkbook.Path & Application.PathSeparator & _
           Range("D12").Value & "_" & Format(Date, "yyyymmdd") & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path
End Sub

Sub ResetWorkbook()
    ' Clear all input cells on all sheets.
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Range("D8,D9,D11,D12,D13").ClearContents
    Next ws
End Sub
```

---

## CSV template (drop into Excel as a starting point)

```csv
section,field,type,validation,notes
cover,site,,,
cover,date,date,,
cover,inspector,text,Tanner Coupland,
cover,reason,choice,Real Estate Sale|Routine Maintenance|Pre-Purchase|Lender Required|Suspected Problem|Other,
cover,address,text,,
cover,county,choice,Palo Pinto County TX|Hood County TX|Erath County TX|Stephens County TX|Young County TX|Jack County TX|Eastland County TX|Mineral Wells TX|Graford TX|Possum Kingdom Lake TX,
tank,size_gal,integer,750|1000|1250|1500|2000|2500,
tank,material,choice,Concrete|Fiberglass|Polyethylene|Steel,
tank,age_years,integer,,
tank,last_pumped,date,,
tank,cracks,choice,Yes|No,
tank,baffles,choice,Yes|No,
tank,inlet_outlet,choice,Yes|No,
tank,risers,choice,Yes|No,
tank,watertight,choice,Pass|Fail,
tank,sludge_in,number,,
tank,scum_in,number,,
tank,liquid_level,choice,Pass|Fail,
drainfield,type,choice,Gravel trench|Chamber|Pipe and gravel|Pressurized,
drainfield,age_years,integer,,
drainfield,wet_surface,choice,Pass|Fail,
drainfield,odors,choice,Yes|No,
drainfield,green_grass,choice,Yes|No,
drainfield,backup_house,choice,Yes|No,
drainfield,backup_surface,choice,Yes|No,
drainfield,lateral_length_ft,integer,,
drainfield,lateral_count,integer,,
drainfield,trenches_inspected,integer,,
drainfield,dbox_ok,choice,Pass|Fail,
drainfield,ports,choice,Yes|No,
drainfield,perc_on_file,choice,Yes|No,
site,water_within_50ft,choice,Yes|No,
site,outdoor_smell,choice,Yes|No,
site,pool_over_field,choice,Yes|No,
site,recent_heavy_use,choice,Yes|No,
site,recent_heavy_rain,choice,Yes|No,
```

*Out West Septic — outwestseptic.com — 940.452.0598 — TCEQ OSSF Installer I OS0041298 — Lone Camp, TX*
*Serving Palo Pinto County, Mineral Wells, Graford, Possum Kingdom Lake, and rural North Texas*