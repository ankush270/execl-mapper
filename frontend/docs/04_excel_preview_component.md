# Frontend Component Documentation: `excel-preview.component.ts`

Yeh document [`excel-preview.component.ts`](file:///e:/Projects/excelmapping/frontend/src/app/components/excel-preview/excel-preview.component.ts) file ka complete detailed technical breakdown hai.

---

## 📌 Component Overview
- **Location:** `src/app/components/excel-preview/`
- **Selector:** `<app-excel-preview>`
- **Imports:** SheetJS Library (`import * as XLSX from 'xlsx'`), PrimeNG `DialogModule`, `TabViewModule`, `ProgressSpinnerModule`.
- **Purpose:** Most Complex UI Component — Browser In-Memory Client-Side Excel Renderer. Downloads raw ArrayBuffer of `.xlsx` / `.xlsm` files, parses sheets, merges cells (`rowspan`/`colspan`), extracts inline cell background colors, borders, font weights, alignments, and builds styled HTML `<table>` elements inside tab views.

---

## 🔑 Data Models & State Properties

```typescript
interface SheetTab {
  name: string;
  html: string;
}

@Input() visible = false;
@Output() visibleChange = new EventEmitter<boolean>();
@Input() fileName = '';

sheets: SheetTab[] = [];
loading = false;
error = '';
```

1. `SheetTab`: Model holding worksheet title `name` and pre-rendered `html` table string.
2. `@Input() visible` & `@Input() fileName`: Controls Dialog modal visibility state and target template filename.
3. `sheets`: Parsed HTML tables list for PrimeNG TabView tabs.

---

## ⚙️ Component Methods Breakdown

### 1. `ngOnChanges(changes: SimpleChanges)`
```typescript
ngOnChanges(changes: SimpleChanges): void {
  if (changes['visible']?.currentValue === true && this.fileName) {
    this.loadPreview();
  }
  if (changes['fileName'] && this.visible && this.fileName) {
    this.loadPreview();
  }
}
```
- **Kya Karta Hai?** Component `@Input()` properties (`visible` / `fileName`) change hone par preview trigger karta hai.
- **Kyun Karta Hai?** Check karta hai dialog opened hai aur valid filename bind hai, to `loadPreview()` call karta hai.
- **Need Kyun Hai?** Reactive preview loading on user interaction.

---

### 2. `loadPreview()`
```typescript
private loadPreview(): void {
  this.loading = true;
  this.service.getTemplateFile(this.fileName).subscribe({
    next: (buffer) => {
      const wb = XLSX.read(buffer, {
        type: 'array', cellStyles: true, cellHTML: false, cellDates: true,
      });
      this.sheets = wb.SheetNames.map(name => ({
        name, html: this.sheetToStyledHtml(wb.Sheets[name], wb)
      }));
      this.loading = false;
    },
    error: (err) => { ... }
  });
}
```
- **Kya Karta Hai?** Service se template file ka raw binary ArrayBuffer download karke SheetJS `XLSX.read()` execute karta hai.
- **Kyun Karta Hai?** In-memory parsing options `cellStyles: true`, `cellDates: true` pass karke SheetJS WorkBook object banata hai, aur har worksheet ke liye `sheetToStyledHtml()` execute karke `sheets` array me store karta hai.
- **Need Kyun Hai?** Pure client-side zero-latency Excel Sheet preview generation.

---

### 3. `sheetToStyledHtml(ws: XLSX.WorkSheet, wb: XLSX.WorkBook): string`
```typescript
private sheetToStyledHtml(ws: XLSX.WorkSheet, wb: XLSX.WorkBook): string
```
- **Kya Karta Hai?** SheetJS WorkSheet object ko fully styled HTML `<table>` string me convert karta hai.
- **Kyun Karta Hai?**
  1. `XLSX.utils.decode_range(ws['!ref'])` se total populated row/col range decode karta hai.
  2. Merged Ranges (`ws['!merge']`) process karke `rowspan` aur `colspan` spans determine karta hai aur non-anchor merged cells ko skip karta hai (`mergedMap.get(cellKey)`).
  3. `<colgroup>` tags me column widths (`!cols`) set karta hai.
  4. Each cell ke liye `buildCellStyle(cell, ws)` call karke CSS styles nikalta hai aur `formatCellValue(cell)` se cell value format karta hai.
- **Need Kyun Hai?** OpenPyXL / Excel like visual grid layout rendering in HTML.

---

### 4. `buildCellStyle(cell: XLSX.CellObject, ws: XLSX.WorkSheet): string`
```typescript
private buildCellStyle(cell: XLSX.CellObject | undefined, ws: XLSX.WorkSheet): string
```
- **Kya Karta Hai?** SheetJS cell style object (`cell.s`) ko decode karke inline CSS string (`style="..."`) banata hai.
- **Kyun Karta Hai?**
  - **Background Fill:** `s.fill.fgColor` decode karke `background-color:#HEX` set karta hai.
  - **Font Styles:** `s.font` inspect karke `font-weight:bold`, `font-style:italic`, `font-size:PX`, `color:#HEX`, `font-family` apply karta hai.
  - **Alignment:** `s.alignment` inspect karke `text-align` (left/center/right), `vertical-align` (top/middle/bottom), `white-space:normal` apply karta hai.
  - **Borders:** `s.border` (top, bottom, left, right) inspect karke CSS borders (`border-left: 1px solid #000000`) apply karta hai.
- **Need Kyun Hai?** Excel template fonts, colors, and borders exact visual match karwane ke liye.

---

### 5. `resolveColor(color: any)` & `formatCellValue(cell)` & `escapeHtml(s)`
```typescript
private resolveColor(color: any): string | null
private formatCellValue(cell: XLSX.CellObject): string
private escapeHtml(s: string): string
```
- **Kya Karta Hai?**
  - `resolveColor`: Excel RGB hexadecimal strings (`FF0000`) ya Excel Theme Color index fallback palette array (`['FFFFFF', '000000', 'EEECE1', ...]`) se Hex Color Code nikalta hai.
  - `formatCellValue`: Cell value type (`cell.t`), formatted text (`cell.w`), ya raw value (`cell.v`) return karta hai.
  - `escapeHtml`: Special Characters (`&`, `<`, `>`) ko HTML entities (`&amp;`, `&lt;`, `&gt;`) me escape karta hai.
- **Kyun Karta Hai?** XSS Security vulnerabilities prevent karne ke liye string sanitization karta hai.
- **Need Kyun Hai?** Security & Accurate visual formatting.

---

### 6. `onClose()`
```typescript
onClose(): void {
  this.visible = false;
  this.visibleChange.emit(false);
  this.sheets = [];
  this.error = '';
}
```
- **Kya Karta Hai?** Dialog Close event on preview modal exit.
- **Kyun Karta Hai?** Resets preview state, clears HTML sheets memory, and emits `visibleChange(false)`.
- **Need Kyun Hai?** Dialog state cleanup & RAM memory deallocation.

---
*Created automatically by Antigravity AI Code Assistant.*
