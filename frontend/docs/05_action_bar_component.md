# Frontend Component Documentation: `action-bar.component.ts`

Yeh document [`action-bar.component.ts`](file:///e:/Projects/excelmapping/frontend/src/app/components/action-bar/action-bar.component.ts) file ka complete detailed technical breakdown hai.

---

## 📌 Component Overview
- **Location:** `src/app/components/action-bar/`
- **Selector:** `<app-action-bar>`
- **Imports:** PrimeNG `ButtonModule`, `ProgressSpinnerModule`, `TooltipModule`, `RippleModule`.
- **Purpose:** Bottom Sticky Action Bar Bar containing the "Generate & Download Excel" action trigger button, Payload compilation, Asynchronous execution tracking, and Browser Blob Auto-Download.

---

## 🔑 State Properties & Inputs

```typescript
@Input() selectedDoc: JsonDocument | null = null;
@Input() templateSel: TemplateSelection | null = null;

generating = false;

get canGenerate(): boolean {
  return !!(
    this.selectedDoc &&
    this.templateSel?.template &&
    this.templateSel?.sheetName?.trim()
  );
}
```

1. `@Input() selectedDoc`: Selected JSON document from `JsonPanelComponent`.
2. `@Input() templateSel`: Selected Excel Template and Sheet Name from `TemplatePanelComponent`.
3. `generating`: Progress spinner loading indicator state boolean.
4. `canGenerate`: Computed Getter property checking whether JSON Document, Template, and non-empty Sheet Name are all selected.

---

## ⚙️ Component Methods Breakdown

### 1. `canGenerate` (Getter)
```typescript
get canGenerate(): boolean
```
- **Kya Karta Hai?** Validation check karta hai ki JSON document selected hai ya nahi, Excel template selected hai ya nahi, aur Sheet Name non-empty string hai ya nahi.
- **Kyun Karta Hai?** Template UI button ke `[disabled]="!canGenerate"` binding ko boolean state return karta hai.
- **Need Kyun Hai?** Incomplete selections par invalid backend requests rokna.

---

### 2. `generate()`
```typescript
generate(): void {
  if (!this.canGenerate) return;
  this.generating = true;

  const payload = {
    document_id: this.selectedDoc!.document_id,
    template_file: this.templateSel!.template.file_name,
    sheet_name: this.templateSel!.sheetName,
  };

  this.service.generateExcel(payload).subscribe({
    next: (res) => {
      const exportId = res.data.export_id;
      this.service.downloadExcel(exportId).subscribe({
        next: (blob) => {
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `${this.templateSel!.template.template_name}_filled.xlsx`;
          a.click();
          URL.revokeObjectURL(url);
          this.msg.add({
            severity: 'success',
            summary: 'Downloaded!',
            detail: `${res.data.mapping_count} fields mapped, ${res.data.written?.['scalar_values_written'] ?? 0} values written`,
            life: 5000,
          });
          this.generating = false;
        },
        error: (err) => { ... }
      });
    },
    error: (err) => { ... }
  });
}
```
- **Kya Karta Hai?** Excel Generation & Auto-Download Execution Workflow:
  1. `payload` construct karta hai (`document_id`, `template_file`, `sheet_name`).
  2. `service.generateExcel(payload)` hit karta hai aur backend se `export_id` receive karta hai.
  3. Immediately `service.downloadExcel(exportId)` hit karke binary `Blob` fetch karta hai.
  4. In-memory Object URL `URL.createObjectURL(blob)` generate karta hai.
  5. HTML Anchor element (`<a>`) create karke programmatic `.click()` trigger karta hai, jisse browser me file download starts ho jati hai.
  6. `URL.revokeObjectURL(url)` call karke browser memory cleanup karta hai.
  7. PrimeNG Success Toast notification display karta hai mentioning total fields mapped & written.
- **Need Kyun Hai?** Application ka core Main User Action workflow execution.

---
*Created automatically by Antigravity AI Code Assistant.*
