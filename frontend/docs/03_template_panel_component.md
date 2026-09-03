# Frontend Component Documentation: `template-panel.component.ts`

Yeh document [`template-panel.component.ts`](file:///e:/Projects/excelmapping/frontend/src/app/components/template-panel/template-panel.component.ts) file ka complete detailed technical breakdown hai.

---

## 📌 Component Overview
- **Location:** `src/app/components/template-panel/`
- **Selector:** `<app-template-panel>`
- **Imports:** PrimeNG `InputTextModule`, `ButtonModule`, `SkeletonModule`, `BadgeModule`, `TooltipModule`, `ExcelPreviewComponent`.
- **Purpose:** Middle-panel UI for selecting Excel Templates (`.xlsx`, `.xlsm`), entering target Worksheet Sheet Name (e.g. `"Auto Policy Application"`), and triggering the In-Browser Excel Preview Modal.

---

## 🔑 Data Models & State Properties

```typescript
export interface TemplateSelection {
  template: ExcelTemplate;
  sheetName: string;
}

@Output() templateSelected = new EventEmitter<TemplateSelection | null>();

templates: ExcelTemplate[] = [];
selectedFile: string | null = null;
sheetName = '';
loading = false;

previewVisible = false;
previewFile = '';
```

1. `TemplateSelection`: Object containing selected template file object and trimmed worksheet name string.
2. `@Output() templateSelected`: Selection state event emitter for `AppComponent`.
3. `templates`: Available templates list from server (`ExcelTemplate[]`).
4. `selectedFile`: Active template filename string (`"SalesTemplate.xlsx"`).
5. `sheetName`: Two-way databound string (`[(ngModel)]="sheetName"`) for target sheet title.
6. `previewVisible` & `previewFile`: Dialog visibility and active preview file binding for `app-excel-preview`.

---

## ⚙️ Component Methods Breakdown

### 1. `ngOnInit()` & `load()`
```typescript
ngOnInit(): void { this.load(); }

load(): void {
  this.loading = true;
  this.service.getTemplates().subscribe({
    next: (res) => { this.templates = res.data ?? []; this.loading = false; },
    error: () => { this.msg.add({ severity: 'error', summary: 'Error', detail: 'Failed to load templates' }); this.loading = false; },
  });
}
```
- **Kya Karta Hai?** Templates list load karta hai backend `getTemplates()` endpoint se.
- **Kyun Karta Hai?** Loading indicator state toggle karta hai aur error aane par PrimeNG Error Toast bhejta hai.
- **Need Kyun Hai?** Available templates UI cards populate karne ke liye.

---

### 2. `selectTemplate(t: ExcelTemplate)`
```typescript
selectTemplate(t: ExcelTemplate): void {
  if (this.selectedFile === t.file_name) {
    this.selectedFile = null;
    this.sheetName = '';
    this.templateSelected.emit(null);
  } else {
    this.selectedFile = t.file_name;
    this.sheetName = '';
    this.emitSelection(t);
  }
}
```
- **Kya Karta Hai?** Template card click par toggle selection perform karta hai.
- **Kyun Karta Hai?** Pehle se selected template par dobara click karne par deselection karta hai. Nayi template click karne par `selectedFile` set karta hai, `sheetName` reset karta hai, aur `emitSelection(t)` execute karta hai.
- **Need Kyun Hai?** Active Excel Template selection management.

---

### 3. `emitSelection(t: ExcelTemplate)`
```typescript
emitSelection(t: ExcelTemplate): void {
  this.templateSelected.emit(
    this.sheetName.trim()
      ? { template: t, sheetName: this.sheetName.trim() }
      : null
  );
}
```
- **Kya Karta Hai?** Check karta hai ki `sheetName` non-empty string hai ya nahi, fir `TemplateSelection` event emit karta hai.
- **Kyun Karta Hai?** Jab tak user Sheet Name enter nahi kar deta, tab tak `null` emit karta hai taaki action bar "Generate" button disabled rahe.
- **Need Kyun Hai?** Validation state syncing between Template Panel and Action Bar.

---

### 4. `openPreview(t: ExcelTemplate, event: Event)`
```typescript
openPreview(t: ExcelTemplate, event: Event): void {
  event.stopPropagation();
  this.previewFile = t.file_name;
  this.previewVisible = true;
}
```
- **Kya Karta Hai?** Template card ke "Eye / Preview" icon button par click karne par In-Browser Excel Preview Dialog Modal open karta hai.
- **Kyun Karta Hai?** `event.stopPropagation()` execute karta hai taaki preview click karne par parent card template selection toggle na ho.
- **Need Kyun Hai?** Excel Preview Modal Trigger.

---
*Created automatically by Antigravity AI Code Assistant.*
