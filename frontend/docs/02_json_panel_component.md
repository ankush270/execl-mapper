# Frontend Component Documentation: `json-panel.component.ts`

Yeh document [`json-panel.component.ts`](file:///e:/Projects/excelmapping/frontend/src/app/components/json-panel/json-panel.component.ts) file ka complete detailed technical breakdown hai.

---

## 📌 Component Overview
- **Location:** `src/app/components/json-panel/`
- **Selector:** `<app-json-panel>`
- **Imports:** PrimeNG `ButtonModule`, `FileUploadModule`, `CardModule`, `SkeletonModule`, `BadgeModule`, `TooltipModule`.
- **Purpose:** Left-side UI Panel for JSON File Drag-and-Drop Uploads, JSON Metadata Documents List, Single-Document Selection State, and Parent Component Event Emitter.

---

## 🔑 Component State Properties

```typescript
@Output() docSelected = new EventEmitter<JsonDocument | null>();

documents: JsonDocument[] = [];
selectedId: string | null = null;
loading = false;
uploading = false;
```

1. `@Output() docSelected`: Selected JSON document object parent `AppComponent` ko bhejta hai.
2. `documents`: Backend se aayi uploaded JSON metadata objects ki list (`JsonDocument[]`).
3. `selectedId`: Currently active/clicked document ID string (ya `null`).
4. `loading`: Skeleton loader spinner state variable.
5. `uploading`: File upload progress spinner state variable.

---

## ⚙️ Component Methods Breakdown

### 1. `ngOnInit()` & `load()`
```typescript
ngOnInit(): void { this.load(); }

load(): void {
  this.loading = true;
  this.service.getAllJson().subscribe({
    next: (res) => { this.documents = res.data ?? []; this.loading = false; },
    error: () => { this.msg.add({ severity: 'error', summary: 'Error', detail: 'Failed to load JSON files' }); this.loading = false; },
  });
}
```
- **Kya Karta Hai?** Component initialization (`ngOnInit`) par backend se JSON files ki list load karta hai.
- **Kyun Karta Hai?** `ExcelMappingService.getAllJson()` subscribe karta hai. Loading indicator toggle karta hai aur error aane par PrimeNG `MessageService` toast error alert display karta hai.
- **Need Kyun Hai?** Uploaded JSON files UI par list karne ke liye.

---

### 2. `selectDoc(doc: JsonDocument)`
```typescript
selectDoc(doc: JsonDocument): void {
  this.selectedId = this.selectedId === doc.document_id ? null : doc.document_id;
  this.docSelected.emit(this.selectedId ? doc : null);
}
```
- **Kya Karta Hai?** User ke dwara kisi JSON card par click karne par single-selection toggle karta hai.
- **Kyun Karta Hai?** Agar pehle se selected doc par dobara click ho to deselection (`null`) karta hai. Nayi file click karne par `selectedId` update karta hai aur `@Output() docSelected` event emit karta hai.
- **Need Kyun Hai?** JSON File Selection State management for Excel generation pipeline.

---

### 3. `onUpload(event: any)`
```typescript
onUpload(event: any): void {
  const files: File[] = Array.from(event.files);
  if (!files.length) return;
  this.uploading = true;
  this.service.uploadJson(files).subscribe({
    next: () => {
      this.msg.add({ severity: 'success', summary: 'Uploaded', detail: `${files.length} file(s) uploaded` });
      this.uploading = false;
      this.load();
    },
    error: (err) => {
      this.msg.add({ severity: 'error', summary: 'Upload Failed', detail: err?.error?.message ?? 'Unknown error' });
      this.uploading = false;
    },
  });
}
```
- **Kya Karta Hai?** PrimeNG `<p-fileUpload>` dropzone se file upload event intercept karta hai aur files array backend `service.uploadJson(files)` ko bhejta hai.
- **Kyun Karta Hai?** Uploading status state update karta hai, PrimeNG Success Toast notification triggers karta hai, aur successful upload par `this.load()` call karke document list auto-refresh karta hai.
- **Need Kyun Hai?** Multi-file Drag & Drop upload execution.

---
*Created automatically by Antigravity AI Code Assistant.*
