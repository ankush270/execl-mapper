# Frontend Service Documentation: `excel-mapping.service.ts`

Yeh document [`excel-mapping.service.ts`](file:///e:/Projects/excelmapping/frontend/src/app/services/excel-mapping.service.ts) file ka complete detailed technical breakdown hai.

---

## 📌 File Overview
- **Location:** `src/app/services/excel-mapping.service.ts`
- **Type:** Angular `@Injectable({ providedIn: 'root' })` Service
- **Purpose:** Backend REST API (`/api/...`) ke sath HTTP requests (`HttpClient`), FormData upload, ArrayBuffer/Blob binary streams, aur TypeScript Data Models ka central communication hub.

---

## 🔑 Data Models & Interfaces

### 1. `JsonDocument`
```typescript
export interface JsonDocument {
  document_id: string;
  file_name: string;
  created_at: string;
}
```
- **Kyun Need Hai:** Backend se milne wale uploaded JSON files ke metadata record ko type-safe representation dene ke liye.

### 2. `ExcelTemplate`
```typescript
export interface ExcelTemplate {
  template_name: string;
  file_name: string;
}
```
- **Kyun Need Hai:** Server directory me available Excel templates ki list model karne ke liye.

### 3. `GenerateExcelRequest` & `GenerateExcelResponse`
```typescript
export interface GenerateExcelRequest {
  document_id: string;
  template_file: string;
  sheet_name: string;
  mapping_id?: string;
}

export interface GenerateExcelResponse {
  export_id: string;
  file_name: string;
  mapping_id: string | null;
  mapping_count: number;
  written: Record<string, number>;
}
```
- **Kyun Need Hai:** `/api/excel_generate` POST API me request payload send karne aur response summary capture karne ke liye.

---

## ⚙️ Service Methods Breakdown

### 1. `uploadJson(files: File[])`
```typescript
uploadJson(files: File[]): Observable<any>
```
- **Kya Karta Hai?** Multiple selected JSON files ko `FormData` object me append karke `/api/json_upload` POST API request bhejta hai.
- **Kyun Karta Hai?** `FormData` multipart/form-data request construct karta hai jisse browser binary files ko backend service layer me pass kar sake.
- **Need Kyun Hai?** JSON File Upload Feature execution.

---

### 2. `getAllJson()`
```typescript
getAllJson(): Observable<{ success: boolean; data: JsonDocument[] }>
```
- **Kya Karta Hai?** GET `/api/json` HTTP call karke uploaded JSON files ki metadata list fetch karta hai.
- **Kyun Karta Hai?** Returns RxJS Observable array `data: JsonDocument[]`.
- **Need Kyun Hai?** JSON Panel UI me file list load karne ke liye.

---

### 3. `getTemplates()`
```typescript
getTemplates(): Observable<{ success: boolean; data: ExcelTemplate[] }>
```
- **Kya Karta Hai?** GET `/api/excel_templates` HTTP call karke server disk par saved Excel template files fetch karta hai.
- **Need Kyun Hai?** Template Panel UI dropdown / list display karne ke liye.

---

### 4. `getTemplateFile(fileName: string)`
```typescript
getTemplateFile(fileName: string): Observable<ArrayBuffer>
```
- **Kya Karta Hai?** GET `/api/excel_templates/<fileName>` endpoint se raw `.xlsx` / `.xlsm` file buffer as `responseType: 'arraybuffer'` download karta hai.
- **Kyun Karta Hai?** `encodeURIComponent(fileName)` apply karta hai taaki special characters URL me escape ho sakein, aur response type `arraybuffer` specifier se browser binary data load kare.
- **Need Kyun Hai?** SheetJS (`xlsx`) library dwara client-side browser in-memory Excel template preview render karne ke liye.

---

### 5. `generateExcel(payload: GenerateExcelRequest)`
```typescript
generateExcel(payload: GenerateExcelRequest): Observable<{ success: boolean; data: GenerateExcelResponse }>
```
- **Kya Karta Hai?** POST `/api/excel_generate` endpoint ko document ID, template filename, aur sheet name ka payload bhejta hai.
- **Kyun Karta Hai?** Backend server par openpyxl fill operation trigger karke database export ID aur execution stats fetch karta hai.
- **Need Kyun Hai?** Excel Generation Execution step.

---

### 6. `downloadExcel(exportId: string)`
```typescript
downloadExcel(exportId: string): Observable<Blob>
```
- **Kya Karta Hai?** GET `/api/excel_download/<exportId>` se final output file ka binary stream as `responseType: 'blob'` fetch karta hai.
- **Kyun Karta Hai?** Returns RxJS Observable containing binary `Blob` object.
- **Need Kyun Hai?** Generated file ko user ke computer par download karwane ke liye (`URL.createObjectURL(blob)`).

---
*Created automatically by Antigravity AI Code Assistant.*
