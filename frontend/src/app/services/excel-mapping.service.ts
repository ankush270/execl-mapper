import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface JsonDocument {
  document_id: string;
  file_name: string;
  created_at: string;
}

export interface ExcelTemplate {
  template_name: string;
  file_name: string;
}

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

@Injectable({ providedIn: 'root' })
export class ExcelMappingService {
  private readonly base = '/api';

  constructor(private http: HttpClient) {}

  uploadJson(files: File[]): Observable<any> {
    const form = new FormData();
    files.forEach(f => form.append('file', f, f.name));
    return this.http.post(`${this.base}/json_upload`, form);
  }

  getAllJson(): Observable<{ success: boolean; data: JsonDocument[] }> {
    return this.http.get<{ success: boolean; data: JsonDocument[] }>(`${this.base}/json`);
  }

  getTemplates(): Observable<{ success: boolean; data: ExcelTemplate[] }> {
    return this.http.get<{ success: boolean; data: ExcelTemplate[] }>(`${this.base}/excel_templates`);
  }

  /** Fetch template file as ArrayBuffer for preview */
  getTemplateFile(fileName: string): Observable<ArrayBuffer> {
    return this.http.get(`${this.base}/excel_templates/${encodeURIComponent(fileName)}`, {
      responseType: 'arraybuffer',
    });
  }

  generateExcel(payload: GenerateExcelRequest): Observable<{ success: boolean; data: GenerateExcelResponse }> {
    return this.http.post<{ success: boolean; data: GenerateExcelResponse }>(
      `${this.base}/excel_generate`,
      payload
    );
  }

  downloadExcel(exportId: string): Observable<Blob> {
    return this.http.get(`${this.base}/excel_download/${exportId}`, {
      responseType: 'blob',
    });
  }
}
