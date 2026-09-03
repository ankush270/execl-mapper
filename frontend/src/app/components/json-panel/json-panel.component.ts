import { Component, OnInit, Output, EventEmitter } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ButtonModule } from 'primeng/button';
import { CardModule } from 'primeng/card';
import { FileUploadModule } from 'primeng/fileupload';
import { SkeletonModule } from 'primeng/skeleton';
import { TooltipModule } from 'primeng/tooltip';
import { BadgeModule } from 'primeng/badge';
import { MessageService } from 'primeng/api';
import { ExcelMappingService, JsonDocument } from '../../services/excel-mapping.service';

@Component({
  selector: 'app-json-panel',
  standalone: true,
  imports: [
    CommonModule, FormsModule,
    ButtonModule, CardModule, FileUploadModule,
    SkeletonModule, TooltipModule, BadgeModule,
  ],
  templateUrl: './json-panel.component.html',
  styleUrls: ['./json-panel.component.scss'],
})
export class JsonPanelComponent implements OnInit {
  @Output() docSelected = new EventEmitter<JsonDocument | null>();

  documents: JsonDocument[] = [];
  selectedId: string | null = null;
  loading = false;
  uploading = false;

  constructor(
    private service: ExcelMappingService,
    private msg: MessageService,
  ) {}

  ngOnInit(): void {
    this.load();
  }

  load(): void {
    this.loading = true;
    this.service.getAllJson().subscribe({
      next: (res) => { this.documents = res.data ?? []; this.loading = false; },
      error: () => { this.msg.add({ severity: 'error', summary: 'Error', detail: 'Failed to load JSON files' }); this.loading = false; },
    });
  }

  selectDoc(doc: JsonDocument): void {
    this.selectedId = this.selectedId === doc.document_id ? null : doc.document_id;
    this.docSelected.emit(this.selectedId ? doc : null);
  }

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
}
