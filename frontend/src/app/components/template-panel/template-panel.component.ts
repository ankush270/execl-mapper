import { Component, OnInit, Output, EventEmitter } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { SkeletonModule } from 'primeng/skeleton';
import { TooltipModule } from 'primeng/tooltip';
import { BadgeModule } from 'primeng/badge';
import { RippleModule } from 'primeng/ripple';
import { MessageService } from 'primeng/api';
import { ExcelMappingService, ExcelTemplate } from '../../services/excel-mapping.service';
import { ExcelPreviewComponent } from '../excel-preview/excel-preview.component';

export interface TemplateSelection {
  template: ExcelTemplate;
  sheetName: string;
}

@Component({
  selector: 'app-template-panel',
  standalone: true,
  imports: [
    CommonModule, FormsModule,
    ButtonModule, InputTextModule, SkeletonModule,
    TooltipModule, BadgeModule, RippleModule,
    ExcelPreviewComponent,
  ],
  templateUrl: './template-panel.component.html',
  styleUrls: ['./template-panel.component.scss'],
})
export class TemplatePanelComponent implements OnInit {
  @Output() templateSelected = new EventEmitter<TemplateSelection | null>();

  templates: ExcelTemplate[] = [];
  selectedFile: string | null = null;
  sheetName = '';
  loading = false;

  previewVisible = false;
  previewFile = '';

  constructor(
    private service: ExcelMappingService,
    private msg: MessageService,
  ) {}

  ngOnInit(): void { this.load(); }

  load(): void {
    this.loading = true;
    this.service.getTemplates().subscribe({
      next: (res) => { this.templates = res.data ?? []; this.loading = false; },
      error: () => { this.msg.add({ severity: 'error', summary: 'Error', detail: 'Failed to load templates' }); this.loading = false; },
    });
  }

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

  emitSelection(t: ExcelTemplate): void {
    this.templateSelected.emit(
      this.sheetName.trim()
        ? { template: t, sheetName: this.sheetName.trim() }
        : null
    );
  }

  openPreview(t: ExcelTemplate, event: Event): void {
    event.stopPropagation();
    this.previewFile = t.file_name;
    this.previewVisible = true;
  }
}
