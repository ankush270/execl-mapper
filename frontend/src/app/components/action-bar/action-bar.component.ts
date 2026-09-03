import { Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ButtonModule } from 'primeng/button';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { TooltipModule } from 'primeng/tooltip';
import { RippleModule } from 'primeng/ripple';
import { MessageService } from 'primeng/api';
import { ExcelMappingService, JsonDocument } from '../../services/excel-mapping.service';
import { TemplateSelection } from '../template-panel/template-panel.component';

@Component({
  selector: 'app-action-bar',
  standalone: true,
  imports: [CommonModule, ButtonModule, ProgressSpinnerModule, TooltipModule, RippleModule],
  templateUrl: './action-bar.component.html',
  styleUrls: ['./action-bar.component.scss'],
})
export class ActionBarComponent {
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

  constructor(
    private service: ExcelMappingService,
    private msg: MessageService,
  ) {}

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
          error: (err) => {
            this.msg.add({ severity: 'error', summary: 'Download Failed', detail: err?.statusText ?? 'Unknown error' });
            this.generating = false;
          },
        });
      },
      error: (err) => {
        this.msg.add({ severity: 'error', summary: 'Generation Failed', detail: err?.error?.message ?? err?.statusText ?? 'Unknown error' });
        this.generating = false;
      },
    });
  }
}
