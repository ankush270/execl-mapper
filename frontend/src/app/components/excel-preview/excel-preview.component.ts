import {
  Component, Input, OnChanges, SimpleChanges,
  ViewChild, ElementRef, AfterViewChecked, ChangeDetectorRef, OnDestroy,
  Output, EventEmitter
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { DialogModule } from 'primeng/dialog';
import { ButtonModule } from 'primeng/button';
import { TabViewModule } from 'primeng/tabview';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { ExcelMappingService } from '../../services/excel-mapping.service';
import * as XLSX from 'xlsx';

interface SheetTab {
  name: string;
  html: string;
}

@Component({
  selector: 'app-excel-preview',
  standalone: true,
  imports: [CommonModule, DialogModule, ButtonModule, TabViewModule, ProgressSpinnerModule],
  templateUrl: './excel-preview.component.html',
  styleUrls: ['./excel-preview.component.scss'],
})
export class ExcelPreviewComponent implements OnChanges, OnDestroy {
  @Input() visible = false;
  @Output() visibleChange = new EventEmitter<boolean>();
  @Input() fileName = '';

  sheets: SheetTab[] = [];
  loading = false;
  error = '';

  constructor(private service: ExcelMappingService) {}

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['visible']?.currentValue === true && this.fileName) {
      this.loadPreview();
    }
    if (changes['fileName'] && this.visible && this.fileName) {
      this.loadPreview();
    }
  }

  ngOnDestroy(): void {}

  private loadPreview(): void {
    this.loading = true;
    this.error = '';
    this.sheets = [];

    this.service.getTemplateFile(this.fileName).subscribe({
      next: (buffer) => {
        try {
          const wb = XLSX.read(buffer, {
            type: 'array',
            cellStyles: true,
            cellHTML: false,
            cellDates: true,
          });

          this.sheets = wb.SheetNames.map(name => {
            const ws = wb.Sheets[name];
            const html = this.sheetToStyledHtml(ws, wb);
            return { name, html };
          });

          this.loading = false;
        } catch (err: any) {
          this.error = `Failed to parse Excel file: ${err?.message ?? err}`;
          this.loading = false;
        }
      },
      error: (err) => {
        this.error = `Could not load template: ${err?.error?.message ?? err?.statusText ?? 'Unknown error'}`;
        this.loading = false;
      }
    });
  }

  /** Build HTML table from worksheet with inline styles from cell.s */
  private sheetToStyledHtml(ws: XLSX.WorkSheet, wb: XLSX.WorkBook): string {
    if (!ws['!ref']) return '<p style="padding:1rem;color:#888">Empty sheet</p>';

    const range = XLSX.utils.decode_range(ws['!ref']);
    const merges: XLSX.Range[] = ws['!merge'] || (ws as any)['!merges'] || [];

    // Track merged cell coverage
    const mergedMap = new Map<string, boolean>();
    const mergeSpans = new Map<string, { rs: number; cs: number }>();

    for (const m of merges) {
      const anchorKey = `${m.s.r}_${m.s.c}`;
      mergeSpans.set(anchorKey, { rs: m.e.r - m.s.r + 1, cs: m.e.c - m.s.c + 1 });
      for (let r = m.s.r; r <= m.e.r; r++) {
        for (let c = m.s.c; c <= m.e.c; c++) {
          if (r !== m.s.r || c !== m.s.c) {
            mergedMap.set(`${r}_${c}`, true);
          }
        }
      }
    }

    // Column widths
    const colWidths = ws['!cols'] || [];
    const rowHeights = ws['!rows'] || [];

    let html = '<table>';

    // colgroup for widths
    html += '<colgroup>';
    for (let c = range.s.c; c <= range.e.c; c++) {
      const colInfo = colWidths[c];
      const w = colInfo?.wpx ? `${colInfo.wpx}px` : (colInfo?.wch ? `${colInfo.wch * 7}px` : '80px');
      html += `<col style="width:${w}">`;
    }
    html += '</colgroup>';

    for (let r = range.s.r; r <= range.e.r; r++) {
      const rowInfo = rowHeights[r];
      const h = rowInfo?.hpx ? `height:${rowInfo.hpx}px;` : '';
      html += `<tr style="${h}">`;

      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellKey = `${r}_${c}`;

        // Skip merged non-anchor cells
        if (mergedMap.get(cellKey)) continue;

        const addr = XLSX.utils.encode_cell({ r, c });
        const cell: XLSX.CellObject | undefined = ws[addr];
        const span = mergeSpans.get(cellKey);

        const rowspan = span ? ` rowspan="${span.rs}"` : '';
        const colspan = span ? ` colspan="${span.cs}"` : '';

        const style = this.buildCellStyle(cell, ws);
        const value = cell ? this.formatCellValue(cell) : '';

        html += `<td${rowspan}${colspan} style="${style}">${value}</td>`;
      }

      html += '</tr>';
    }

    html += '</table>';
    return html;
  }

  private buildCellStyle(cell: XLSX.CellObject | undefined, ws: XLSX.WorkSheet): string {
    const styles: string[] = [];

    if (!cell?.s) {
      styles.push('padding:3px 6px');
      return styles.join(';');
    }

    const s: any = cell.s;

    // Background fill
    if (s.fill) {
      const fg = s.fill.fgColor;
      if (fg) {
        const rgb = this.resolveColor(fg);
        if (rgb && rgb !== 'FFFFFF' && rgb !== 'ffffff') {
          styles.push(`background-color:#${rgb}`);
        }
      }
    }

    // Font
    if (s.font) {
      const f = s.font;
      if (f.bold)   styles.push('font-weight:bold');
      if (f.italic) styles.push('font-style:italic');
      if (f.underline) styles.push('text-decoration:underline');
      if (f.sz)     styles.push(`font-size:${Math.round(f.sz * 1.33)}px`);
      if (f.color) {
        const rgb = this.resolveColor(f.color);
        if (rgb) styles.push(`color:#${rgb}`);
      }
      if (f.name)   styles.push(`font-family:'${f.name}',Calibri,sans-serif`);
    }

    // Alignment
    if (s.alignment) {
      const a = s.alignment;
      if (a.horizontal) {
        const hMap: Record<string, string> = {
          left: 'left', center: 'center', right: 'right',
          justify: 'justify', general: 'left',
        };
        styles.push(`text-align:${hMap[a.horizontal] ?? 'left'}`);
      }
      if (a.vertical) {
        const vMap: Record<string, string> = {
          top: 'top', center: 'middle', bottom: 'bottom',
        };
        styles.push(`vertical-align:${vMap[a.vertical] ?? 'middle'}`);
      }
      if (a.wrapText) styles.push('white-space:normal');
    }

    // Border
    if (s.border) {
      const b = s.border;
      const bStyles = ['top','bottom','left','right'] as const;
      for (const side of bStyles) {
        const bd = (b as any)[side];
        if (bd?.style) {
          const rgb = bd.color ? this.resolveColor(bd.color) : '000000';
          const cssStyle = bd.style === 'thin' ? 'solid' : bd.style === 'medium' ? 'solid' : bd.style === 'thick' ? 'solid' : 'solid';
          const width = bd.style === 'thin' ? '1px' : bd.style === 'medium' ? '2px' : bd.style === 'thick' ? '3px' : '1px';
          styles.push(`border-${side}:${width} ${cssStyle} #${rgb ?? '000000'}`);
        }
      }
    }

    styles.push('padding:3px 6px');
    return styles.join(';');
  }

  private resolveColor(color: any): string | null {
    if (!color) return null;
    if (color.rgb) return color.rgb.length === 8 ? color.rgb.slice(2) : color.rgb;
    if (color.theme !== undefined) {
      // Basic theme color fallback palette
      const palette = ['FFFFFF','000000','EEECE1','1F497D','4F81BD','C0504D','9BBB59','8064A2','4BACC6','F79646'];
      return palette[color.theme] ?? null;
    }
    return null;
  }

  private formatCellValue(cell: XLSX.CellObject): string {
    if (cell.t === 'z' || cell.v === null || cell.v === undefined) return '';
    if (cell.t === 'e') return '';
    if (cell.w !== undefined) return this.escapeHtml(cell.w);
    return this.escapeHtml(String(cell.v));
  }

  private escapeHtml(s: string): string {
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  onClose(): void {
    this.visible = false;
    this.visibleChange.emit(false);
    this.sheets = [];
    this.error = '';
  }
}
