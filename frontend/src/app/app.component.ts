import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ToastModule } from 'primeng/toast';
import { JsonPanelComponent } from './components/json-panel/json-panel.component';
import { TemplatePanelComponent, TemplateSelection } from './components/template-panel/template-panel.component';
import { ActionBarComponent } from './components/action-bar/action-bar.component';
import { JsonDocument } from './services/excel-mapping.service';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [
    CommonModule,
    ToastModule,
    JsonPanelComponent,
    TemplatePanelComponent,
    ActionBarComponent,
  ],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  selectedDoc: JsonDocument | null = null;
  templateSel: TemplateSelection | null = null;

  onDocSelected(doc: JsonDocument | null): void {
    this.selectedDoc = doc;
  }

  onTemplateSelected(sel: TemplateSelection | null): void {
    this.templateSel = sel;
  }
}
