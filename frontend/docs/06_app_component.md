# Frontend Component Documentation: `app.component.ts`

Yeh document [`app.component.ts`](file:///e:/Projects/excelmapping/frontend/src/app/app.component.ts) file ka complete detailed technical breakdown hai.

---

## 📌 Component Overview
- **Location:** `src/app/app.component.ts`
- **Selector:** `<app-root>`
- **Imports:** PrimeNG `ToastModule`, `JsonPanelComponent`, `TemplatePanelComponent`, `ActionBarComponent`.
- **Purpose:** Main Root Layout Shell Component jo left-side JSON panel, middle Template panel, aur bottom Action Bar ke beech Shared Application State ko orchestrate karta hai.

---

## 🔑 Shared State Properties

```typescript
selectedDoc: JsonDocument | null = null;
templateSel: TemplateSelection | null = null;
```

1. `selectedDoc`: Global active selected JSON document state object (`JsonDocument` or `null`).
2. `templateSel`: Global active selected Excel template and target sheet name object (`TemplateSelection` or `null`).

---

## ⚙️ Component Methods Breakdown

### 1. `onDocSelected(doc: JsonDocument | null)`
```typescript
onDocSelected(doc: JsonDocument | null): void {
  this.selectedDoc = doc;
}
```
- **Kya Karta Hai?** `<app-json-panel>` ke `@Output() docSelected` event ko handle karke global `selectedDoc` state update karta hai.
- **Kyun Karta Hai?** Child JsonPanel component se aaye document selection ya deselection (`null`) event ko parent state me store karta hai.
- **Need Kyun Hai?** State sharing between JsonPanel and ActionBar.

---

### 2. `onTemplateSelected(sel: TemplateSelection | null)`
```typescript
onTemplateSelected(sel: TemplateSelection | null): void {
  this.templateSel = sel;
}
```
- **Kya Karta Hai?** `<app-template-panel>` ke `@Output() templateSelected` event ko handle karke global `templateSel` state update karta hai.
- **Kyun Karta Hai?** Selected Excel template file aur input Sheet Name ko parent component state me store karta hai.
- **Need Kyun Hai?** State sharing between TemplatePanel and ActionBar.

---

## 📐 Layout Architecture

`app.component.html` layout structure:
```html
<p-toast></p-toast>
<div class="app-layout">
  <app-json-panel (docSelected)="onDocSelected($event)"></app-json-panel>
  <app-template-panel (templateSelected)="onTemplateSelected($event)"></app-template-panel>
  <app-action-bar [selectedDoc]="selectedDoc" [templateSel]="templateSel"></app-action-bar>
</div>
```
- **Toast:** Global Toast alert notifications display karta hai.
- **Grid Layout:** Flexbox / Grid split view containing JSON Panel on Left, Template Panel on Right, and Fixed Action Bar at Bottom.

---
*Created automatically by Antigravity AI Code Assistant.*
