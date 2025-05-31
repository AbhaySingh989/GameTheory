# IPD Simulator UI Theming Strategy (Light Theme)

This document outlines the color palette and theming strategy for the Iterated Prisoner's Dilemma (IPD) simulator's user interface, focusing on a modern, clean light theme.

## Color Palette Definition

| Role                  | Hex Code | Description                                            |
|-----------------------|----------|--------------------------------------------------------|
| Primary Color         | `#007BFF`  | Vibrant blue, for main actions, headers                |
| Secondary Color       | `#6C757D`  | Cool gray, for less prominent elements, borders        |
| Accent Color          | `#17A2B8`  | Teal/cyan, for highlights, selected items              |
| Background Color      | `#F8F9FA`  | Very light gray/off-white, for main window, content    |
| Text Color (Primary)  | `#212529`  | Dark gray (almost black), for main text                |
| Text Color (Secondary)| `#495057`  | Medium gray, for less important text                   |
| Success Color         | `#28A745`  | Green, for positive feedback, 'Cooperate' move         |
| Warning/Neutral Color | `#FFC107`  | Amber/Yellow, for draws, informational messages        |
| Error/Defect Color    | `#DC3545`  | Red, for negative feedback, 'Defect' move              |

## Theming Strategy Application

Below describes how the defined colors will be applied to various UI elements. This strategy aims for clarity, usability, and a modern aesthetic. Qt StyleSheets will be the primary mechanism for applying these themes where possible.

### 1. Main Window
*   **Background:** `Background Color (#F8F9FA)`

### 2. Tabs (QTabWidget)
*   **Active Tab Background:** `Primary Color (#007BFF)`
*   **Active Tab Text:** `#FFFFFF` (White) for good contrast against the primary color.
*   **Inactive Tab Background:** A lighter shade of `Background Color` (e.g., `#E9ECEF`) or `Secondary Color (#6C757D)` if more distinction is needed.
*   **Inactive Tab Text:** `Text Color (Secondary) (#495057)` or `Text Color (Primary) (#212529)` if using a very light inactive background.
*   **Tab Pane Background:** `Background Color (#F8F9FA)`

### 3. GroupBoxes (QGroupBox)
*   **Background:** `Background Color (#F8F9FA)` or transparent to inherit from the parent.
*   **Border:** `Secondary Color (#6C757D)` (potentially a lighter tint like `#CED4DA`).
*   **Title Text Color:** `Text Color (Primary) (#212529)`

### 4. Buttons (QPushButton)
*   **Default Button:**
    *   Background: `Primary Color (#007BFF)`
    *   Text: `#FFFFFF` (White)
    *   Border: `Primary Color (#007BFF)`
    *   Hover Background: `#0069D9` (Slightly lighter/brighter blue)
    *   Pressed Background: `#0056B3` (Slightly darker blue)
*   **Secondary/Utility Button (e.g., Reset, Clear):**
    *   Background: `Secondary Color (#6C757D)`
    *   Text: `#FFFFFF` (White)
    *   Border: `Secondary Color (#6C757D)`
    *   Hover Background: `#5A6268` (Slightly lighter/brighter gray)
    *   Pressed Background: `#545B62` (Slightly darker gray)
*   **Disabled Button:**
    *   Background: `#E9ECEF` (Light gray)
    *   Text: `Text Color (Secondary) (#495057)`
    *   Border: `#CED4DA` (Lighter gray)

### 5. Input Fields (QLineEdit, QSpinBox, QComboBox)
*   **Background:** `#FFFFFF` (White)
*   **Border:** `Secondary Color (#6C757D)` (e.g., `#CED4DA`)
*   **Focused Border:** `Primary Color (#007BFF)` (thin border)
*   **Text Color:** `Text Color (Primary) (#212529)`
*   **QComboBox Arrow Button Background:** `Background Color (#F8F9FA)` or `Secondary Color (#6C757D)`

### 6. Text Elements (QLabel, QTextBrowser)
*   **QLabel (default):**
    *   Text Color: `Text Color (Primary) (#212529)`
    *   Background: Transparent (inherits from parent)
*   **QLabel (secondary text):**
    *   Text Color: `Text Color (Secondary) (#495057)`
*   **QTextBrowser (e.g., for logs, detailed descriptions):**
    *   Background: `#FFFFFF` (White) or a very light shade of `Background Color` (e.g., `#F0F2F4`)
    *   Text Color: `Text Color (Primary) (#212529)`
    *   Border: `Secondary Color (#6C757D)` (e.g., `#CED4DA`)

### 7. VisualizationWidget (Custom widget for game history)
*   **Background:** `Background Color (#F8F9FA)`
*   **History Bars:**
    *   'Cooperate' Move: `Success Color (#28A745)`
    *   'Defect' Move: `Error/Defect Color (#DC3545)`
    *   'Draw'/'Neutral' Outcome (if applicable): `Warning/Neutral Color (#FFC107)`
*   **Text (labels, annotations, scores):** `Text Color (Primary) (#212529)`
*   **Axes/Grid Lines:** `Secondary Color (#6C757D)` (lighter tint, e.g., `#ADB5BD`)

### 8. TableWidget (QTableWidget for displaying strategies, results)
*   **Header Background:** `Secondary Color (#6C757D)` or a muted `Primary Color`.
*   **Header Text:** `Text Color (Primary) (#212529)` or `#FFFFFF` (White) if using a darker header background.
*   **Row Background (Odd):** `Background Color (#F8F9FA)`
*   **Row Background (Even - Alternating):** `#FFFFFF` (White) or a slightly darker shade of `Background Color` (e.g., `#E9ECEF`).
*   **Row Text:** `Text Color (Primary) (#212529)`
*   **Selected Row Background:** `Accent Color (#17A2B8)`
*   **Selected Row Text:** `#FFFFFF` (White)
*   **Grid Lines:** `Secondary Color (#6C757D)` (lighter tint, e.g., `#CED4DA`)
*   **Hover Row Background:** A light shade of `Accent Color` (e.g., `#A0D8E0`) or `Primary Color` (e.g., `#B3D7FF`).

### General Notes
*   **Consistency:** Apply colors consistently across all UI elements.
*   **Contrast:** Ensure sufficient contrast between text and background colors for accessibility (WCAG AA guidelines as a reference).
*   **Feedback States:** Use colors to provide clear visual feedback for user actions (hover, pressed, disabled, success, error, warning).
*   **Qt StyleSheets:** Leverage Qt StyleSheets extensively to implement this theming strategy. This allows for dynamic styling and easier maintenance.
*   **Icons:** Icons should be chosen to complement this color scheme, preferably SVG icons that can be styled with CSS or color-filtered. Consider a set that works well on light backgrounds.
*   **Future Dark Theme:** While this document focuses on a light theme, designing with CSS custom properties (if applicable via a preprocessor or future Qt versions) or a structured stylesheet can make it easier to introduce a dark theme later by redefining the core color variables.
