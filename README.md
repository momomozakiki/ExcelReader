
1. search_keyword   # user put the keyword that he want to search in the Excel file (example : product, brand, etc...)
2. file_path       # last open or save file path to be saved and load the setting during close or open the application
3. folder_path     # last open or save folder path to be saved and load the setting during close or open the application
4. file_extensions # default is Excel (.xlsx) for file and folder, will be used for specific Excel file maybe will include .csv fil in future
5. row_start       # limit the Excel file row starting search limit 
6. row_end         # limit the Excel file row ending search limit 
7. col_start       # limit the Excel file col starting search limit 
8. col_end         # limit the Excel file col ending search limit 
9. cell_start      # cell start will get from the combination of user input row_start, col_start
10. cell_end       # cell ending will get from the combination of user input row_end, col_end
11. save_folder_path # destination folder to keep the result of keyword with the content in .json file 

# QSettings Key Map Suggestion
* settings.setValue("paths/last_file_path", file_path)
* settings.setValue("paths/last_folder_path", folder_path)
* settings.setValue("paths/save_folder_path", save_folder_path)
* settings.setValue("search/keywords", ",".join(search_keywords))
* settings.setValue("search/file_extensions", file_extensions)
* settings.setValue("search/row_start", row_start)
* settings.setValue("search/row_end", row_end)
* settings.setValue("search/col_start", col_start)
* settings.setValue("search/col_end", col_end)


# How the application work
1. user open the application, the application will load the last save file and folder path and the last search keyword
2. user key in the new search keyword or use the last save keyword. multiple keyword (product, brand, etc...)
3. user select new file or folder to do the keyword searching.
4. when finish searching the found keyword content will show in the text box display in .Json format. (keyword : content)
5. the .json file will be saved at new path that user select (save_folder_path), the file name will be same as selected file name.
6. user select the file from the file path, in this case the application only do searching on one file
7. user select the folder and the application will do the searching all the file in that folder instead of single file
8. Default search area is whole Excel file

# AI prompt
You are helping me build a Python desktop application using PySide6 for the GUI. Please follow these instructions:

* Follow best coding practices 
  * use clean, modular, and maintainable code structures (PEP 8 compliant, proper naming conventions, error handling, etc.).
* Avoid code duplication 
  * if a feature or logic already exists (DRY principle), suggest how to reuse or refactor rather than duplicating code.
* Use QSettings (from PySide6) to persist user configuration, 
  * such as the last-used file or folder paths, and automatically load them on application startup and save them on exit.
* The entire GUI should be built using PySide6 only (no other GUI frameworks).

* Please explain and include code examples where necessary, and ensure all suggestions align with professional development standards.













Here’s a **detailed breakdown of all the container widgets** you’ll find in **Qt Designer (used with PySide6)**, including their **purpose**, **function**, and **real-world use cases**.

---

## 🧱 What Are Container Widgets?

In Qt (and thus PySide6), **container widgets** are UI elements that **hold and organize other widgets**. They are used to structure layouts, manage child widgets, and provide interaction contexts (e.g., tabs, scroll areas).

In Qt Designer, you’ll find these under the **“Containers”** section.

---

## 📦 Full List of PySide6 Container Widgets (from Qt Designer)

1. **Group Box (`QGroupBox`)**
2. **Scroll Area (`QScrollArea`)**
3. **Tool Box (`QToolBox`)**
4. **Tab Widget (`QTabWidget`)**
5. **Stacked Widget (`QStackedWidget`)**
6. **Frame (`QFrame`)**
7. **Widget (`QWidget`)** — generic container
8. **Main Window (`QMainWindow`)**
9. **Dock Widget (`QDockWidget`)**
10. **MDI Area (`QMdiArea`)**

---

## 🔍 Detailed Explanation for Each

### 1. **Group Box (`QGroupBox`)**

* **Function**: Groups related widgets together visually and logically.
* **UI**: Has a title at the top; looks like a bordered frame.
* **Use Case**:

  * Grouping settings (e.g., "Login Info", "Network Options")
  * Can be collapsible (with a checkbox if `checkable=True`)
* **Layout Required**: Yes, set layout inside to hold child widgets.

---

### 2. **Scroll Area (`QScrollArea`)**

* **Function**: Allows a large widget or layout to be scrollable inside a fixed area.
* **UI**: Adds vertical and/or horizontal scrollbars automatically.
* **Use Case**:
  * Displaying large images, forms, or lists.
  * Useful when window size is limited.
* **Note**: You add a widget inside the scroll area, which becomes the scrollable content.

---

### 3. **Tool Box (`QToolBox`)**

* **Function**: Stack of pages where each page has a header (like an accordion).
* **UI**: Looks like collapsible panels, one visible at a time.
* **Use Case**:

  * Sidebar settings with grouped options.
  * Preferences panels with categories.
* **Interaction**: Click header to switch panel.

---

### 4. **Tab Widget (`QTabWidget`)**

* **Function**: Provides a tabbed interface, like a notebook with multiple pages.
* **UI**: Tabs across top (or sides), click to switch pages.
* **Use Case**:

  * Settings, dashboards, multi-step forms.
  * Logical separation of content (e.g., “General | Advanced | Help”)
* **Behavior**:

  * Each tab is a separate QWidget container.
  * You can access and manage tabs dynamically.

---

### 5. **Stacked Widget (`QStackedWidget`)**

* **Function**: Shows one widget at a time from a stack.
* **UI**: No UI by itself – switching is controlled programmatically or with logic.
* **Use Case**:

  * Wizard-style navigation
  * Custom tab-like or menu-driven interfaces (you show/hide based on selection)
* **How to use**:

  * Switch page using `.setCurrentIndex()` or `.setCurrentWidget()`

---

### 6. **Frame (`QFrame`)**

* **Function**: A simple container with optional visual borders.
* **UI**: Often used for visual separation or layout grouping.
* **Use Case**:

  * Dividers between sections
  * Grouping widgets without titles
* **Types**: Can customize border styles (`Box`, `HLine`, `VLine`, `StyledPanel`, etc.)

---

### 7. **Widget (`QWidget`)**

* **Function**: Generic container. All other widgets inherit from QWidget.
* **UI**: Blank space; use to structure other widgets/layouts.
* **Use Case**:

  * When you need a layout space but don't want a visible group box or frame.
  * Custom UI containers or dynamic widget holders.

---

### 8. **Main Window (`QMainWindow`)**

* **Function**: Base container for top-level application windows.
* **Special Feature**: Has built-in layout areas (menu bar, tool bar, central widget, status bar, dock areas).
* **Use Case**:

  * Full desktop apps.
  * You must assign a `centralWidget` where you place your layout.

---

### 9. **Dock Widget (`QDockWidget`)**

* **Function**: Panels that can be docked, floated, or hidden.
* **UI**: Typically used in side panels or toolbars.
* **Use Case**:

  * IDE-like interfaces (e.g., sidebars in Qt Creator or VS Code).
  * Optional tools or info panels (e.g., "Object Inspector", "Log Console").

---

### 10. **MDI Area (`QMdiArea`)**

* **Function**: Multi-document interface container.
* **UI**: Allows multiple subwindows inside a parent window.
* **Use Case**:

  * Applications with multiple open documents/windows (like Photoshop or AutoCAD).
  * You add `QMdiSubWindow` objects inside.

---

## 🧭 Choosing the Right Container

| Goal                                          | Container             |
| --------------------------------------------- | --------------------- |
| Group related settings visually               | `QGroupBox`           |
| Scroll large content                          | `QScrollArea`         |
| Show multiple pages with tabs                 | `QTabWidget`          |
| Switch views programmatically                 | `QStackedWidget`      |
| Expandable side panels                        | `QDockWidget`         |
| Multi-document apps                           | `QMdiArea`            |
| Complex main window with menu/tool/status bar | `QMainWindow`         |
| Simple section or layout area                 | `QFrame` or `QWidget` |

---

## 💡 Tips for Using in Designer

* Always **set a layout inside** the container widget after placing it.
* Use **promote to custom widget** if using advanced subclassed containers.
* Remember that some containers (like `QMainWindow`) are **meant to be top-level windows**, not embedded.

---
