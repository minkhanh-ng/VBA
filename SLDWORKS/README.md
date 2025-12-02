# SLDWORKS VBA Macros Collection

A collection of VBA macros for automating SOLIDWORKS (SLDWORKS) design tasks, particularly focused on drawing management, model population, and design table operations.

## Overview

This folder contains VBA macro files (`.bas`) and an Excel workbook (`.xlsm`) that work together to automate various SOLIDWORKS workflows. These macros are designed to handle batch operations on 3D models and 2D drawings.

## Files

### Macro Files (.bas)

| File | Description |
|------|-------------|
| `M_CopyWithErrorHandling.bas` | File system operations with error handling for copying folders while detecting file locks |
| `M_DesignTablePopulate.bas` | Populates and manipulates SOLIDWORKS Design Tables from Excel spreadsheets |
| `M_DrawingPopulate.bas` | Automates drawing creation by copying template drawings and updating references |
| `M_DrawingPopulateConfigs.bas` | Advanced drawing population with configuration management, view positioning, and weldment table handling |
| `M_DrawingSheetCopyAndRename.bas` | Copies and renames drawing sheets, updates view names and configurations |
| `M_LinkDesignTableToModel.bas` | Links external Excel design tables to SOLIDWORKS models |
| `M_ModelMassRebuild.bas` | Batch rebuilds SOLIDWORKS models across multiple configurations |
| `M_ModelsPopulate.bas` | Uses Pack and Go to create model variants with renamed components |
| `M_MyCutList.bas` | Inserts weldment cut list tables into drawings |
| `M_UpdateReferenceLink.bas` | Refreshes external references in Excel files across folder hierarchies |
| `M_Utilities.bas` | Common utility functions for folder/file operations |
| `M_ViewScaleProcess.bas` | Manages drawing view scaling and positioning |
| `M_ViewSequentialLabeling.bas` | Automatically labels detail and section views sequentially |
| `M_WeldPartDrawingProp.bas` | Extracts cut-list properties from weldment parts and applies them to drawings |

### Excel Workbook

| File | Description |
|------|-------------|
| `SLDWORKS-CollectionMacro.xlsm` | Excel workbook containing the VBA macros with data input sheets |

## Key Features

### Drawing Automation
- Copy and populate drawings from templates
- Automatically rename sheets and views
- Update view configurations
- Reposition views based on seed drawings
- Sequential labeling for detail and section views

### Model Management
- Pack and Go operations with component renaming
- Design table integration
- Mass rebuild across configurations
- Reference link management

### Weldment Support
- Cut-list property extraction
- Weldment table insertion
- Body-specific property management

## Usage

1. Open `SLDWORKS-CollectionMacro.xlsm` in Excel
2. Configure the data in Sheet1 with your model/drawing information
3. Ensure SOLIDWORKS is running
4. Press `Alt+F11` to open the VBA editor
5. Select the desired macro module
6. Press `F5` to run the selected macro

## Requirements

- Microsoft Excel (2016 or later) with VBA support enabled
- SOLIDWORKS (2018 or later with API access)
- Windows operating system (uses Windows Script Host for file operations)

## Notes

- Most macros expect SOLIDWORKS to be already running
- File paths are often configured in the Excel workbook
- Macros use the SOLIDWORKS API for model/drawing manipulation
- Some macros include delays (`Sleep`) to handle SOLIDWORKS processing time
