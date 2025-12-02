## Weld Part Drawing Properties VBA Macro

This VBA macro (`M_WeldPartDrawingProp.bas`) is designed to work with SolidWorks drawings. It automates the extraction and transfer of custom properties from weldment part cut list items into drawing properties, making it easier to document key attributes such as part numbers, names, materials, and weight.

### Features

- Connects to an active SolidWorks session and the currently open drawing.
- Identifies the first visible component in a drawing view, and opens the associated part.
- Extracts custom properties from the part’s cut list, including:
  - SW-Part Number
  - ITEM NAME
  - RAW PART
  - RAW MATERIAL
  - WEIGHT (KG & LBS)
  - RAW EQUI MATERIAL
- Transfers these properties to the drawing-level custom properties for documentation.
- Handles both single-body and multi-body weldment parts.
- Supports automatic rebuild of the drawing after updating properties.

### How it Works

1. **Initialization**
   - Connects to the running instance of SolidWorks.
   - Accesses the active document (drawing).

2. **Component & Body Extraction**
   - Locates the active view and its visible components.
   - Opens the referenced part in the correct configuration.
   - Identifies the cut list body matching the view.

3. **Custom Property Transfer**
   - Traverses the part’s features to find cut list folders.
   - Extracts custom properties tied to the cut list item.
   - Updates the drawing’s custom properties accordingly.

4. **Rebuild & Logging**
   - Forces a rebuild of the drawing.
   - Optionally, logs activity (logging code commented out).

### Usage

_Summary instructions:_
- Open a drawing in SolidWorks that references a weldment part.
- Run the macro.
- The drawing’s custom properties will be updated based on the weld part/cut list item properties.

_Developer notes:_
- The macro relies on SolidWorks API objects—ensure relevant references are set (`SldWorks`).
- Logging to a text file can be enabled by uncommenting relevant lines.

### Key Subroutines

- `Main`: Orchestrates extraction and property transfer.
- `GetBodies`: Identifies relevant bodies in a view (handles flat pattern views).
- `Edit_Properties`: Edits drawing-level custom properties.
- `S_GetCutListProperties`, `TraverseFeatures`, `DoTheWork`, `GetFeatureCustomProps`: Traverse the model’s features and extract cut list custom properties.

### Example Cut List Properties Transferred

| Drawing Property      | Cut List Property       |
|----------------------|------------------------|
| PartNo               | SW-Part Number         |
| PartName             | ITEM NAME              |
| RawPart              | RAW PART               |
| RawMaterial          | RAW MATERIAL           |
| PartWeightKG         | WEIGHT (in KG)         |
| PartWeightLBS        | WEIGHT LBS             |
| RawMaterialEqui      | RAW EQUI MATERIAL      |

---

**Note:**  
- This macro is tailored for weldment workflows in SolidWorks and may need adjustments for other use cases.
- Error handling for missing views, unsupported configurations, and suppressed features is included.
