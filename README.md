# SolidWorks Configurations to STL Export Macro

## Overview
This SolidWorks VBA macro exports **all configurations** of the active part or assembly to individual STL files.  
It is designed for users who manage multiple configurations (e.g., via Design Tables) and want a quick, automated way to generate STL files for each one.

**Key features:**
- Exports every configuration as an STL file
- Filenames are based on configuration names (invalid characters replaced)
- Creates an `STL_Exports` folder next to your model automatically
- Uses your last-set STL export settings (resolution, units, binary/ASCII)
- Supports both parts and assemblies (assemblies exported as one STL file)

---

## Requirements
- SolidWorks (any recent version that supports VBA macros)
- A saved SolidWorks part (`.sldprt`) or assembly (`.sldasm`) file with multiple configurations
- STL export settings already configured in SolidWorks

---

## Installation
1. Download the `config2stl.swp` macro file from this repository.
2. Open SolidWorks and your part/assembly file.
3. Go to **Tools → Macro → Run...**
4. Select the downloaded `config2stl.swp` file and click **Open**.

---

## Usage
1. Before running the macro, set your STL export options:
   - **File → Save As → STL → Options...**
   - Configure resolution, units, and binary/ASCII settings
   - Click **OK** and cancel the Save As dialog (settings will be saved)
2. Run the macro.
3. The macro will:
   - Loop through all configurations
   - Export each as an STL file
   - Place them in a subfolder `STL_Exports` in the same location as your model

---

## Customization
- **Output folder name:** Change the `OUTPUT_SUBFOLDER` constant in the macro code.
- **Assembly export mode:** Macro currently exports assemblies as a single STL; this can be modified if needed.
- **Filename sanitization:** The macro replaces invalid filename characters with underscores.

---

## Example
If your file is located at:
```text
C:\Projects\MyPart.SLDPRT
```

After running the macro, you’ll get:
```text
C:\Projects\STL_Exports\
    Config1.stl
    Config2.stl
    Config3.stl
```
