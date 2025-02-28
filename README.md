**GRAPHICAL CONVERSION SCRIPTS AND PROCESSES FLOW SUMMARY**

### A. Run `parse_class_names.py`
- **Summary:** Loops through all the SCIDs and creates available port lookup file.
- **Output:** Flat file of available port names.

### B. Run Conversion Scripts
- **Command:**
  - For one display file: `python 0_run_conversion_main <fileName>`
  - For entire satellite displays: `python process_scid.py <dir>`
- **Summary:** Runs all the processes listed below accordingly and converts the file to SVG.
- **Output:** Converted SVG file.

#### Steps:
1. **Flat to JSON**
   - **Summary:** Parses the flat file and organizes shapes with attributes into JSON.
   - **Inputs:** Flat file name.
   - **Outputs:** JSON file with flat file objects parsed out.

2. **Get Unique Names (Node Script)**
   - **Summary:** Retrieves the `labelId` from SVG creator to map flat file objects to SVG elements.
   - **Inputs:** Flat file name.
   - **Outputs:** Flat file of distinct object names for mapping.

3. **Get Unique Names (Python Script)**
   - **Summary:** Loads the test file and maps connected shapes and ports.
   - **Inputs:** Flat file name.
   - **Outputs:** JSON object names for mapping.

4. **Get Connection location(Python Script)**
   - **Summary:** Updates JSON file with unique names, calculates X, Y coordinates of each port, and finds connecting objects.
   - **Inputs:** JSON with objects, flat file with distinct names.
   - **Outputs:** JSON file with connection info.

5. **Convert ascii flat file to svg (Node Script)**
   - **Summary:** Parses flat file into an SVG to be used in Visio, including connection info in custom properties.
   - **Inputs:** Flat file, JSON file with connection info.
   - **Outputs:** SVG file.

6. **Format Names(Python Script)**
   - **Summary:** Replaces hyphens (`-`) with underscores (`_`) for shapes within the SVG.

### C. Open SVG in Visio
- **Action:** Move objects to fit display and flip misaligned objects.
- **Macro:** Import `LineConnector.bas` and run `RunLines`.
- **Summary:** Connects objects, draws connection lines, and populates dynamic connector properties.
- **Inputs:** SVG file.
- **Outputs:** Visio file.

### D. Run Switch Replacement Macro
- **Macro:** `replaceShapesWithStencilObjects.bas`
- **Summary:** Replaces switches with stencil objects, copies all properties to new objects, and removes old ones.
  **Note:** Line connections with switches will break—manual reconnection is required. Verify switch rotation.
- **Inputs:** SVG file.
- **Outputs:** Visio file.

### E. Export Connection Information
- **Macro:** `ReportPageShapes.bas`
- **Summary:** Dumps JSON file with connection information.
- **Inputs:** SVG file.
- **Outputs:** JSON file.

### F. Nest Parent-Child Relationships
- **Macro:** `NestReadoutandExport.bas`
- **Summary:** Nests parent-to-child relationships, renames readouts, and exports the final SVG and Visio source document.
- **Inputs:** SVG file.
- **Outputs:** SVG and Visio (`.vsd`) files.

### G. Export SVG from Visio

### H. Run Hot Arrow Property Update
- **Script:** `updateHotArrowProps`
- **Summary:** Updates hot arrow property hierarchy.

