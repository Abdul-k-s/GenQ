##  GenQ – Revit BOQ Automation Tool

GenQ is a Revit add-in that automates Bill of Quantities (BOQ) extraction directly from models, reducing manual work and improving accuracy.

###  Key Features
- Extract quantities from multiple categories
- Handle linked models
- Export to Excel (EPPlus)
- Tree-based element selection UI

###  Demo
[![Watch the demo](https://img.youtube.com/vi/0UPMaIKpWcA/0.jpg)](https://www.youtube.com/watch?v=0UPMaIKpWcA)
<img width="1024" height="791" alt="1" src="https://github.com/user-attachments/assets/b26cf5d6-1cbd-42a5-b10f-78cd83ddd80c" />
<img width="1024" height="791" alt="2" src="https://github.com/user-attachments/assets/3fed3545-f5b2-4ed6-abbc-eafb807f1e43" />
<img width="1024" height="791" alt="3" src="https://github.com/user-attachments/assets/d39fc6ad-72c7-4cc8-acf0-531d905751e1" />
<img width="966" height="640" alt="4" src="https://github.com/user-attachments/assets/eb9698cc-4d2b-49c7-8138-b179fc26b3ed" />


###  Use Cases
- Quantity Surveying
- Tender preparation
- Cost estimation workflows

---

##  Installation Instructions

### Option 1: Build from Source (Developers)

#### Step 1: Install Prerequisites

2. **Install Autodesk Revit 2025**
   - Ensure Revit 2025 is installed at the default location:
     ```
     C:\Program Files\Autodesk\Revit 2025\
     ```

#### Step 2: Clone or Download the Project

```bash
git clone https://github.com/Abdul-k-s/GenQ.git
```

Or download and extract the ZIP file.

#### Step 3: Build the Project

1. Select **Release** configuration (or Debug for testing)
2. Select **x64** platform
3. Build the solution: `Build > Build Solution` (or press `Ctrl+Shift+B`)

The build process will automatically:
- Compile the add-in
- Copy files to: `%AppData%\Autodesk\Revit\Addins\2025\GenQ\`
- Copy the `.addin` manifest to: `%AppData%\Autodesk\Revit\Addins\2025\`

#### Step 4: Start Revit

1. Launch Revit 2025
2. Look for the **"ITI - GenQ"** tab in the ribbon
3. Click the **"GenQ"** button to start

---

### Option 2: Manual Installation (Pre-built Release)

If you have pre-built files:

#### Step 1: Locate the Add-in Files

You should have these files from the build output:
```
GenQ.dll                    (main add-in)
GenQ.pdb                    (debug symbols - optional)
BOQ_Gen.addin               (manifest file)
CSI-MasterFormat2018.xlsx   (MasterFormat data)
EPPlus.dll                  (Excel library dependency)
```

#### Step 2: Copy to Revit Addins Folder

1. **Create the add-in folder:**
   ```
   %AppData%\Autodesk\Revit\Addins\2025\GenQ\
   ```
   Full path example:
   ```
   C:\Users\<YourUsername>\AppData\Roaming\Autodesk\Revit\Addins\2025\GenQ\
   ```

2. **Copy the following files to the `GenQ` subfolder:**
   - `GenQ.dll`
   - `CSI-MasterFormat2018.xlsx`
   - `EPPlus.dll`
   - All other `.dll` files from the build output

3. **Copy the manifest to the 2025 folder:**
   - Copy `BOQ_Gen.addin` to:
     ```
     %AppData%\Autodesk\Revit\Addins\2025\BOQ_Gen.addin
     ```

#### Step 3: Verify Installation

Your folder structure should look like:
```
%AppData%\Autodesk\Revit\Addins\2025\
├── BOQ_Gen.addin                     ← Manifest file
└── GenQ\                             ← Add-in folder
    ├── GenQ.dll                      ← Main add-in
    ├── CSI-MasterFormat2018.xlsx     ← MasterFormat data
    ├── EPPlus.dll                    ← Excel library
    └── [other dependencies]
```

#### Step 4: Start Revit

1. Launch Revit 2025
2. If prompted about loading the add-in, select **"Always Load"**
3. Find the **"GenQ"** tab in the ribbon

---

##  Usage Guide

### Basic Workflow

1. **Open a Revit Project** with model elements

2. **Click "GenQ"** button in the "ITI - GenQ" ribbon tab

3. **Select Levels:**
   - Choose specific levels or "All Levels"
   - Check "Include Links" to include linked models

4. **Select Elements:**
   - Expand categories in the tree view
   - Check the element types you want to include

5. **Configure Each Type:**
   - **Unit Type:** Area, Length, Volume, Perimeter, or Count
   - **CSI Division:** Select from MasterFormat 2018 divisions
   - **Cost:** Enter unit cost
   - **Description:** Add optional description

6. **Generate BOQ:**
   - Click "OK" to calculate quantities
   - Preview the results
   - Export to Excel

### Supported Element Categories

GenQ supports quantity takeoff for:
- Walls
- Floors
- Roofs
- Doors
- Windows
- Structural Columns
- Structural Foundations
- Structural Framing
- Generic Models
- Railings
- And more...

### Unit Types

| Unit Type | Conversion | Use Case |
|-----------|------------|----------|
| **Area** | ft² → m² | Walls, Floors, Roofs |
| **Volume** | ft³ → m³ | Concrete, Fill |
| **Length** | ft → m | Framing, Pipes |
| **Perimeter** | ft → m | Edge conditions |
| **Count** | No conversion | Doors, Windows, Fixtures |

---

## 🔧 Troubleshooting

### Add-in Not Loading

1. **Check the .addin file path:**
   - Must be in `%AppData%\Autodesk\Revit\Addins\2025\`
   - File must have `.addin` extension (not `.addin.txt`)

2. **Check Assembly path in manifest:**
   - Open `BOQ_Gen.addin` in a text editor
   - Verify the `<Assembly>` path matches your installation:
     ```xml
     <Assembly>GenQ\GenQ.dll</Assembly>
     ```

3. **Check Revit's Add-in Manager:**
   - Go to Add-ins tab > Add-in Manager
   - Look for "GenQ" - check its status

### "Could not load file or assembly"

This usually means a dependency is missing:
1. Ensure all DLLs from the build output are in the `GenQ` folder
2. Check that EPPlus.dll is present
3. Ensure .NET 8.0 Desktop Runtime is installed

### "No valid 3D view found"

GenQ requires a 3D view for level-based filtering:
1. Create a 3D view in your project
2. Ensure it's not a template view
3. Try using "All Levels" option

### EPPlus License Warning

EPPlus 5+ requires a license context. GenQ uses NonCommercial mode:
- This is fine for personal/educational use
- For commercial use, obtain an EPPlus license

---

##  Project Structure

```
GenQ-master/
├── GenQ.sln                          # Visual Studio solution
├── README.md                         # This file
├── LICENSE.txt                       # License information
└── TreeViewWithCheckBoxes/           # Main project folder
    ├── GenQ.csproj                   # Project file (.NET 8.0)
    ├── BOQ_Gen.addin                 # Revit manifest
    ├── Apps.cs                       # IExternalApplication entry point
    ├── Class1.cs                     # IExternalCommand handler
    ├── BOQ_Generator.cs              # Quantity calculation logic
    ├── FooViewModel.cs               # Element tree view model
    ├── LevelViewModel.cs             # Level selection view model
    ├── ExcelGen.cs                   # Excel export (EPPlus)
    ├── ReadMasterFormat.cs           # CSI MasterFormat parser
    ├── Window2.xaml / .cs            # Main UI window
    ├── CSI-MasterFormat2018.xlsx     # MasterFormat data
    └── Resources/                    # Icons
        ├── GenQ.png
        └── GenQLogo.ico
---

##  Version History

### v2.0.0 (Revit 2025)
- **Upgraded to .NET 8.0** for Revit 2025 compatibility
- **Performance optimizations:**
  - HashSet for category filtering (O(1) lookup)
  - Dictionary-based type-to-category mapping
  - Reduced redundant Revit API calls
  - Efficient element deduplication
- **Added unit conversion constants** (replaced magic numbers)
- **Improved error handling** with try-catch blocks
- **Updated EPPlus** to v7.0.10
- **Code cleanup** and documentation

### v1.x (Revit 2024 and earlier)
- .NET Framework 4.8
- Original implementation

---

##  Support

For issues and questions:
- Open an issue on GitHub
- Contact: [abdul.khaled.sultan@gmail.com]

---

**GenQ** - Simplifying quantity takeoffs in Revit
