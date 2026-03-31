##  GenQ – Revit BOQ Automation Tool

GenQ is a Revit add-in that automates Bill of Quantities (BOQ) extraction directly from models, reducing manual work and improving accuracy.

###  Key Features
- Extract quantities from multiple categories
- Handle linked models
- Export to Excel (EPPlus)
- Tree-based element selection UI

###  Demo
(Add GIF or video here)

###  Use Cases
- Quantity Surveying
- Tender preparation
- Cost estimation workflows

## 🔧 System Requirements

### For Building (Development)
- **Windows 10/11** (64-bit)
- **Visual Studio 2022** (v17.8 or later recommended)
  - Workloads: ".NET Desktop Development"
- **.NET 8.0 SDK** (included with VS 2022 17.8+)
- **Autodesk Revit 2025** (for API references)

### For Running (End Users)
- **Windows 10/11** (64-bit)
- **Autodesk Revit 2025**
- **.NET 8.0 Desktop Runtime** (usually included with Revit 2025)

---

## 📦 Installation Instructions

### Option 1: Build from Source (Developers)

#### Step 1: Install Prerequisites

1. **Install Visual Studio 2022**
   - Download from: https://visualstudio.microsoft.com/
   - During installation, select:
     - ✅ ".NET Desktop Development" workload
     - ✅ ".NET 8.0 Runtime" (individual component)

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

#### Step 3: Open in Visual Studio

1. Open `GenQ.sln` in Visual Studio 2022
2. Wait for NuGet packages to restore automatically

#### Step 4: Build the Project

1. Select **Release** configuration (or Debug for testing)
2. Select **x64** platform
3. Build the solution: `Build > Build Solution` (or press `Ctrl+Shift+B`)

The build process will automatically:
- Compile the add-in
- Copy files to: `%AppData%\Autodesk\Revit\Addins\2025\GenQ\`
- Copy the `.addin` manifest to: `%AppData%\Autodesk\Revit\Addins\2025\`

#### Step 5: Start Revit

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
3. Find the **"ITI - GenQ"** tab in the ribbon

---

## 🚀 Usage Guide

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

## 📁 Project Structure

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
```

---

## ⚠️ Files to Delete (IMPORTANT - Legacy Files)

**Before building**, delete these obsolete files that cause conflicts:

| File | Location | Why Delete |
|------|----------|------------|
| `packages.config` | `TreeViewWithCheckBoxes/` | Old NuGet format - conflicts with PackageReference |
| `app.config` | `TreeViewWithCheckBoxes/` | Old .NET Framework 4.8 config - not needed for .NET 8 |
| `Levels.cs` | Root folder | Orphan empty class file |

**To delete:**
1. Open File Explorer
2. Navigate to `GenQ-master/TreeViewWithCheckBoxes/`
3. Delete `packages.config` and `app.config`
4. Navigate to `GenQ-master/`
5. Delete `Levels.cs`

Or in PowerShell:
```powershell
cd path\to\GenQ-master
Remove-Item "TreeViewWithCheckBoxes\packages.config" -ErrorAction SilentlyContinue
Remove-Item "TreeViewWithCheckBoxes\app.config" -ErrorAction SilentlyContinue
Remove-Item "Levels.cs" -ErrorAction SilentlyContinue
```

---

## 📝 Version History

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

## 📄 License

See [LICENSE.txt](LICENSE.txt) for license information.

---

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with Revit 2025
5. Submit a pull request

---

## 📧 Support

For issues and questions:
- Open an issue on GitHub
- Contact: [Your contact information]

---

**GenQ** - Simplifying quantity takeoffs in Revit
