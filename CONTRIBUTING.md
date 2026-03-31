# Contributing to GenQ

Thank you for your interest in contributing to GenQ! This document provides guidelines and instructions for contributing.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [How to Contribute](#how-to-contribute)
- [Pull Request Process](#pull-request-process)
- [Coding Standards](#coding-standards)
- [Testing Guidelines](#testing-guidelines)

---

## Code of Conduct

By participating in this project, you agree to maintain a respectful and inclusive environment. Please be considerate of others and focus on constructive feedback.

---

## Getting Started

### Prerequisites

- **Windows 10/11** (64-bit)
- **Visual Studio 2022** (v17.8 or later)
  - Workload: ".NET Desktop Development"
- **.NET 8.0 SDK**
- **Autodesk Revit 2025**

### Fork and Clone

1. Fork the repository on GitHub
2. Clone your fork locally:
   ```bash
   git clone https://github.com/YOUR-USERNAME/GenQ.git
   cd GenQ
   ```
3. Add the upstream remote:
   ```bash
   git remote add upstream https://github.com/Abdul-k-s/GenQ.git
   ```

---

## Development Setup

### Step 1: Install Revit API References

The project requires Revit 2025 API DLLs. These should be in the `lib/` folder:
- `RevitAPI.dll`
- `RevitAPIUI.dll`

If missing, copy them from:
```
C:\Program Files\Autodesk\Revit 2025\
```

### Step 2: Open Solution

1. Open `GenQ.sln` in Visual Studio 2022
2. Wait for NuGet packages to restore
3. Select **Release | x64** configuration
4. Build the solution (`Ctrl+Shift+B`)

### Step 3: Test in Revit

1. The build automatically copies files to Revit's Addins folder
2. Launch Revit 2025
3. Find the "ITI - GenQ" tab

---

## How to Contribute

### Reporting Bugs

Before creating a bug report, please check existing issues. When filing a bug:

1. Use a clear, descriptive title
2. Include steps to reproduce
3. Describe expected vs. actual behavior
4. Include Revit version and OS details
5. Attach screenshots or error logs if applicable

### Suggesting Features

Feature requests are welcome! Please:

1. Check if the feature was already requested
2. Describe the use case
3. Explain how it benefits BOQ workflows
4. Consider implementation complexity

### Code Contributions

1. **Pick an Issue**: Look for issues labeled `good first issue` or `help wanted`
2. **Comment**: Let others know you're working on it
3. **Branch**: Create a feature branch from `main`
4. **Code**: Implement your changes
5. **Test**: Test with Revit 2025
6. **PR**: Submit a pull request

---

## Pull Request Process

### Branch Naming

Use descriptive branch names:
- `feature/sql-export` - New features
- `bugfix/excel-crash` - Bug fixes
- `docs/update-readme` - Documentation
- `refactor/boq-generator` - Code refactoring

### Before Submitting

- [ ] Code compiles without errors
- [ ] Tested with Revit 2025
- [ ] No breaking changes to existing functionality
- [ ] Code follows project style guidelines
- [ ] Updated documentation if needed

### PR Description Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Documentation update
- [ ] Refactoring

## Testing
How was this tested?

## Screenshots (if applicable)
```

### Review Process

1. Maintainers will review your PR
2. Address any requested changes
3. Once approved, it will be merged

---

## Coding Standards

### C# Style Guidelines

Follow Microsoft's C# coding conventions:

```csharp
// Use PascalCase for public members
public string ElementName { get; set; }

// Use camelCase for private fields with underscore prefix
private readonly Document _document;

// Use meaningful names
public void CalculateTotalArea() { }  // Good
public void Calc() { }                 // Bad

// Add XML documentation for public methods
/// <summary>
/// Generates BOQ for selected elements.
/// </summary>
/// <param name="elements">List of Revit elements</param>
/// <returns>BOQ data structure</returns>
public BOQData Generate(List<Element> elements)
```

### File Organization

```
TreeViewWithCheckBoxes/
├── Apps.cs                 # Entry point
├── Class1.cs               # Command handlers
├── BOQ_Generator.cs        # Core BOQ logic
├── ViewModels/             # MVVM ViewModels
├── Views/                  # XAML Windows
├── Data/                   # Data access layer
└── Resources/              # Images, icons
```

### XAML Guidelines

- Use meaningful `x:Name` attributes
- Keep code-behind minimal; prefer data binding
- Use resource dictionaries for styles

---

## Testing Guidelines

### Manual Testing Checklist

Before submitting a PR, test these scenarios:

- [ ] Load add-in in Revit 2025
- [ ] Select elements from different categories
- [ ] Test with linked models
- [ ] Export to Excel
- [ ] Test with empty selection
- [ ] Test with large models (performance)

### Test Projects

If you have sample Revit projects, test with:
- Small residential project
- Large commercial project
- Project with linked models

---

## Project Architecture

### Key Components

| File | Purpose |
|------|---------|
| `Apps.cs` | `IExternalApplication` - Ribbon setup |
| `Class1.cs` | `IExternalCommand` - Button click handler |
| `BOQ_Generator.cs` | Quantity calculation logic |
| `FooViewModel.cs` | Element tree view model |
| `ExcelGen.cs` | Excel export using EPPlus |
| `Window2.xaml` | Main UI window |

### Data Flow

```
User Click → Class1.Execute() → Window2 (UI)
                                    ↓
                              FooViewModel (Element Selection)
                                    ↓
                              BOQ_Generator (Calculate Quantities)
                                    ↓
                              ExcelGen (Export)
```

---

## Questions?

- Open a [GitHub Discussion](https://github.com/Abdul-k-s/GenQ/discussions)
- Contact the maintainers

---

Thank you for contributing to GenQ!
