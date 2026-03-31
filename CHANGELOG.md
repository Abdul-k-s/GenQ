# Changelog

All notable changes to GenQ will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Planned
- PDF export support
- Custom BOQ templates
- Multi-language support
- Revit 2024 backward compatibility

---

## [2.0.0] - 2025-03-25

### Added
- **Revit 2025 Support** - Full compatibility with Autodesk Revit 2025
- **SQL Server Integration** - Export BOQ directly to SQL Server database
- **SQL Connection Dialog** - UI for configuring database connections
- **Connection Settings** - Save and load database connection preferences
- **Installer** - Windows installer for easy deployment

### Changed
- **Upgraded to .NET 8.0** - Migrated from .NET Framework 4.8
- **Performance Optimizations**
  - HashSet for O(1) category filtering
  - Dictionary-based type-to-category mapping
  - Reduced redundant Revit API calls
  - Efficient element deduplication
- **Updated EPPlus** to v7.0.10 for Excel generation
- **Added Microsoft.Data.SqlClient** v5.2.2 for SQL connectivity
- **Replaced magic numbers** with named constants for unit conversions
- **Improved error handling** with comprehensive try-catch blocks

### Fixed
- Element duplication when multiple levels selected
- Memory leaks in large model processing
- Excel export failures with special characters

### Removed
- Legacy .NET Framework 4.8 support
- Deprecated `packages.config` (migrated to PackageReference)
- Obsolete `app.config` file

---

## [1.1.0] - 2024-12-15

### Added
- CSI MasterFormat 2018 integration
- Cost estimation per element type
- Preview window before export
- Level-based filtering
- Linked model support

### Changed
- Improved tree view performance
- Better category organization

### Fixed
- Wall area calculation accuracy
- Roof element detection

---

## [1.0.0] - 2024-10-01

### Added
- Initial release
- Basic BOQ generation from Revit models
- Excel export functionality
- Category selection UI
- Support for common element categories:
  - Walls
  - Floors
  - Ceilings
  - Doors
  - Windows
  - Structural elements
- Unit type selection (Area, Volume, Length, Count)
- Revit 2024 compatibility

---

## Version History Summary

| Version | Date | Revit | .NET | Key Features |
|---------|------|-------|------|--------------|
| 2.0.0 | 2025-03-25 | 2025 | 8.0 | SQL export, Performance, Installer |
| 1.1.0 | 2024-12-15 | 2024 | 4.8 | MasterFormat, Cost, Preview |
| 1.0.0 | 2024-10-01 | 2024 | 4.8 | Initial release |

---

## Upgrade Guide

### From v1.x to v2.0

1. **Uninstall v1.x**
   - Remove files from `%AppData%\Autodesk\Revit\Addins\2024\`

2. **Install .NET 8.0 Runtime** (if not already installed)
   - Usually included with Revit 2025

3. **Install v2.0**
   - Run the installer or copy files to:
   - `%AppData%\Autodesk\Revit\Addins\2025\GenQ\`

4. **Migrate Settings**
   - v2.0 uses new configuration format
   - Re-configure your preferences in the UI

---

[Unreleased]: https://github.com/Abdul-k-s/GenQ/compare/v2.0.0...HEAD
[2.0.0]: https://github.com/Abdul-k-s/GenQ/compare/v1.1.0...v2.0.0
[1.1.0]: https://github.com/Abdul-k-s/GenQ/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/Abdul-k-s/GenQ/releases/tag/v1.0.0
