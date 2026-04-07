# ExcelMacros

![Repo Size](https://img.shields.io/github/repo-size/drussell33/ExcelMacros)
![Last Commit](https://img.shields.io/github/last-commit/drussell33/ExcelMacros)
![Top Language](https://img.shields.io/github/languages/top/drussell33/ExcelMacros)

## Overview

ExcelMacros is a repository of Microsoft Excel VBA modules focused on automating spreadsheet workflows and improving efficiency when working with Excel-based data. The project emphasizes reusable, modular VBA code that can be integrated into multiple workbooks. It serves as a practical toolkit for handling Excel objects such as worksheets, ranges, and shapes.

## Key Features

- VBA modules for Excel automation
- Reusable procedures and helper functions
- Interaction with Excel objects (worksheets, ranges, shapes)
- Code structured for reuse across multiple workbooks
- Focus on maintainability and modular design

## Tech Stack

**Backend / Logic**
- VBA (Visual Basic for Applications)

**Frontend**
- Microsoft Excel (Workbook interface, worksheets, shapes)

**Database**
- None (Excel workbook data model)

**Tools / Services**
- Microsoft Excel
- VBA Editor (VBE)

## Architecture Overview

The project is built around Excel's VBA execution model:

- **Frontend (Excel UI):** Users interact with Excel workbooks, worksheets, and shapes.
- **Backend (VBA Modules):** Logic is implemented in VBA modules that manipulate Excel objects.
- **Data Layer:** Data is stored directly within Excel worksheets.

The structure follows a modular pattern where logic is separated into reusable procedures and functions. VBA modules act similarly to service layers, encapsulating operations such as data manipulation and UI interactions within Excel. There is no external database or API layer—Excel itself acts as both the UI and data store.

## Project Structure

```tree
ExcelMacros/
├── Modules/          # VBA modules containing reusable macro logic
├── Forms/            # (If present) UserForms for UI interactions
├── Workbook/         # Workbook-level event logic
└── README.md         # Project documentation
```

> Note: Structure may vary depending on how modules are exported from Excel.

## Getting Started

### Prerequisites

- Microsoft Excel (Desktop version with VBA support)
- Basic familiarity with VBA

### Installation

```bash
git clone https://github.com/drussell33/ExcelMacros.git
```

1. Open Microsoft Excel
2. Open the VBA Editor (ALT + F11)
3. Import `.bas`, `.cls`, or `.frm` files from the repository into your workbook

### Usage

- Run macros directly from the VBA editor (F5)
- Attach macros to buttons or shapes within Excel
- Call reusable functions from other VBA modules

## Roadmap

- [x] Core VBA macro modules
- [x] Reusable helper functions
- [ ] Additional documentation for each module
- [ ] Example Excel workbook demonstrating usage
- [ ] Expanded utility functions for data processing
- [ ] Improved naming and organization of modules

## Contributing

Contributions are welcome.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes
4. Push to your branch
5. Open a Pull Request

## Screenshots / Demo

_Add screenshots or demo GIFs here to showcase macro functionality._

## Contact

GitHub: https://github.com/drussell33
