# MilSymbol Extension for LibreOffice

A LibreOffice extension for generating military symbols in documents and presentations.

This extension uses the [milsymbol](https://github.com/spatialillusions/milsymbol/) library to create NATO standard military symbols directly within LibreOffice applications.

## Installation

### Prerequisites

First, install LibreOffice:

**Windows / macOS**

Download and install LibreOffice from https://www.libreoffice.org/

**Linux**

Install LibreOffice from your package manager. Make sure that the script provider for JavaScript is installed:

* Debian/Ubuntu: `sudo apt install libreoffice-script-provider-js`
* Fedora: `sudo dnf install libreoffice-rhino`

### Extension Installation

1. [Download the extension](https://github.com/allotropia/milsym-extension/releases) from the releases page
2. Install it via `Tools > Extensions` in LibreOffice
3. Alternatively, simply open the `.oxt` file from your file manager

### Examples and documentation

The extension addes a new global menu, left of the Window and Help
menus on all supported platforms. It provides two entries, either
adding a single military symbol as a draw shape, or a group of symbols
in an org chart layout (called an "ORBAT", as shorthand for Order of
Battle chart).

Both result in data-enhanced SVG draw shapes, that store meta data
alongside its graphical representation, when saved to ODF.

When selected, both shapes add a number of context menu entries, for
updating and editing. A single military tactical symbol for both cases
is heavily configurable via the symbol dialog, the orbat tree has its
own tree editing dialog. Besides standard LibreOffice keyboard
shortcuts, the orbat tree dialog supports the following special keys:

| Key          | Behaviour                                             |
|--------------|-------------------------------------------------------|
| Delete       | delete select shapes(s)                               |
| Return       | open symbol dialog for selected shape                 |
| '+'          | add a child symbol below the currently-selected shape |
| '*'          | add currently selected symbol to favourites           |

With Ctrl key pressed:

| Key          | Behaviour                                     |
|--------------|-----------------------------------------------|
| Cursor left  | move selected shapes one hierarchy level up   |
| Cursor right | move selected shapes one hierarchy level down |
| Cursor up    | move selected shapes one row up               |
| Cursor down  | move selected shapes one row down             |
| 'C'          | copy selected lines to clipboard              |
| 'X'          | cut selected lines, and add to clipboard      |
| 'V'          | paste current clipboard at cursor pos         |
| 'Z'          | undo last orbat dialog operation              |
| 'Y'          | redo last orbat dialog operation              |

There is additionally a special feature available in Calc, activated
in the context menu when used on a multi-cell selection: "Generate
Military Symbols from Rows". It can be used to mass-generate military
tactical symbols, by using standardized Symbol Identification Codes
(SIDC) for the symbol details. See file `example_symbols_list.csv`.

## Building from Source

To build the extension from source:

```bash
./build.sh
```

Then install the resulting `milsymbol-extension.oxt` file using one of these methods:

* Using the LibreOffice extension manager: `Tools > Extensions`
* Using unopkg command line tool:
  ```bash
  unopkg add -f milsymbol-extension.oxt
  ```

## Development

### Autocomplete Support

For development with autocomplete suggestions, install [types-unopy](https://pypi.org/project/types-unopy/) and restart your LSP:

```bash
pip install types-unopy
```
