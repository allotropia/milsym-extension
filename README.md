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
