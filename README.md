# MilSymbol Extension for LibreOffice

Extension to generate military symbols in LibreOffice.

Uses the [milsymbol](https://github.com/spatialillusions/milsymbol/) library.

## Installing

First, install LibreOffice.

**Windows / macOS**

Download and install LibreOffice from https://www.libreoffice.org/

**Linux**

Install LibreOffice from your package manager. Make sure that the script provider for Javascript is installed.

* Debian/Ubuntu: `sudo apt install libreoffice-script-provider-js`
* Fedora: `sudo dnf install libreoffice-rhino`

Then [download the extension](https://github.com/allotropia/milsym-extension/releases) and install it via `Tools->Extensions` in LibreOffice or simply by opening the `.oxt` file from your file manager.

## Building

Build the extension with:

`./build.sh`

Then install the resulting `milsymbol-extension.oxt` using the LibreOffice extension manager, or use unopkg:

`unopkg add -f milsymbol-extension.oxt`

## Autocomplete suggestions

Install [types-unopy](https://pypi.org/project/types-unopy/) and restart your LSP.
