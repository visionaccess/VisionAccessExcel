# VisionAccess Excel Enhancements for NVDA

An NVDA add-on that makes Microsoft Excel faster and easier to use with a screen reader. It brings several JAWS-style Excel features to NVDA.

- **Author:** Mukesh Jain, VisionAccess
- **Email:** info@visionaccess.org
- **Website:** [www.visionaccess.org](https://www.visionaccess.org)
- **Requirements:** NVDA 2023.1 or later, Microsoft Excel 2013 or later (including Microsoft 365)
- **Licence:** GNU General Public License, version 2

## Download

Download the latest `.nvda-addon` file from the [Releases page](../../releases/latest), then press Enter on it to install. NVDA asks you to confirm and then restarts.

## Keyboard commands

All commands work only in Excel. If you are editing a cell, press Escape first.

| Command | What it does |
|---|---|
| Control+Shift+D | Lists every cell with data in the current worksheet, with its value, and says "formula" for formula cells. Press Enter on a cell to move to it. |
| Control+Shift+R | Lists the cells with data in the current row, with column headers. Press Enter to move to a cell. |
| Control+Shift+C | Lists the cells with data in the current column, with row headers. Press Enter to move to a cell. |
| Control+Shift+S | Opens the Sheet Manager: go to, rename, move up or down, protect or unprotect, hide or unhide, insert and delete sheets. |
| NVDA+Alt+T | Reports the name and range of the table at the current cell. In the Create Table or Resize Table dialog, reports the table range. |
| NVDA+Alt+U | Reports the used range of the sheet. |
| NVDA+Alt+S | Reports a summary of the selection: cell count, cells with data, sum, average, minimum and maximum. |

None of NVDA's default commands are changed. You can reassign any of these keys in NVDA menu, Preferences, Input gestures, category "VisionAccess Excel".

## Other features

- **Blank announcement:** NVDA says "blank" when you press Delete on a cell and it becomes empty, and when you delete the last character in a cell, the formula bar or any Excel edit field.
- **Table range in dialogs:** when the Create Table (Control+T) or Resize Table dialog opens, NVDA announces the range the table covers, for example "Table range: A1:D20".
- **Settings:** NVDA menu, Preferences, Settings, "VisionAccess Excel". Turn each announcement on or off, show or hide headers in cell lists, optionally say "blank" when moving to an empty cell, and set the maximum number of cells listed.

## Building from source

The add-on source is in the `addon` folder. To build the package yourself, zip the contents of `addon` (`manifest.ini`, `appModules`, `globalPlugins`, `doc`) and rename the zip file to end in `.nvda-addon`.

Releases are built automatically: pushing a tag such as `v1.0.2` runs the GitHub Actions workflow, which builds the package and attaches it to a new release.

## Feedback

Please report problems or suggestions through the [Issues page](../../issues), or email info@visionaccess.org.
