# Class NamedRange

NamedRange

Create, access and modify named ranges in a spreadsheet. Named ranges are ranges that have
associated string aliases. They can be viewed and edited via the Sheets UI under the **Data \>
Named ranges...** menu.  

### Methods

|                                                     Method                                                     |                                        Return type                                         |               Brief description                |
|----------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------|------------------------------------------------|
| [getName()](https://developers.google.com/apps-script/reference/spreadsheet/named-range#getName())             | `String`                                                                                   | Gets the name of this named range.             |
| [getRange()](https://developers.google.com/apps-script/reference/spreadsheet/named-range#getRange())           | [Range](https://developers.google.com/apps-script/reference/spreadsheet/range)             | Gets the range referenced by this named range. |
| [remove()](https://developers.google.com/apps-script/reference/spreadsheet/named-range#remove())               | `void`                                                                                     | Deletes this named range.                      |
| [setName(name)](https://developers.google.com/apps-script/reference/spreadsheet/named-range#setName(String))   | [NamedRange](https://developers.google.com/apps-script/reference/spreadsheet/named-range#) | Sets/updates the name of the named range.      |
| [setRange(range)](https://developers.google.com/apps-script/reference/spreadsheet/named-range#setRange(Range)) | [NamedRange](https://developers.google.com/apps-script/reference/spreadsheet/named-range#) | Sets/updates the range for this named range.   |

## Detailed documentation

### `get``Name()`

Gets the name of this named range.

#### Return


`String` --- the name of this named range

#### Authorization

Scripts that use this method require authorization with one or more of the following [scopes](https://developers.google.com/apps-script/concepts/scopes#setting_explicit_scopes):

- `https://www.googleapis.com/auth/spreadsheets.currentonly`
- `https://www.googleapis.com/auth/spreadsheets`

*** ** * ** ***

### `get``Range()`

Gets the range referenced by this named range.

#### Return


[Range](https://developers.google.com/apps-script/reference/spreadsheet/range) --- the spreadsheet range that is associated with this named range

#### Authorization

Scripts that use this method require authorization with one or more of the following [scopes](https://developers.google.com/apps-script/concepts/scopes#setting_explicit_scopes):

- `https://www.googleapis.com/auth/spreadsheets.currentonly`
- `https://www.googleapis.com/auth/spreadsheets`

*** ** * ** ***

### `remove()`

Deletes this named range.

```javascript
// The code below deletes all the named ranges in the spreadsheet.
const namedRanges = SpreadsheetApp.getActive().getNamedRanges();
for (let i = 0; i < namedRanges.length; i++) {
  namedRanges[i].remove();
}
```

#### Authorization

Scripts that use this method require authorization with one or more of the following [scopes](https://developers.google.com/apps-script/concepts/scopes#setting_explicit_scopes):

- `https://www.googleapis.com/auth/spreadsheets.currentonly`
- `https://www.googleapis.com/auth/spreadsheets`

*** ** * ** ***

### `set``Name(name)`

Sets/updates the name of the named range.

```javascript
// The code below updates the name for the first named range.
const namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
if (namedRanges.length > 1) {
  namedRanges[0].setName('UpdatedNamedRange');
}
```

#### Parameters

|  Name  |   Type   |           Description            |
|--------|----------|----------------------------------|
| `name` | `String` | The new name of the named range. |

#### Return


[NamedRange](https://developers.google.com/apps-script/reference/spreadsheet/named-range#) --- the range whose name was set by the call

#### Authorization

Scripts that use this method require authorization with one or more of the following [scopes](https://developers.google.com/apps-script/concepts/scopes#setting_explicit_scopes):

- `https://www.googleapis.com/auth/spreadsheets.currentonly`
- `https://www.googleapis.com/auth/spreadsheets`

*** ** * ** ***

### `set``Range(range)`

Sets/updates the range for this named range.

#### Parameters

|  Name   |                                      Type                                      |                        Description                        |
|---------|--------------------------------------------------------------------------------|-----------------------------------------------------------|
| `range` | [Range](https://developers.google.com/apps-script/reference/spreadsheet/range) | The spreadsheet range to associate with this named range. |

#### Return


[NamedRange](https://developers.google.com/apps-script/reference/spreadsheet/named-range#) --- the named range for which the spreadsheet range was set

#### Authorization

Scripts that use this method require authorization with one or more of the following [scopes](https://developers.google.com/apps-script/concepts/scopes#setting_explicit_scopes):

- `https://www.googleapis.com/auth/spreadsheets.currentonly`
- `https://www.googleapis.com/auth/spreadsheets`