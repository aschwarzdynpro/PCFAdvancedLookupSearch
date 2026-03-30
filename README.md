# PCFAdvancedLookupSearch

A Power Apps Component Framework (PCF) field control that provides a sophisticated search experience over the Dataverse table linked through a lookup field.

## Features

- **Bound to a lookup column** — replaces the standard lookup control on any Dataverse model-driven form
- **Multi-column search** — searches across one or more columns via a configurable semicolon-separated list (e.g. `fullname;emailaddress1;telephone1`)
- **Multi-column display** — shows configurable columns per result row (primary + secondary labels)
- **Debounced live search** — results appear 300 ms after the user stops typing, with a loading indicator
- **Keyboard navigation** — `↑`/`↓` to move between results, `Enter` to select, `Escape` to dismiss
- **Inline clear / edit** — once a value is selected, clear it or start a new search with one click
- **Respects disabled state** — renders read-only when the field is locked on the form

## Control Properties

| Property | Type | Required | Default | Description |
|---|---|---|---|---|
| `lookupField` | `Lookup.Simple` | ✅ | — | The lookup field to bind the control to |
| `searchColumns` | `SingleLine.Text` | ✅ | `name` | Semicolon-separated columns to search in (e.g. `fullname;emailaddress1`) |
| `displayColumns` | `SingleLine.Text` | | _(same as searchColumns)_ | Semicolon-separated columns to display in each result row |
| `placeholderText` | `SingleLine.Text` | | `Search...` | Placeholder shown in the search input |
| `maxResults` | `Whole.None` | | `10` | Maximum number of records returned per search (max 100) |

## Installation (Dataverse Solution)

A pre-built managed Dataverse solution is attached to every [GitHub Release](../../releases) as `PCFAdvancedLookupSearch_managed.zip`. To install it:

1. Go to the [Releases page](../../releases) and download `PCFAdvancedLookupSearch_managed.zip` from the latest release.
2. Open [make.powerapps.com](https://make.powerapps.com) and select your target environment.
3. Navigate to **Solutions** in the left-hand menu.
4. Click **Import solution** and select the downloaded `.zip` file.
5. Follow the import wizard and click **Import** to complete the installation.
6. After the import finishes, the **Advanced Lookup Search** control will be available to add to lookup columns on model-driven app forms.

## Prerequisites

- [Node.js](https://nodejs.org/) ≥ 18
- [Microsoft Power Platform CLI](https://aka.ms/PowerAppsCLI) (`pac`) for solution packaging

## Development

```bash
# Install dependencies
npm install

# Start the test harness (hot-reload)
npm run start:watch

# Production build
npm run build
```

## Deploying to Dataverse

1. Build the control: `npm run build`
2. Create a solution project: `pac solution init --publisher-name <Name> --publisher-prefix <prefix>`
3. Add the control: `pac solution add-reference --path .`
4. Build the solution: `msbuild /t:build /restore`
5. Import the generated `.zip` into your environment via the [Power Apps maker portal](https://make.powerapps.com/)
6. Add the control to a lookup column on a model-driven form and configure the properties