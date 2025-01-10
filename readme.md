# Neobroker Portfolio Importer

## Description

This web-scraping tool aims to extract portfolio asset information (such as stocks, cryptos and ETFs) from Scalable Capital and Trade Republic, given that both neobrokers currently do not feature a portfolio value export. The main features are:

- Import of portfolio asset of both Scalable Capital and Trade Republic, incl. assets name, ISINs and assets current value.
- Semi-automatic login option, where login and password are automatically filled if user adds the parameters `login` and `password`. It waits until 2FA login confirmation.
- Save of the imported assets as an Excel, .csv or a simply copy it as a table to the system clipboard.

## Privacy

The code runs locally in the user's machine and imitates, via Chrome WebDriver and the Selenium library, user's behavior and extracts the assets information from Scalable Capital and Trade Republic. No information is collected and send externally.

For security reason, it is recommended to keep the default parameters `login = None` and `password = None`.

# Usage

## C# dependencies

```.ps1
dotnet add package Selenium.WebDriver
dotnet add package Selenium.WebDriver.ChromeDriver
dotnet add package Selenium.WebDriver.GeckoDriver
dotnet add package EPPlus
```

## Methods

### `ScalableCapitalPortfolioImport`

```.cs
ScalableCapitalPortfolioImport(string login = null, string password = null, string fileType = ".xlsx", string outputPath = null, bool returnDataTable = false)
```

<br>

### `TradeRepublicPortfolioImport`

```.cs
TradeRepublicPortfolioImport(string login = null, string password = null, string fileType = ".xlsx", string outputPath = null, bool returnDataTable = false)
```

#### Description

- Scraps and imports portfolio asset information from Scalable Capital and Trade Republic.

#### Parameters

- `login`: _string_, default: _null_. If defined (e.g. `login = "email@email.com"`), login information is automatically filled; otherwise, user needs to manually add them once the WebDriver initiates.
- `password`: _string_, default: _null_. If defined (e.g. `password = "12345"`), password information is automatically filled; otherwise, user needs to manually add them once the WebDriver initiates.
- `fileType`: _string_, options: _".xlsx"_, _".csv"_ and _null_, default: _".xlsx"_. If _null_, imported assets dataset is copied to the system clipboard.
- `outputPath`: _string_, default: _null_. If _null_, imported assets dataset is copied to the system clipboard.
- `returnDataTable`: _bool_, default: _false_. Returns DataTable from method.

<br>

### `SeleniumWebDriverQuit`

```.cs
SeleniumWebDriverQuit()
```

#### Description

- Terminates the WebDriver session.

#### Parameters

- None.

## Code Workflow Example

```.cs
// Scrap, import and save as .csv portfolio asset information from Scalable Capital
var importer = new NeobrokerPortfolioImporter();
importer.ScalableCapitalPortfolioImport(
    login: null,
    password: null,
    fileType: ".xlsx",
    outputPath: Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Downloads",
        "Assets Scalable Capital.xlsx"
    ),
    returnDataTable: false
);

// Scrap, import and save as .csv portfolio asset information from Trade Republic
importer.TradeRepublicPortfolioImport(
    login: null,
    password: null,
    fileType: ".xlsx",
    outputPath: Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Downloads",
        "Assets Trade Republic.xlsx"
    ),
    returnDataTable: false
);

// Quit WebDriver
importer.SeleniumWebDriverQuit();
```

# See also

[pytr](https://github.com/marzzzello/pytr): Use Trade Republic in terminal.
