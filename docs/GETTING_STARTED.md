# Getting started

## Requirements

- Windows with the .NET 10 SDK for the WPF application.
- Python and the Backend requirements only if you want to run the exported Python tests.
- Desktop Microsoft Word only for local Word/PDF generation.

The public UI does not need Yardi, BuildingLink, Outlook or SMTP credentials.

## Open the application

Double-click **Open Public Demo.cmd** in the repository root. The launcher builds the entire desktop application and opens it. Registry Sync is selected initially; use the left sidebar to explore the other workspaces.

The launcher puts build outputs under `%LOCALAPPDATA%\PropertyManagerDemo-Build`, outside the source tree. Application data and credential targets use separate demo namespaces.

You can also build and run from PowerShell:

```powershell
dotnet build PropertyManager.App/PropertyManager.App.csproj -c Release
dotnet run --project PropertyManager.App/PropertyManager.App.csproj -c Release
```

These direct commands use the usual local `bin` and `obj` directories, which are ignored by Git.

## Explore a sample campaign

1. Open **Personalized Email** from the sidebar.
2. Create a new draft and give it a name.
3. On **Recipients**, choose **Try sample data**.
4. Review the fictional recipients, then continue through the message and attachment steps.
5. Build previews on **Review & test** and select different messages to inspect their merged values.

The synthetic example includes two units sharing an email address and a print-only recipient. It illustrates why recipient grouping and duplicate handling matter.

## Public-edition capabilities

| Available for exploration | Intentionally blocked |
| --- | --- |
| Workspace navigation, appearance and local editors | Registry/API synchronization |
| Synthetic recipient data, mapping and merge previews | BuildingLink downloads and reminder extraction |
| Address book and envelope/label layout code | Classic Outlook connections |
| Local document-generation workflows when Word is installed | SMTP connections and message submission |
| Protected demo storage and selected offline tests | Windows task execution and direct printer submission |

Blocks are enforced at the live service/interop entry points by [PublicEdition.cs](../PropertyManager.App/PublicEdition.cs), rather than relying on the absence of credentials alone. Direct Python CLI worker execution is also blocked.

This is a source showcase. Adapting the integration code for another organization requires configuring and validating its identities, data formats and service behavior; replacing the example names alone is not sufficient.

## Tests

```powershell
python -m pip install -r PropertyManager.App/Backend/requirements.txt
python -m unittest discover -s tests -q
```

See [Testing](TESTING.md) for what the exported tests cover.
