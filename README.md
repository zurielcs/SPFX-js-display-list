# js-display-list

## How to start

Installs the required dependencies. This usually takes 1-3 minutes depending on your internet connection.
```
npm install
```

## Run local

Start the local web server & launch the hosted workbench
```
gulp serve
```

## Package the solution - Local

The following commands bundle your client-side solution, it creates the following package: ./sharepoint/solution/js-display-list.sppkg
```
gulp bundle
gulp package-solution
```

## Package the solution - Cloud tenant

The following commands bundle online solution, it creates the following package: ./sharepoint/solution/js-display-list.sppkg
```
gulp bundle --ship
gulp package-solution --ship
```

## Deploy solution

Upload file ./sharepoint/solution/js-display-list.sppkg in AppCatalog on Sharepoint


## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Build your first SharePoint client-side web](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development