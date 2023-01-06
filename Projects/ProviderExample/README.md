# provider-example

## Summary

This project shows a minimal setup of an SPFx WebPart with PnPjs and serves as a reference point and best practices for future SPFx WebParts.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.13-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> NodeJs >= 14.19.3

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- Adjust ```"initialPage"``` in ```serve.json``` file to your local workbench
- In the command-line run:
  - ```npm install``` or ```pnpm install```
  - ```gulp serve```

### Additional Commands
- ```npm run prod``` - bundles and packages the SPFx WebPart for deployment

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- Custom Hooks
- PnPjs SharePoint & Graph Provider (as top level provider)
- (minimal) SharePoint Log Framework

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [PnPjs - Getting Started](https://pnp.github.io/pnpjs/getting-started/)
