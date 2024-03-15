# sp-js-provisioning [![version](https://img.shields.io/badge/version-1.2.2-green.svg)](https://semver.org)

## Description

This project is a SharePoint provisioning tool that uses the SharePoint Framework (SPFx) and Patterns & Practices (PnP) to provision SharePoint sites. It includes various handlers for provisioning different SharePoint components such as files, custom actions, and more.

## Installation

To install the project, you need to have Node.js and npm installed on your machine. After that, you can clone the repository and install the dependencies:

```sh
git clone git://github.com/Puzzlepart/pnp-js-provisioning
cd pnp-js-provisioning
npm install
```

### Usage

Add the npm packages to your project

```shell
npm install sp-js-provisioning --save
```

Here is an example of how you might define navigation in a provisioning template, first in XML and then in JSON (used by sp-js-provisioning).

**XML:**

```xml
<pnp:Navigation>
  <pnp:CurrentNavigation NavigationType="Structural">
    <pnp:StructuralNavigation RemoveExistingNodes="true">
      <pnp:NavigationNode Title="Home" Url="{site}" />
      <pnp:NavigationNode Title="About" Url="{site}/about" />
    </pnp:StructuralNavigation>
  </pnp:CurrentNavigation>
</pnp:Navigation>
```

**JSON:**
  
```json
{
  "Navigation": {
    "CurrentNavigation": {
      "NavigationType": "Structural",
      "StructuralNavigation": {
        "RemoveExistingNodes": true,
        "NavigationNode": [
          {
            "Title": "Home",
            "Url": "{site}"
          },
          {
            "Title": "About",
            "Url": "{site}/about"
          }
        ]
      }
    }
  }
}
```

## Contributing

Contributions are welcome. Please open an issue or submit a pull request on the [GitHub repository](https://github.com/Puzzlepart/pnp-js-provisioning).
