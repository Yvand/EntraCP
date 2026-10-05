# EntraCP for SharePoint Subscription / 2019 / 2016

**Please visit [entracp.yvand.net](https://entracp.yvand.net/) to download EntraCP and find documentation.**

## Latest stable release

![GitHub Release Date](https://img.shields.io/github/release-date/Yvand/AzureCP.svg)
![GitHub release](https://img.shields.io/github/release/Yvand/AzureCP.svg)
![Latest release downloads](https://img.shields.io/github/downloads/Yvand/AzureCP/latest/total.svg)

## Version in development

![GitHub (Pre-)Release Date](https://img.shields.io/github/release-date-pre/Yvand/AzureCP.svg)
![GitHub release](https://img.shields.io/github/release-pre/Yvand/AzureCP.svg)


## Maintainer workflows

- **CI:** pushing a `v`-prefixed version tag creates a draft prerelease. To publish manually from a branch, enable `publish_release` and supply `release_tag` (for example, `v1.0.0.0`). Existing tags must point to the built commit; new release tags target that commit. Reruns refresh release notes and replace matching assets.
- **Prepare test environment:** `repository_dispatch` accepts `client_payload.sharepoint_versions` as a JSON array or JSON-encoded array of `"SE"`, `"2019"`, and `"2016"`, and `client_payload.skip_create_environment` as a boolean. Omitted values default to `["SE"]` and `false`.
- **Optional runner registration:** set repository variable `DTL_REGISTER_RUNNER` to `true`, variable `DTL_REGISTER_RUNNER_SCRIPT_URL` to the registration script URL, and secret `GH_TOKEN_ADD_RUNNER` to enable it. Otherwise registration is skipped.
- **Run Visual Studio tests:** assign a custom environment label to the intended self-hosted Windows x64 runner and supply that label as `runner_environment`. Test binaries and `DTLServer.runsettings` must already exist under `C:\drop\unit-tests`. Environment preparation does not assign this custom label automatically.

## Miscellaneous

![GitHub issues](https://img.shields.io/github/issues/Yvand/AzureCP.svg)
![GitHub](https://img.shields.io/github/license/Yvand/AzureCP.svg)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/Yvand/AzureCP.svg)
