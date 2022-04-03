# Automated App Registration Secrets & Certificates Expiration Notifier

## Table of Contents
  - [Summary](#summary)
  - [Applies to](#applies-to)
  - [Prerequisites](#prerequisites)
  - [Solution](#solution)
  - [Version history](#version-history)
  - [How to implement](#how-to-implement)

## Summary

Automate fetching and displaying an overview of all expiring Azure App Registration secrets and certificates within one or more configured tenants with this Power Automate workflow.

**[<img src="https://external-content.duckduckgo.com/iu/?u=https%3A%2F%2Fwww.iconsdb.com%2Ficons%2Fpreview%2Froyal-blue%2Fdata-transfer-download-xxl.png&f=1&nofb=1" alt="Download .sppkg file" style="width:15px;margin-right:10px;"/>__Download the .zip file here!__](AppRegistrationSecretsAndCertificatesExpirationNotifier.zip)**

## Applies to

- [Power Automate](https://powerautomate.microsoft.com/en-us/)
- [Azure App Registrations](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)

## Prerequisites

> - A Power Automate per user or per flow plan that allows you to use the HTTP connector (see [Power Automate pricing](https://powerautomate.microsoft.com/en-us/pricing/))
> - An Outlook Online mailbox to send the notification to
> - An Azure App Registration within the tenant of which you want to receive automated notifications from. This App Registration should have ONE of the following Microsoft Graph Application type (not Delegated) Permissions: Application.Read.All, Application.ReadWrite.All, Directory.Read.All or Directory.AccessAsUser.All

## Solution

Solution|Author(s)
--------|---------
Automated App Registration Secrets & Certificates Expiration Notifier | cup o'365 ([contact](mailto:info@cupo365.gg), [website](https://cupo365.gg/))

## Version history

Version|Date|Comments
-------|----|--------
1.0|April 3, 2022|Initial release

---

## How to implement

TODO