# Replay the requests EntraCP sends to Microsoft Graph

This folder is a **[Bruno](https://www.usebruno.com/)** collection which contains the typical requests EntraCP sends to Microsoft Graph.  
You can use it to test / replay the requests sent by EntraCP, and compare against the values returned to SharePoint.

## Usage

To use this collection:

- Copy this folder locally
- Rename/copy the file `.env.example` to `.env` and edit it with your tenant data
- Open this collection in **Bruno**
- Select the environment `entracp-environment` and review its variables
- Go to folder **authentication** > **get access token**: Click tab **Auth** and **Get Access Token**

You can now run any request.
