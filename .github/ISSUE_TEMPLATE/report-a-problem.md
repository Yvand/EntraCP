---
name: Report a problem
about: EntraCP throws an error, or causes an error in SharePoint, or cannot be configured
title: ''
labels: ''
assignees: ''

---

**Describe the problem**
A clear and concise description of what the problem is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '....'
3. Scroll down to '....'
4. See error

**Screenshots**
If applicable, add screenshots to help explain your problem.

**Version of EntraCP:**
Run the script below:
```powershell
$dll = [System.Reflection.Assembly]::Load("Yvand.EntraCP, Version=1.0.0.0, Culture=neutral, PublicKeyToken=65dc6b5903b51636")
Get-ChildItem -Path $dll.Location  | Select-Object -ExpandProperty VersionInfo
```

**Relevant logs:**
Include the relevant messages from the SharePoint logs.  
EntraCP records its activity in the Product/Area "EntraCP".  
You can use [ULS Viewer](https://www.microsoft.com/en-us/download/details.aspx?id=44020&msockid=0e32b08e13e8640e3148a5e312a96567) to easily filter the logs.
