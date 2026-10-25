# Tools By Sanat

An Excel add-in (Windows desktop Excel) that adds a **Tools By Sanat** tab with report and PO formatting tools.

## Install

Download **`ToolsBySanat-Setup-x.y.z.exe`** from the [latest release](https://github.com/sanatpattnaik1-droid/tools-by-sanat/releases/latest)
and run it once. Install only from this page.

Windows may show "unknown publisher" because the installer is not code-signed. Before you run it, compare its fingerprint with the
SHA-256 printed in the release notes:

```powershell
Get-FileHash .\ToolsBySanat-Setup-x.y.z.exe -Algorithm SHA256
```

If the two values differ, do not run the file.

## Updates

After the first install the add-in checks this repository for new versions when Excel starts. An update is only installed when it
is **digitally signed by the owner's key** (the public key is built into the add-in), comes from this repository's releases, and
matches the signed SHA-256. Anything else is ignored.

## Your data stays on your PC

Email addresses and customer details are not part of the add-in. They are read from two files on each PC
(`settings.csv`, `party_master.csv`, opened with the **Settings** button) and are never uploaded.

## Reporting a problem

Open an issue on this repository. Please do not post customer data, email addresses or GSTIN numbers.
