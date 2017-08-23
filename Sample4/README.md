# How to run ASP.NET Core app for Office 365 addin

## Prerequisites

- Install Dotnet core 2.0

## Start service to host Office 365 addin

```sh
cd dotnetcore_web
dotnet build
dotnet run
```

## Enable https

- install https package

```sh
dotnet add package Microsoft.AspNetCore.Server.Kestrel.Https
```

- create self-signed certificate by openssl
- add pfx file (certificate and private key) to folder `dotnetcore_web/keys/`
