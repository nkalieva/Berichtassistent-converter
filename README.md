# Release Instructions

```sh
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

wird verwendet, um das Projekt als ausführbare .exe-Datei zu erstellen, die:

• ohne installiertes .NET auf dem Zielsystem läuft (self-contained),

• in einer einzigen Datei verpackt ist (Single File Deployment),

• für Windows 64-Bit (win-x64) vorgesehen ist,

• im Release-Modus kompiliert wird – optimiert für die endgültige Nutzung.


the output file is located at:
`\bin\Release\net8.0-windows\win-x64\publish\Resave.exe`