# PCF Project Success Guide: Code-Only (VS Code) Environment

This guide documents the specific project structure and configuration required to successfully build and deploy (`pac pcf push`) PCF controls on this machine, avoiding dependencies on Visual Studio Power Platform Tools.

## 1. Project Structure
The key to success is placing source files in a **subfolder**, not the root.

```text
/Root
  ├── .gitignore
  ├── package.json
  ├── pcfconfig.json               <-- REQUIRED
  ├── tsconfig.json                <-- Paths must point to Subfolder
  ├── eslint.config.mjs            <-- REQUIRED (Modern Flat Config)
  ├── MyControl.pcfproj            <-- "Old Style" MSBuild format
  └── /MyControl (Subfolder)       <-- ALL Source Code Goes Here
       ├── ControlManifest.Input.xml
       ├── index.ts
       ├── Component.tsx
       ├── Component.css
       └── /generated              <-- Type definitions
```

## 2. Configuration Files

### A. `pcfconfig.json` (Root)
Likely missing in failed builds. Tells pcf-scripts where to output.
```json
{
    "outDir": "./out/controls"
}
```

### B. `tsconfig.json` (Root)
Must explicitly include sources in the subfolder.
```json
{
    "compilerOptions": {
        "target": "ES2015",
        "module": "es2015",
        "moduleResolution": "node",
        "lib": ["es2015", "dom"],
        "jsx": "react",
        "strict": true,
        "esModuleInterop": true,
        "skipLibCheck": true,
        "forceConsistentCasingInFileNames": true,
        "noImplicitAny": true
    },
    "include": [
        "MyControl/index.ts",        <-- UPDATE THIS PATH
        "MyControl/Component.tsx",   <-- UPDATE THIS PATH
        "MyControl/generated/**/*.ts"
    ]
}
```

### C. `eslint.config.mjs` (Root)
Create this file if missing. Modern `pcf-scripts` strictly enforces linting.
```javascript
import eslintjs from "@eslint/js";
import microsoftPowerApps from "@microsoft/eslint-plugin-power-apps";
import pluginPromise from "eslint-plugin-promise";
import reactPlugin from "eslint-plugin-react";
import globals from "globals";
import typescriptEslint from "typescript-eslint";

export default [
  { ignores: ["**/generated"] },
  eslintjs.configs.recommended,
  ...typescriptEslint.configs.recommendedTypeChecked,
  ...typescriptEslint.configs.stylisticTypeChecked,
  pluginPromise.configs["flat/recommended"],
  microsoftPowerApps.configs.paCheckerHosted,
  reactPlugin.configs.flat.recommended,
  {
    plugins: { "@microsoft/power-apps": microsoftPowerApps },
    languageOptions: {
      globals: { ...globals.browser, ComponentFramework: true },
      parserOptions: {
        ecmaVersion: 2020,
        sourceType: "module",
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // Add "off" rules here if build fails on linting
      "@typescript-eslint/no-explicit-any": "off"
    },
    settings: { react: { version: "detect" } },
  },
];
```

### D. `.pcfproj` (Root)
Do **NOT** use the `<Project Sdk="...">` style. Use this explicit reference style to avoid "MSBuild missing targets" errors.

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <PowerAppsTargetsPath>$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v$(VisualStudioVersion)\PowerApps</PowerAppsTargetsPath>
  </PropertyGroup>

  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props" />
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Pcf.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Pcf.props')" />

  <PropertyGroup>
    <Name>MyControlName</Name>
    <ProjectGuid>[GUID]</ProjectGuid>
    <OutputPath>$(MSBuildThisFileDirectory)out\controls</OutputPath>
    <TargetFrameworkVersion>v4.6.2</TargetFrameworkVersion>
    <TargetFramework>net462</TargetFramework>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.PowerApps.MSBuild.Pcf" Version="1.*" />
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.0" PrivateAssets="All" />
  </ItemGroup>

  <ItemGroup>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\.gitignore" />
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\bin\**" />
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\obj\**" />
    <ExcludeDirectories Include="$(OutputPath)\**" />
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\node_modules\**" />
  </ItemGroup>

  <ItemGroup>
    <None Include="$(MSBuildThisFileDirectory)\**" Exclude="@(ExcludeDirectories)" />
  </ItemGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets" />
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Pcf.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Pcf.targets')" />
</Project>
```

## 3. Required Dependencies
Run this to ensure all linting tools are present:
```bash
npm install --save-dev @eslint/js @microsoft/eslint-plugin-power-apps eslint-plugin-promise eslint-plugin-react globals typescript-eslint
```

## 4. Why This Works
The default `pac pcf init` or older templates might misconfigure the project for a non-VS environment. 
- **Subfolder**: `pcf-scripts` seems to reliably detect the manifest when it's in a subfolder and `None Include` recursively grabs it.
- **Config Files**: Explicitly adding `pcfconfig.json` and `eslint.config.mjs` prevents the build tools from crashing or searching for default configs provided by Visual Studio extensions.
- **Dependencies**: The modern build process fails silently or throws generic errors if ESLint plugins are missing.
