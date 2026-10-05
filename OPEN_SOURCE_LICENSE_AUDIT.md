# Open-source license audit

Audit started 2026-10-05; reconciled 2026-10-06. Scope: current application/build/installer/website source (existing public application repository: https://github.com/unrealvijay1/decision-risk-intelligence-excel; website/release repository: https://github.com/unrealvijay1/telivu), every package in restored project.assets.json, packed payload allowlists, shipped artwork and runtime prerequisites. This is an engineering inventory and compatibility assessment for owner review, not a license grant.

## Decision

**Apache-2.0 approval is not ready.** No application LICENSE exists and none has been added. Most identified code dependencies use permissive licenses, but Office interop terms and artwork ownership remain unresolved. Removing activation is independent of this legal decision. Do not advertise Apache-2.0 permission or source-redistribution/commercial-use rights until the owner resolves these points and approves the license.

## Runtime, installer and assets

| Dependency/component | Current license/evidence | Telivu usage | Redistribution and Apache-2.0 assessment | Attribution/NOTICE and concerns |
| --- | --- | --- | --- | --- |
| Telivu Core/Excel/website source | No selected repository license | First-party simulation, fitting, scenarios, optimizer, SPC, WinForms UI, static site | Owner must approve terms and confirm source ownership; Apache-2.0 remains candidate | Confirm contributor rights; add approved LICENSE and copyright/NOTICE policy later |
| ExcelDna.AddIn 1.9.0 / ExcelDna.Integration 1.9.0 | Zlib in both installed nuspec files; [pinned upstream license](https://github.com/Excel-DNA/ExcelDna/blob/v1.9.0/LICENSE.txt) | Build tasks, native x86/x64 XLL loaders, managed integration and host components | Permissive, appears compatible with Apache-2.0 for Telivu's own code; upstream retains Zlib | Preserve upstream copyright/origin and source notices; do not claim authorship or relicense Excel-DNA; unmodified loaders retained |
| Excel-DNA build tools / LZMA packing | AddIn package Zlib declaration; SevenZip helper inside build tasks | Compression/decompression during packaging and verification; no standalone optimizer or chart library | Tool-only helper; validate upstream embedded third-party notices when redistributing the tool itself | Customer payload does not include build task DLLs; preserve package notices in developer distributions |
| Microsoft.Office.Interop.Excel 16.0.18925.20022 | Installed nuspec has no license field; describes an unsupported Office-assembly repackaging; [package record](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel/16.0.18925.20022) | Compile-only COM interfaces/delegates with EmbedInteropTypes=true; PIA/office DLLs are explicitly excluded from packed/customer dependencies | **Unresolved**. Excluding the DLL does not establish rights for embedded metadata; obtain authoritative Microsoft Office SDK/PIA terms or replace with a verified source in a separate reviewed change | Package author metadata is not proof of permission. Cannot conclude Apache compatibility or binary redistribution rights from this package |
| .NET 10 runtime / WinForms / System.Drawing/GDI+ | [Runtime MIT](https://github.com/dotnet/runtime/blob/main/LICENSE.TXT), [WinForms MIT](https://github.com/dotnet/winforms/blob/main/LICENSE.TXT), runtime third-party notices; Windows system components under platform terms | Shared Windows Desktop framework, native WinForms and custom GDI+ charts | MIT code appears compatible; Windows/Excel remain platform prerequisites; no third-party chart/UI library | Retain Microsoft runtime distribution/license/third-party notices; framework files are supplied by the unmodified runtime installer, not relicensed as Telivu |
| Microsoft Windows Desktop Runtime 10.0.12 x86/x64 bootstrapper EXEs | Official Microsoft signed/hash-pinned binaries in Installer/runtime-manifest.json; [runtime notices](https://github.com/dotnet/runtime/blob/v10.0.0/THIRD-PARTY-NOTICES.TXT) | Bundled prerequisite installers, deployed only when missing | Separate Microsoft distribution, not Apache-covered Telivu source | Preserve original installers and their terms/notice flows; signing and SHA-512 checks retained. Confirm exact shipped patch distribution terms at release approval |
| Inno Setup 7.1 | Installed License.txt and [upstream license](https://github.com/jrsoftware/issrc/blob/main/license.txt) | Build-only compiler plus generated installer/uninstaller engine | Permissive custom terms allow use and redistribution subject to retained notices/origin; appears compatible for separate Telivu source | Do not remove installer engine notices. [Commercial-use purchase request](https://jrsoftware.org/isorder.php) is not strictly mandatory per the publisher; donations count toward its revenue threshold. Compiler activation is unrelated to Telivu activation and must not be altered |
| Twelve embedded PNG ribbon icons | Supplied assets, no provenance/license metadata or manifest found | Native ribbon buttons | **Unresolved ownership**; owner must confirm rights or replace with verified assets before blanket Apache approval | Confirm creator, permitted distribution/modification and whether attribution is required |
| Website SVG brand mark and social.png | Existing artwork, no recorded license/provenance | Website branding/social cards | **Unresolved ownership** | Confirm owner rights; Apache cannot automatically cover trademarks or artwork without permission |
| Results/optimizer WebP captures | Captures of real product development UI with synthetic fixture data; no external screenshot service | Website product illustration | Likely first-party captures but depend on UI/icon rights; no customer data found | Preserve honest captions; owner confirms underlying asset permissions |
| Microsoft Office ribbon imageMso FileProperties, Segoe UI/Arial | Built-in Office/Windows assets, not copied files | About ribbon icon and system fonts | Platform provided; do not claim Apache licensing of Microsoft assets | Office/Windows prerequisites and Microsoft/Excel trademark attribution retained |
| Algorithms/copied source | No imported-source attribution header or source URL found in application .cs scan; maths/SPC/search are implemented in repo | Sampling, fitting, correlation, XmR, optimizer | No separate MathNet/optimization/charting NuGet library. Absence of headers is not ownership proof | Owner confirms source and contributor provenance, especially any adapted implementations; do not copy copyrighted explanatory text |
| Python build scripts and Node/Playwright checks | First-party scripts; Python standard library; Playwright only local optional validation | Build/link check/browser QA | No Python/Node/Playwright package shipped in website or installer | Python/Node tool distributions retain their own notices. Playwright is Apache-2.0; browser binaries retain browser terms. No JS package is bundled into the public site |

## Complete restored NuGet inventory

The following direct/transitive packages are developer/test-only except the Excel-DNA runtime and compile-only PIA already assessed above. Version/license evidence comes from installed nuspec files and each project's restored lock graph. Customer payload allowlists exclude test/build tools. Compatible labels are provisional engineering assessments; upstream components keep their own licenses.

| Package/version | Declared license | Usage/redistribution | Apache-2.0 assessment and notice requirement |
| --- | --- | --- | --- |
| ExcelDna.AddIn/1.9.0 | Zlib | runtime/build | See runtime table |
| ExcelDna.Integration/1.9.0 | Zlib | runtime/build | See runtime table |
| Microsoft.CodeCoverage/17.14.1 | MIT | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| Microsoft.NET.Test.Sdk/17.14.1 | MIT | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| Microsoft.Office.Interop.Excel/16.0.18925.20022 | UNKNOWN | compile-only metadata | See runtime table |
| Microsoft.TestPlatform.ObjectModel/17.14.1 | MIT | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| Microsoft.TestPlatform.TestHost/17.14.1 | MIT | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| Newtonsoft.Json/13.0.3 | MIT | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| coverlet.collector/6.0.4 | MIT | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| xunit.abstractions/2.0.3 | Legacy licenseUrl only; verify pinned 2.0.3 notice before distributing developer binaries | developer/test-only; not customer payload | Version-specific notice unresolved; test-only and not shipped |
| xunit.analyzers/1.18.0 | Apache-2.0 | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| xunit.assert/2.9.3 | Apache-2.0 | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| xunit.core/2.9.3 | Apache-2.0 | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| xunit.extensibility.core/2.9.3 | Apache-2.0 | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| xunit.extensibility.execution/2.9.3 | Apache-2.0 | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| xunit.runner.visualstudio/3.1.4 | Apache-2.0 | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |
| xunit/2.9.3 | Apache-2.0 | developer/test-only; not customer payload | MIT/Apache permissive where declared; preserve copyright, license and any upstream NOTICE in developer distributions |

## Required owner decisions and release gates

1. Resolve Office interop embedded metadata/PIA rights using authoritative terms; do not silently change COM packaging during activation removal.
2. Confirm first-party code/contributor rights and supplied icon/brand ownership or replace unresolved artwork.
3. Review exact runtime/tool notices for any developer distribution and preserve existing vendor installer notices. Add complete customer third-party notices when the approved license/distribution plan is applied.
4. Approve Apache-2.0 or another license explicitly, then add the corresponding LICENSE and any required NOTICE/attributions. This task does not decide the legal license.
5. Authorize an activation-free public release separately. The existing v0.2.1 asset and all release history remain unchanged; public disclosure stays until that release.

No donation provider is configured. No fake payment destination is introduced. Donation-supported funding does not alter any dependency's license terms.
