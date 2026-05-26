# Graph Report - .  (2026-05-25)

## Corpus Check
- 121 files · ~110,267 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 278 nodes · 291 edges · 62 communities (33 shown, 29 thin omitted)
- Extraction: 64% EXTRACTED · 36% INFERRED · 0% AMBIGUOUS · INFERRED: 104 edges (avg confidence: 0.81)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_DNS Enrichment Core (Private Helpers)|DNS Enrichment Core (Private Helpers)]]
- [[_COMMUNITY_DNS Enrichment Bugs & Design Rationale|DNS Enrichment Bugs & Design Rationale]]
- [[_COMMUNITY_NextCloud Integration (Private)|NextCloud Integration (Private)]]
- [[_COMMUNITY_WorkstationReport & NextCloud Public API|WorkstationReport & NextCloud Public API]]
- [[_COMMUNITY_Printer Management|Printer Management]]
- [[_COMMUNITY_DNS Log Analysis Public API|DNS Log Analysis Public API]]
- [[_COMMUNITY_DNS Enrichment SQLite Backend|DNS Enrichment SQLite Backend]]
- [[_COMMUNITY_DNS Log Analysis Private Helpers|DNS Log Analysis Private Helpers]]
- [[_COMMUNITY_Server Inventory Orchestrator|Server Inventory Orchestrator]]
- [[_COMMUNITY_Repository Governance|Repository Governance]]
- [[_COMMUNITY_Printer Registry Pattern|Printer Registry Pattern]]
- [[_COMMUNITY_Module Component 25|Module Component 25]]
- [[_COMMUNITY_Module Component 26|Module Component 26]]
- [[_COMMUNITY_Module Component 27|Module Component 27]]
- [[_COMMUNITY_Module Component 28|Module Component 28]]
- [[_COMMUNITY_Module Component 29|Module Component 29]]
- [[_COMMUNITY_Module Component 30|Module Component 30]]
- [[_COMMUNITY_Module Component 31|Module Component 31]]
- [[_COMMUNITY_Module Component 32|Module Component 32]]
- [[_COMMUNITY_Module Component 33|Module Component 33]]
- [[_COMMUNITY_Module Component 34|Module Component 34]]
- [[_COMMUNITY_Module Component 35|Module Component 35]]
- [[_COMMUNITY_Module Component 36|Module Component 36]]
- [[_COMMUNITY_Module Component 37|Module Component 37]]
- [[_COMMUNITY_Module Component 38|Module Component 38]]
- [[_COMMUNITY_Module Component 39|Module Component 39]]
- [[_COMMUNITY_Module Component 40|Module Component 40]]
- [[_COMMUNITY_Module Component 41|Module Component 41]]
- [[_COMMUNITY_Module Component 42|Module Component 42]]
- [[_COMMUNITY_Module Component 43|Module Component 43]]
- [[_COMMUNITY_Module Component 44|Module Component 44]]
- [[_COMMUNITY_Module Component 45|Module Component 45]]
- [[_COMMUNITY_Module Component 46|Module Component 46]]
- [[_COMMUNITY_Module Component 47|Module Component 47]]
- [[_COMMUNITY_Module Component 48|Module Component 48]]
- [[_COMMUNITY_Module Component 49|Module Component 49]]
- [[_COMMUNITY_Module Component 50|Module Component 50]]
- [[_COMMUNITY_Module Component 59|Module Component 59]]
- [[_COMMUNITY_Module Component 60|Module Component 60]]
- [[_COMMUNITY_Module Component 61|Module Component 61]]

## God Nodes (most connected - your core abstractions)
1. `Get-VBServerInventory()` - 30 edges
2. `Invoke-VBIPEnrichment Function (Orchestrator)` - 19 edges
3. `Invoke-VBIPEnrichment()` - 18 edges
4. `New-VBLayerResult()` - 12 edges
5. `VB.WindowsDNSLogAnalysis Module` - 12 edges
6. `Invoke-VBWorkstationReport()` - 10 edges
7. `Publish to PowerShell Gallery GitHub Actions Workflow` - 9 edges
8. `DNSEnrichment Module Design Specification v2.0 (Final)` - 9 edges
9. `VB.NextCloud Module` - 7 edges
10. `ConvertTo-VBNormalisedMAC()` - 6 edges

## Surprising Connections (you probably didn't know these)
- `VB.ServerInventory Legacy Publish Workflow` --semantically_similar_to--> `Publish to PowerShell Gallery GitHub Actions Workflow`  [INFERRED] [semantically similar]
  VB.ServerInventory/Publishing PS Module from Github/publish-to-psgallery.yml → .github/workflows/publish-modules.yml
- `Test-VBPrivateIP in WindowsDNSLogAnalysis` --semantically_similar_to--> `Test-VBPrivateIP Private Helper`  [INFERRED] [semantically similar]
  VB.WindowsDNSLogAnalysis/README.md → VB.DNSEnrichment/CLAUDE.md
- `DNSEnrichment SQLite Schema (Enrichment + EnrichmentHistory)` --semantically_similar_to--> `Parse-Once SQLite Backend Strategy`  [INFERRED] [semantically similar]
  VB.DNSEnrichment/Documentation/dnsenrichment-module-design-v2.md → VB.WindowsDNSLogAnalysis/Documentation/Module_Design.md
- `Publish to PowerShell Gallery GitHub Actions Workflow` --references--> `VB.DNSEnrichment Module`  [INFERRED]
  .github/workflows/publish-modules.yml → VB.DNSEnrichment/README.md
- `VB.WindowsDNSLogAnalysis Module` --conceptually_related_to--> `VB.DNSEnrichment Module`  [INFERRED]
  VB.WindowsDNSLogAnalysis/README.md → VB.DNSEnrichment/README.md

## Hyperedges (group relationships)
- **DNSEnrichment Passive Layer Functions (AD, DHCP, PTR, ARP)** — get_vbadcomputer, get_vbdhcplease, get_vbptrrecord, get_vbarpentry [EXTRACTED 1.00]
- **DNSEnrichment Active Probe Layers (TCP, HTTP, SNMP, RTSP, mDNS, Switch, OUI)** — get_vbtcpfingerprint, get_vbhttpbanner, get_vbsnmpidentity, get_vbrtspbanner, get_vbmdnsrecord, get_vbswitcharp, get_vbouivendor [EXTRACTED 1.00]
- **All Modules Published to PSGallery via CI/CD** — vb_admintools, vb_dnsenrichment, vb_nextcloud, vb_serverinventory, vb_windowsdnsloganalysis, vb_workstationreport [EXTRACTED 1.00]

## Communities (62 total, 29 thin omitted)

### Community 0 - "DNS Enrichment Core (Private Helpers)"
Cohesion: 0.07
Nodes (20): ConvertTo-VBNormalisedMAC(), Get-VBARPTable(), Invoke-VBLoadOUITable(), Invoke-VBmDNSBrowse(), Invoke-VBSwitchSNMPWalk(), New-VBLayerResult(), Write-VBEnrichmentProgress(), Get-VBADComputer() (+12 more)

### Community 1 - "DNS Enrichment Bugs & Design Rationale"
Cohesion: 0.08
Nodes (33): Bug C-01: TCP Timeout 300ms Too Short, Bug C-02: SQL Injection in SQLite Cache-Load, Bug C-03: mDNS Guard Condition Logically Inverted, Bug C-04: Parallel Runspaces Rebuild All One-Shot Caches, ConvertTo-VBNormalisedMAC Private Helper, DHCP Churn Detection via EnrichmentHistory, VB.DNSEnrichment Developer Fix TODO (v0.4.0 to v0.5.0), VB.DNSEnrichment CLAUDE.md (Developer Guidance) (+25 more)

### Community 2 - "NextCloud Integration (Private)"
Cohesion: 0.10
Nodes (10): New-VBNextcloudFolder(), Set-VBNextcloudFile(), Start-VBNextcloudUpload(), Get-DSRegstatus(), Get-VBNetworkInterface(), Get-VBOneDriveFolderBackupStatus(), Get-VBSyncCenterStatus(), Get-VBUserFolderRedirections() (+2 more)

### Community 3 - "WorkstationReport & NextCloud Public API"
Cohesion: 0.18
Nodes (17): Get-VBNextcloudFiles Function, Get-VBPendingReboot Function, Invoke-VBWorkstationReport Orchestrator Function, Auto-detect Module Directories Pattern, VB.NextCloud Soft Runtime Dependency on WorkstationReport, PowerShell Gallery (PSGallery), PSGallery API Key Secret, Publish to PowerShell Gallery GitHub Actions Workflow (+9 more)

### Community 4 - "Printer Management"
Cohesion: 0.16
Nodes (7): Add-VBUserPrinter(), Get-VBUserPrinterMappings(), Set-VBUserPrinterMigration(), Dismount-VBUserHive(), Mount-VBUserHive(), Update-VBUserPrinterRegistry(), Get-VBUserProfile()

### Community 5 - "DNS Log Analysis Public API"
Cohesion: 0.18
Nodes (13): SHA-256 File Deduplication Pattern, Export-VBDNSLogReport Function, Get-VBDNSLogStatistics Function, Get-VBEnrichmentResult Function, Import-VBDNSLog Function, Invoke-VBDNSLogParser Private (StreamReader + Regex Engine), Invoke-VBDNSLogQuery Function, Invoke-VBSqliteCommand Private Helper (+5 more)

### Community 6 - "DNS Enrichment SQLite Backend"
Cohesion: 0.17
Nodes (6): ConvertFrom-VBSqliteEnrichmentRow(), Format-VBEnrichmentContext(), Invoke-VBSqliteCommand(), Get-VBEnrichmentContext(), Get-VBEnrichmentResult(), Initialize-VBEnrichmentDatabase()

### Community 7 - "DNS Log Analysis Private Helpers"
Cohesion: 0.20
Nodes (5): ConvertFrom-VBDNSName(), Get-VBImportStatus(), Invoke-VBBulkInsert(), Invoke-VBDNSLogParser(), Import-VBDNSLog()

### Community 8 - "Server Inventory Orchestrator"
Cohesion: 0.25
Nodes (4): Get-VBServerInventory(), Get-VBDHCPInformation(), Get-VBWindowsFeaturesInfo(), Get-VBFirewallPortRules()

### Community 10 - "Repository Governance"
Cohesion: 0.50
Nodes (4): Bug Report Issue Template, Contributor Covenant Code of Conduct, Contributing Guide, Feature Request Issue Template

### Community 11 - "Printer Registry Pattern"
Cohesion: 0.67
Nodes (4): Add-VBUserPrinter Function, Mount-VBUserHive / Dismount-VBUserHive Functions, Windows Printer Registry Key Pattern, Set-VBUserPrinterMigration Function

## Knowledge Gaps
- **19 isolated node(s):** `Contributor Covenant Code of Conduct`, `Bug Report Issue Template`, `Feature Request Issue Template`, `Get-VBPendingReboot Function`, `Resolve-VBDeviceClass Function` (+14 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **29 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Get-VBServerInventory()` connect `Server Inventory Orchestrator` to `Module Component 25`, `Module Component 26`, `Module Component 27`, `Module Component 28`, `Module Component 29`, `Module Component 30`, `Module Component 31`, `Module Component 32`, `Module Component 33`, `Module Component 34`, `Module Component 35`, `Module Component 36`, `Module Component 37`, `Module Component 38`, `Module Component 39`, `Module Component 40`, `Module Component 41`, `Module Component 42`, `Module Component 43`, `Module Component 44`, `Module Component 45`, `Module Component 46`, `Module Component 47`, `Module Component 48`, `Module Component 49`, `Module Component 50`?**
  _High betweenness centrality (0.044) - this node is a cross-community bridge._
- **Why does `Invoke-VBIPEnrichment Function (Orchestrator)` connect `DNS Enrichment Bugs & Design Rationale` to `DNS Log Analysis Public API`?**
  _High betweenness centrality (0.024) - this node is a cross-community bridge._
- **Why does `Invoke-VBIPEnrichment()` connect `DNS Enrichment Core (Private Helpers)` to `DNS Enrichment SQLite Backend`?**
  _High betweenness centrality (0.023) - this node is a cross-community bridge._
- **Are the 29 inferred relationships involving `Get-VBServerInventory()` (e.g. with `Get-VBSystemInfo()` and `Get-VBNetworkInformation()`) actually correct?**
  _`Get-VBServerInventory()` has 29 INFERRED edges - model-reasoned connections that need verification._
- **Are the 17 inferred relationships involving `Invoke-VBIPEnrichment()` (e.g. with `Get-VBEnrichmentContext()` and `Invoke-VBSqliteCommand()`) actually correct?**
  _`Invoke-VBIPEnrichment()` has 17 INFERRED edges - model-reasoned connections that need verification._
- **Are the 11 inferred relationships involving `New-VBLayerResult()` (e.g. with `Get-VBADComputer()` and `Get-VBARPEntry()`) actually correct?**
  _`New-VBLayerResult()` has 11 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Contributor Covenant Code of Conduct`, `Bug Report Issue Template`, `Feature Request Issue Template` to the rest of the system?**
  _30 weakly-connected nodes found - possible documentation gaps or missing edges._