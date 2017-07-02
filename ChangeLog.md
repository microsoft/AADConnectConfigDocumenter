## Azure AD Connect Configuration Documenter Change Log

All notable changes to AADConnectConfigDocumenter project will be documented in this file. The "Unreleased" section at the top is for keeping track of important changes that might make it to upcoming releases.

### Version [Unreleased]

* Support for adding report metadata to include section / concept contents.

------------

### Version 1.17.0702.0

#### Added
- Added a new section for documenting End-to-End attribute flows.

#### Fixed
- Fixed tool crash when config exports have internal inconsistency when directory schema extension attributes are selected and then removed.

#### Changed
- Table widths in various sections.

------------

### Version 1.17.0412.0

#### Fixed
- Fixed issue where the "Selected Attributes" section of a connector may display incorrect attribute type information.

------------

### Version 1.17.0314.0

#### Added
- Added support for complete documentation of GLDAP, GSQL, PowerShell and WebServices connectors.

#### Changed
- Variablized $syncRuleId and $syncRulePrecedence to improve portability of the sync rule changes script.

------------

### Version 1.17.0208.0

#### Fixed
- Fixed broken bookmarks in TOC Level 3 entries.

------------

### Version 1.17.0207.0

#### Fixed
- Synchronisation Rule changes script is now correctly downloads in IE.
- Corrected Add-ADSyncAttributeFlowMapping cmdlet script for -Source parameter when the Transformation expression involved more than one attribute.

#### Changed
- TOC now formats entries for Additions and Deletions

------------

### Version 1.16.1123.0

#### Added

* Implemented capability to generate PowerShell scripts for deploying synchronisation rule changes.

------------

### Version 1.16.0919.0

#### Fixed

* Fixed missing Outbound Sync Rule Summary sections.

------------

### Version 1.16.0915.0

#### Fixed

* Corrected XPath condition for filtering out disabled sync rules so that it's backward compatible to the AADSync versions which did not have this capability.
* Fixed the issue of Selected Attributes section not showing Import/Export status in the "Flows Configured?" column.

#### Added

* Support for allowing more than one instances running concurrently.

------------

### Version 1.16.0815.0

#### Fixed

* Active Directory connector showing false positives for inclusion / exclusion OUs in some cases. 

------------

### Version 1.16.0812.0


#### Changed

* Disabled Sync Rules are now not listed in the metaverse attribute precedence information or in the rules summary sections. They are only documented in the Details section.

#### Fixed

* Active Directory connector showing false positives for inclusion / exclusion OUs in some cases. 
------------

### Version 1.16.0708.0

#### Added

* The report now provides a check-box to show only changes between "pilot" and "production" environment config.
 
------------

### Version 1.16.0608.0

#### Fixed

* Metaverse Object Type  section now correctly ranks the sync rules on each attribute as per their precedence.
* The source and target attributes in the join rules of outbound sync rules are now in their correct columns.

------------

### Version 1.16.0603.0

#### Added

* The report header section now lists the pilot and production environment config folders used as input to the tool as well as the version of the AAD Connect in these environments.

#### Changed

* Global Settings section now moved from the Metaverse Configuration section to its own section.

#### Fixed

* Metaverse Object Deletion Rules section now does not mark text as updates when there is no change to the config.
* Selected Attributes section of each connector now correctly shows which attributes are configured with Import/Export flows
* The vanity row of an empty table is not printed if it is to appears as a deleted row in the report.

------------

### Version 1.15.1030.0

#### Added

* Baseline version check-in.

------------
