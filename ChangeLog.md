## Azure AD Connect Configuration Documenter Change Log

All notable changes to AADConnectConfigDocumenter project will be documented in this file. The "Unreleased" section at the top is for keeping track of important changes that might make it to upcoming releases.

### Version [Unreleased]

* Support for adding report metadata to include section / concept contents.

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
