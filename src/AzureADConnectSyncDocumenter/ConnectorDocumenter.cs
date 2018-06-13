//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ConnectorDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync Connector Configuration Documenter
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace AzureADConnectConfigDocumenter
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Configuration;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The ConnectorDocumenter is an abstract base class for all types of AAD Connect sync Connectors.
    /// </summary>
    internal abstract class ConnectorDocumenter : Documenter
    {
        /// <summary>
        /// The logger context item connector name
        /// </summary>
        protected const string LoggerContextItemConnectorName = "Connector Name";

        /// <summary>
        /// The logger context item connector unique identifier
        /// </summary>
        protected const string LoggerContextItemConnectorGuid = "Connector Guid";

        /// <summary>
        /// The logger context item connector category
        /// </summary>
        protected const string LoggerContextItemConnectorCategory = "Connector Category";

        /// <summary>
        /// The logger context item connector sub type
        /// </summary>
        protected const string LoggerContextItemConnectorSubType = "Connector SubType";

        /// <summary>
        /// The data source object type currently being processed
        /// </summary>
        private string currentDataSourceObjectType;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConnectorDocumenter" /> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The connector name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        protected ConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
        {
            Logger.Instance.WriteMethodEntry("Connector Name: '{0}'. Config Environment: '{1}'.", connectorName, configEnvironment);

            try
            {
                this.PilotXml = pilotXml;
                this.ProductionXml = productionXml;
                this.ConnectorName = connectorName;
                this.Environment = configEnvironment;

                string xpath = "/ma-data[name ='" + this.ConnectorName + "']";
                var connector = configEnvironment == ConfigEnvironment.ProductionOnly ? this.ProductionXml.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(false) + xpath, Documenter.NamespaceManager) : this.PilotXml.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(true) + xpath, Documenter.NamespaceManager);

                this.ConnectorGuid = (string)connector.Element("id");
                this.ConnectorCategory = (string)connector.Element("category");
                this.ConnectorSubType = (string)connector.Element("subtype");

                // Set Logger call context items
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorName, this.ConnectorName);
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorGuid, this.ConnectorGuid);
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorCategory, this.ConnectorCategory);
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorSubType, this.ConnectorSubType);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Connector Name: '{0}'. Config Environment: '{1}'.", connectorName, configEnvironment);
            }
        }

        /// <summary>
        /// Gets the name of the connector.
        /// </summary>
        /// <value>
        /// The name of the connector.
        /// </value>
        public string ConnectorName { get; private set; }

        /// <summary>
        /// Gets the connector unique identifier.
        /// </summary>
        /// <value>
        /// The connector unique identifier.
        /// </value>
        public string ConnectorGuid { get; private set; }

        /// <summary>
        /// Gets the connector category.
        /// </summary>
        /// <value>
        /// The connector category.
        /// </value>
        public string ConnectorCategory { get; private set; }

        /// <summary>
        /// Gets the type of the connector sub.
        /// </summary>
        /// <value>
        /// The type of the connector sub.
        /// </value>
        public string ConnectorSubType { get; private set; }

        /// <summary>
        /// Gets the sync rule xpath.
        /// </summary>
        /// <param name="currentConnectorGuid">The current connector unique identifier.</param>
        /// <param name="direction">The direction.</param>
        /// <param name="reportType">Type of the report.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        /// <returns>The sync rule xpath.</returns>
        protected static string GetSyncRuleXPath(string currentConnectorGuid, SyncRuleDocumenter.SyncRuleDirection direction, SyncRuleDocumenter.SyncRuleReportType reportType, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Current Connector Guid: '{0}'. Sync Rule Direction: '{1}'.  Sync Rule Report Type: '{2}'. Pilot Config: '{3}'.", currentConnectorGuid, direction, reportType, pilotConfig);

            var xpath = Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + currentConnectorGuid + "' and direction = '" + direction.ToString() + "'";
            try
            {
                switch (reportType)
                {
                    case SyncRuleDocumenter.SyncRuleReportType.ProvisioningSection:
                        xpath += " and linkType = 'Provision' " + Documenter.SyncRuleDisabledCondition;
                        break;
                    case SyncRuleDocumenter.SyncRuleReportType.StickyJoinSection:
                        xpath += " and linkType = 'StickyJoin' " + Documenter.SyncRuleDisabledCondition;
                        break;
                    case SyncRuleDocumenter.SyncRuleReportType.ConditionalJoinSection:
                        xpath += " and linkType = 'Join' " + Documenter.SyncRuleDisabledCondition;
                        xpath += " and (count(synchronizationCriteria/conditions/scope) != 0 or  count(relationshipCriteria/conditions/condition) != 0)";
                        break;
                }

                xpath += "]";

                return xpath;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Current Connector Guid: '{0}'. Sync Rule Direction: '{1}'.  Sync Rule Report Type: '{2}'. Pilot Config: '{3}'. XPath: '{4}'.", currentConnectorGuid, direction, reportType, pilotConfig, xpath);
            }
        }

        /// <summary>
        /// Gets the sync rule section title.
        /// </summary>
        /// <param name="reportType">Type of the report.</param>
        /// <returns>
        /// The sync rule section title.
        /// </returns>
        protected static string GetSyncRuleSectionTitle(SyncRuleDocumenter.SyncRuleReportType reportType)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Report Type: '{0}'.", reportType);

            var sectionTitle = "Synchronization Rules";

            try
            {
                switch (reportType)
                {
                    case SyncRuleDocumenter.SyncRuleReportType.ProvisioningSection:
                        sectionTitle = "Provisioning Rules Summary";
                        break;
                    case SyncRuleDocumenter.SyncRuleReportType.StickyJoinSection:
                        sectionTitle = "Sticky Join Rules Summary";
                        break;
                    case SyncRuleDocumenter.SyncRuleReportType.ConditionalJoinSection:
                        sectionTitle = "Conditional Join Rules Summary";
                        break;
                }

                return sectionTitle;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Sync Rule Report Type: '{0}'. Sync Rule Section Title: '{1}'.", reportType, sectionTitle);
            }
        }

        /// <summary>
        /// Gets the type of the run profile step.
        /// </summary>
        /// <param name="runProfileStepType">Type of the run profile step.</param>
        /// <returns>The type of the run profile step.</returns>
        protected static string GetRunProfileStepType(XElement runProfileStepType)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Report Type: '{0}'.", runProfileStepType != null ? (string)runProfileStepType.Attribute("type") : null);

            var stepType = string.Empty;

            try
            {
                if (runProfileStepType != null)
                {
                    var type = ((string)runProfileStepType.Attribute("type") ?? string.Empty).ToUpperInvariant();
                    switch (type)
                    {
                        case "DELTA-IMPORT":
                            {
                                var importType = ((string)runProfileStepType.Element("import-subtype") ?? string.Empty).ToUpperInvariant();
                                stepType = importType == "TO-CS" ? "Delta Import (Stage Only)" : "Delta Import and Delta Synchronization";
                                break;
                            }

                        case "FULL-IMPORT":
                            {
                                var importType = ((string)runProfileStepType.Element("import-subtype") ?? string.Empty).ToUpperInvariant();
                                stepType = importType == "TO-CS" ? "Full Import (Stage Only)" : "Full Import and Delta Synchronization";
                                break;
                            }

                        case "EXPORT":
                            {
                                stepType = "Export";
                                break;
                            }

                        case "FULL-IMPORT-REEVALUATE-RULES":
                            {
                                stepType = "Full Import and Full Synchronization";
                                break;
                            }

                        case "APPLY-RULES":
                            {
                                var subType = ((string)runProfileStepType.Element("apply-rules-subtype") ?? string.Empty).ToUpperInvariant();
                                stepType = subType == "APPLY-PENDING" ? "Delta Synchronization" : subType == "REEVALUATE-FLOW-CONNECTORS" ? "Full Synchronization" : subType;
                                break;
                            }

                        default:
                            {
                                stepType = type;
                                break;
                            }
                    }
                }

                return stepType;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Sync Rule Report Type: '{0}'.", stepType);
            }
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        protected override void WriteSectionHeader(string title, int level)
        {
            this.WriteSectionHeader(title, level, title, this.ConnectorGuid);
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        /// <param name="bookmark">The section bookmark</param>
        protected new void WriteSectionHeader(string title, int level, string bookmark)
        {
            this.WriteSectionHeader(title, level, bookmark, this.ConnectorGuid);
        }

        /// <summary>
        /// Writes the connector report header.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        protected void WriteConnectorReportHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                var sectionTitle = this.ConnectorName + " Connector Configuration";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 2, this.ConnectorName, this.ConnectorGuid);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Writes the connector synchronize rules header.
        /// </summary>
        /// <param name="reportType">Type of the report.</param>
        protected void WriteConnectorSyncRulesHeader(SyncRuleDocumenter.SyncRuleReportType reportType)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Report Type: '{0}'.", reportType);

            try
            {
                var sectionTitle = ConnectorDocumenter.GetSyncRuleSectionTitle(reportType);

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Sync Rule Report Type: '{0}'.", reportType);
            }
        }

        #region Connector Properties

        /// <summary>
        /// Processes the connector properties.
        /// </summary>
        protected void ProcessConnectorProperties()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Properties.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Property Name, 3 = Value

                this.FillConnectorPropertiesDataSet(true);
                this.FillConnectorPropertiesDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintConnectorProperties();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector properties data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorPropertiesDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var setting = (string)connector.Element("name");
                    Documenter.AddRow(table, new object[] { 0, "Connector Name", setting });

                    setting = (string)connector.Element("category");
                    Documenter.AddRow(table, new object[] { 1, "Connector Type", setting });

                    setting = (string)connector.Element("description");
                    Documenter.AddRow(table, new object[] { 2, "Description", setting });

                    setting = (string)connector.Element("subtype");
                    if (!string.IsNullOrEmpty(setting))
                    {
                        Documenter.AddRow(table, new object[] { 3, "Sub Type", setting });
                    }

                    setting = (string)connector.Element("ma-listname");
                    if (!string.IsNullOrEmpty(setting))
                    {
                        Documenter.AddRow(table, new object[] { 4, "List Name", setting });
                    }

                    setting = (string)connector.Element("ma-companyname");
                    if (!string.IsNullOrEmpty(setting))
                    {
                        Documenter.AddRow(table, new object[] { 5, "Company", setting });
                    }

                    setting = (string)connector.Element("creation-time");
                    Documenter.AddRow(table, new object[] { 6, "Creation Time", setting });

                    setting = (string)connector.Element("last-modification-time");
                    Documenter.AddRow(table, new object[] { 7, "Last Modification Time", setting });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the connector properties.
        /// </summary>
        protected void PrintConnectorProperties()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Properties";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 50 }, { "Configuration", 50 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Connector Properties

        #region Provisioning Hierarchy

        /// <summary>
        /// Processes the connector provisioning hierarchy configuration.
        /// </summary>
        protected void ProcessConnectorProvisioningHierarchyConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Provisioning Hierarchy Configuration.");

                this.CreateSimpleSettingsDataSets(2); // 1 = DN Component, 2 = Object Class Mapping

                this.FillConnectorProvisioningHierarchyConfigurationDataSet(true);
                this.FillConnectorProvisioningHierarchyConfigurationDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorProvisioningHierarchyConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector provisioning hierarchy configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorProvisioningHierarchyConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var mappings = config.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']/component_mappings/mapping");

                foreach (var mapping in mappings)
                {
                    Documenter.AddRow(table, new object[] { (string)mapping.Element("dn_component"), (string)mapping.Element("object_class") });
                }

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the connector provisioning hierarchy configuration.
        /// </summary>
        protected void PrintConnectorProvisioningHierarchyConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Provisioning Hierarchy";

                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "DN Component", 50 }, { "Object Class Mapping", 50 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
                else
                {
                    this.WriteContentParagraph("The provisioning hierarchy is not enabled.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Provisioning Hierarchy

        #region Container Selections

        /// <summary>
        /// Processes the connector partitions.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        protected void ProcessConnectorPartitionContainers(string partitionName)
        {
            this.CreateConnectorPartitionContainersDataSets();

            this.FillConnectorPartitionContainersDataSet(partitionName, true);
            this.FillConnectorPartitionContainersDataSet(partitionName, false);

            this.CreateConnectorPartitionContainersDiffgram();

            this.PrintConnectorPartitionContainers();
        }

        /// <summary>
        /// Creates the connector partition containers data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorPartitionContainersDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("Containers") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Container"); // Container
                var column2 = new DataColumn("Setting"); // Include / Exclude
                var column3 = new DataColumn("DNPart1"); // Sort Column 1
                var column4 = new DataColumn("DNPart2"); // Sort Column 2
                var column5 = new DataColumn("DNPart3"); // Sort Column 3
                var column6 = new DataColumn("DNPart4"); // Sort Column 4
                var column7 = new DataColumn("DNPart5"); // Sort Column 5
                var column8 = new DataColumn("DNPart6"); // Sort Column 6
                var column9 = new DataColumn("DNPart7"); // Sort Column 7
                var column10 = new DataColumn("DNPart8"); // Sort Column 8
                var column11 = new DataColumn("DNPart9"); // Sort Column 9
                var column12 = new DataColumn("DNPart10"); // Sort Column 10

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.Columns.Add(column5);
                table.Columns.Add(column6);
                table.Columns.Add(column7);
                table.Columns.Add(column8);
                table.Columns.Add(column9);
                table.Columns.Add(column10);
                table.Columns.Add(column11);
                table.Columns.Add(column12);
                table.PrimaryKey = new[] { column1, column2 };

                this.PilotDataSet = new DataSet("Containers") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorPartitionContainersPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector partition containers print table.
        /// </summary>
        /// <returns>The connector partition containers print table</returns>
        protected DataTable GetConnectorPartitionContainersPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Container
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Include / Exclude
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Sort Column1
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column3
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 4 }, { "Hidden", true }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column4
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 5 }, { "Hidden", true }, { "SortOrder", 3 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column5
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 6 }, { "Hidden", true }, { "SortOrder", 4 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column6
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 7 }, { "Hidden", true }, { "SortOrder", 5 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column7
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 8 }, { "Hidden", true }, { "SortOrder", 6 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column8
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 9 }, { "Hidden", true }, { "SortOrder", 7 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column9
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 10 }, { "Hidden", true }, { "SortOrder", 8 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column10
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 11 }, { "Hidden", true }, { "SortOrder", 9 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector partition containers data set.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorPartitionContainersDataSet(string partitionName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var partition = connector.XPathSelectElement("ma-partition-data/partition[selected = 1 and name = '" + partitionName + "']");

                    if (partition != null)
                    {
                        var table = dataSet.Tables[0];

                        var inclusions = partition.XPathSelectElements("filter/containers/inclusions/inclusion");

                        var columnCount = table.Columns.Count;
                        foreach (var inclusion in inclusions)
                        {
                            var distinguishedName = (string)inclusion;
                            if (!string.IsNullOrEmpty(distinguishedName))
                            {
                                var row = this.GetContainerSelectionRow(distinguishedName, true, columnCount);
                                Documenter.AddRow(table, row);
                            }
                        }

                        var exclusions = partition.XPathSelectElements("filter/containers/exclusions/exclusion");

                        foreach (var exclusion in exclusions)
                        {
                            var distinguishedName = (string)exclusion;
                            if (!string.IsNullOrEmpty(distinguishedName))
                            {
                                var row = this.GetContainerSelectionRow(distinguishedName, false, columnCount);
                                Documenter.AddRow(table, row);
                            }
                        }

                        table.AcceptChanges();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);
            }
        }

        /// <summary>
        /// Gets the container selection row.
        /// </summary>
        /// <param name="distinguishedName">The distinguished name of the container.</param>
        /// <param name="inclusion">True if the container is included.</param>
        /// <param name="columnCount">The column count.</param>
        /// <returns>The container selection row</returns>
        protected object[] GetContainerSelectionRow(string distinguishedName, bool inclusion, int columnCount)
        {
            Logger.Instance.WriteMethodEntry("Container: '{0}'. Included: '{1}'.", distinguishedName, inclusion);

            try
            {
                var row = new object[columnCount];
                var distinguishedNameParts = distinguishedName.Split(new string[] { "OU=" }, StringSplitOptions.None);
                var partsCount = distinguishedNameParts.Length;
                row[0] = distinguishedName;
                row[1] = inclusion ? "Include" : "Exclude";

                if (partsCount > Documenter.MaxSortableColumns)
                {
                    Logger.Instance.WriteInfo("Container: '{0}' is deeper than '{1}' levels. Display sequence may be a little out-of-order.", distinguishedName, Documenter.MaxSortableColumns);
                }

                for (var i = 0; i < row.Length - 2 && i < Documenter.MaxSortableColumns; ++i)
                {
                    row[2 + i] = string.Empty;
                    if (i < partsCount)
                    {
                        if (partsCount == 1)
                        {
                            row[2 + i] = " " + distinguishedNameParts[0]; // so that the domain root is always sorted first
                        }
                        else
                        {
                            row[2 + i] = distinguishedNameParts[partsCount - 1 - i];
                        }
                    }
                }

                return row;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Container: '{0}'. Included: '{1}'.", distinguishedName, inclusion);
            }
        }

        /// <summary>
        /// Creates the connector partition containers diffgram.
        /// </summary>
        protected void CreateConnectorPartitionContainersDiffgram()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = Documenter.GetDiffgram(this.PilotDataSet, this.ProductionDataSet);
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector partition containers.
        /// </summary>
        protected void PrintConnectorPartitionContainers()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Container", 70 }, { "Include / Exclude", 30 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Container Selections

        #region Selected Object Types

        /// <summary>
        /// Processes the selected object types.
        /// </summary>
        protected void ProcessConnectorSelectedObjectTypes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Selected Object Types.");

                this.CreateSimpleSettingsDataSets(1);

                this.FillConnectorSelectedObjectTypesDataSet(true);
                this.FillConnectorSelectedObjectTypesDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorSelectedObjectTypes();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the selected object types data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorSelectedObjectTypesDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var objectTypes = config.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']/ma-partition-data/partition[position() = 1]/filter/object-classes/object-class");

                foreach (var objectType in objectTypes)
                {
                    Documenter.AddRow(table, new object[] { (string)objectType });
                }

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the selected object types.
        /// </summary>
        protected void PrintConnectorSelectedObjectTypes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Selected Object Types";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Object Types", 100 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Selected Object Types

        #region Selected Attributes

        /// <summary>
        /// Processes the connector selected attributes.
        /// </summary>
        protected void ProcessConnectorSelectedAttributes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Selected Attributes. This may take a few minutes...");

                this.CreateSimpleSettingsDataSets(4);

                this.FillConnectorSelectedAttributesDataSet(true);
                this.FillConnectorSelectedAttributesDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorSelectedAttributes();

                this.ProcessConnectorAttributeFlowsSummary();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector selected attributes data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorSelectedAttributesDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];
                    var currentConnectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant(); // This may be pilot or production GUID

                    foreach (var attribute in connector.XPathSelectElements("attribute-inclusion/attribute"))
                    {
                        var attributeName = (string)attribute;

                        var attributeInfo = connector.XPathSelectElement(".//dsml:attribute-type[dsml:name = '" + attributeName + "']", Documenter.NamespaceManager);
                        if (attributeInfo != null)
                        {
                            var hasInboundFlows = config.XPathSelectElements(Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + currentConnectorGuid + "' and direction = 'Inbound' " + Documenter.SyncRuleDisabledCondition + "]/attribute-mappings/mapping/src[attr = '" + attributeName + "']").Count() != 0;
                            var hasOutboundFlows = config.XPathSelectElements(Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + currentConnectorGuid + "' and direction = 'Outbound' " + Documenter.SyncRuleDisabledCondition + "]/attribute-mappings/mapping[dest = '" + attributeName + "']").Count() != 0;

                            var row = table.NewRow();

                            row[0] = attributeName;
                            var attributeSyntax = (string)attributeInfo.Element(Documenter.DsmlNamespace + "syntax");
                            row[1] = Documenter.GetAttributeType(attributeSyntax, (string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexable"));
                            row[2] = ((string)attributeInfo.Attribute("single-value") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";

                            row[3] = hasInboundFlows && hasOutboundFlows ? "Import / Export" : hasInboundFlows ? "Import" : hasOutboundFlows ? "Export" : "No";

                            Documenter.AddRow(table, row);
                        }
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the connector selected attributes.
        /// </summary>
        protected void PrintConnectorSelectedAttributes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Selected Attributes";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Attribute Name", 40 }, { "Type", 35 }, { "Multi-valued", 10 }, { "Flows Configured?", 15 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Selected Atrributes

        #region Attribute Flows Summary

        /// <summary>
        /// Processes the connector import and export attribute flows.
        /// </summary>
        protected void ProcessConnectorAttributeFlowsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter.WriteBeginTag("div");
                this.ReportWriter.WriteAttribute("class", "EndToEndFlowsSummary");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportToCWriter.WriteBeginTag("div");
                this.ReportToCWriter.WriteAttribute("class", "EndToEndFlowsSummary");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);

                var sectionTitle = "End-to-End Attribute Flows Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3);

                var configSetting = ConfigurationManager.AppSettings["SuppressConnectorImportExportEndToEndFlows"];
                if (!string.IsNullOrEmpty(configSetting) && configSetting.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Instance.WriteWarning(string.Format(CultureInfo.InvariantCulture, "!!WARNING!! SuppressConnectorImportExportEndToEndFlows = '{0}'.", configSetting));
                    this.WriteContentParagraph("The documentation of this section is suppressed via config setting.", "Highlight");
                    return;
                }

                var objectTypeXPath = "/ma-data[name ='" + this.ConnectorName + "']/ma-partition-data/partition[position() = 1]/filter/object-classes/object-class";
                var pilotObjectTypes = from objectClass in this.PilotXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(true) + objectTypeXPath)
                                       let objectType = (string)objectClass
                                       orderby objectType
                                       select objectType;

                foreach (var objectType in pilotObjectTypes)
                {
                    this.currentDataSourceObjectType = objectType;
                    this.ProcessConnectorObjectAttributeFlowsSummary();
                }

                var productionObjectTypes = from objectClass in this.ProductionXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(false) + objectTypeXPath)
                                            let objectType = (string)objectClass
                                            orderby objectType
                                            select objectType;

                productionObjectTypes = productionObjectTypes.Where(productionObjectType => !pilotObjectTypes.Contains(productionObjectType));

                foreach (var objectType in productionObjectTypes)
                {
                    this.currentDataSourceObjectType = objectType;
                    this.ProcessConnectorObjectAttributeFlowsSummary();
                }
            }
            finally
            {
                this.ReportWriter.WriteEndTag("div");
                this.ReportToCWriter.WriteEndTag("div");

                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes connector object attribute flows summary
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "Reviewed. Ignored.")]
        protected void ProcessConnectorObjectAttributeFlowsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ProcessConnectorObjectImportAttributeFlowsSummary();
                this.ProcessConnectorObjectExportAttributeFlowsSummary();

                if (this.DiffgramDataSets[0].Tables[0].Rows.Count != 0 || this.DiffgramDataSets[1].Tables[0].Rows.Count != 0)
                {
                    this.WriteSectionHeader(this.currentDataSourceObjectType, 4);
                }

                if (this.DiffgramDataSets[0].Tables[0].Rows.Count != 0)
                {
                    this.WriteSectionHeader("Import Flows", 5, "Import Flows", this.ConnectorGuid + this.currentDataSourceObjectType);

                    this.DiffgramDataSet = this.DiffgramDataSets[0];
                    this.PrintCreateConnectorObjectImportAttributeFlowsSummary();
                }

                if (this.DiffgramDataSets[1].Tables[0].Rows.Count != 0)
                {
                    this.WriteSectionHeader("Export Flows", 5, "Export Flows", this.ConnectorGuid + this.currentDataSourceObjectType);

                    this.DiffgramDataSet = this.DiffgramDataSets[1];
                    this.PrintCreateConnectorObjectExportAttributeFlowsSummary();
                }
            }
            catch (Exception e)
            {
                Logger.Instance.WriteError("Error while processing end-to-end attribute flows. Current Object Type: {0}. Error: {1}", this.currentDataSourceObjectType, e.ToString());
            }
            finally
            {
                this.ResetDiffgram();
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Import Attribute Flows Summary

        /// <summary>
        /// Processes connector object import attribute flows summary
        /// </summary>
        protected void ProcessConnectorObjectImportAttributeFlowsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing End to End Import Attribute Flows Summary. This may take a few minutes...");

                this.CreateConnectorObjectImportAttributeFlowsSummaryDataSets();

                this.FillConnectorObjectImportAttributeFlowsSummary(true);
                this.FillConnectorObjectImportAttributeFlowsSummary(false);

                this.CreateConnectorObjectImportAttributeFlowsSummaryDiffgram();

                ////this.PrintCreateConnectorObjectImportAttributeFlowsSummary();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates connector object import attribute flows summary
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorObjectImportAttributeFlowsSummaryDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("MetaverseAttributes") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("MetaverseAttribute");
                var column2 = new DataColumn("FlowDirection");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("InboundSyncRules") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("MetaverseAttribute");
                var column22 = new DataColumn("Source");
                var column32 = new DataColumn("InboundSyncRule");
                var column42 = new DataColumn("InboundSyncRulePrecedence", typeof(int));
                var column52 = new DataColumn("InboundSyncRuleScopingCondition");
                var column62 = new DataColumn("InboundSyncRuleGuid");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.Columns.Add(column52);
                table2.Columns.Add(column62);
                table2.PrimaryKey = new[] { column12, column32 };

                var table3 = new DataTable("MetaverseAttributes2") { Locale = CultureInfo.InvariantCulture };

                var column13 = new DataColumn("MetaverseAttribute"); // needed for cascading - see table 4 comments
                var column23 = new DataColumn("InboundSyncRule");
                var column33 = new DataColumn("FlowDirection");

                table3.Columns.Add(column13);
                table3.Columns.Add(column23);
                table3.Columns.Add(column33);
                table3.PrimaryKey = new[] { column13 };

                var table4 = new DataTable("MetaverseObjectTypeOutboundFlows") { Locale = CultureInfo.InvariantCulture };

                var column14 = new DataColumn("MetaverseAttribute");
                var column24 = new DataColumn("InboundSyncRule");
                var column34 = new DataColumn("Source");
                var column44 = new DataColumn("TargetAttribute");
                var column54 = new DataColumn("OutboundSyncRulePrecedence", typeof(int));
                var column64 = new DataColumn("OutboundSyncRule");
                var column74 = new DataColumn("TargetConnector");
                var column84 = new DataColumn("OutboundSyncRuleSyncRuleScopingCondition");
                var column94 = new DataColumn("TargetConnectorGuid");
                var column104 = new DataColumn("OutboundSyncRuleSyncRuleGuid");

                table4.Columns.Add(column14);
                table4.Columns.Add(column24);
                table4.Columns.Add(column34);
                table4.Columns.Add(column44);
                table4.Columns.Add(column54);
                table4.Columns.Add(column64);
                table4.Columns.Add(column74);
                table4.Columns.Add(column84);
                table4.Columns.Add(column94);
                table4.Columns.Add(column104);
                table4.PrimaryKey = new[] { column14, column44, column64, column74 }; // column24 is excluded as we'll insert only one row in the the parent table.

                this.PilotDataSet = new DataSet("ImportAttributeFlowsSummary") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);
                this.PilotDataSet.Tables.Add(table3);
                this.PilotDataSet.Tables.Add(table4);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);
                var dataRelation23 = new DataRelation("DataRelation23", new[] { column12, column32 }, new[] { column13, column23 }, false);
                var dataRelation34 = new DataRelation("DataRelation34", new[] { column13, column23 }, new[] { column14, column24 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);
                this.PilotDataSet.Relations.Add(dataRelation23);
                this.PilotDataSet.Relations.Add(dataRelation34);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorObjectImportAttributeFlowsSummaryPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector object import attribute flows summary print table.
        /// </summary>
        /// <returns>The connector object import attribute flows summary print table.</returns>
        protected DataTable GetConnectorObjectImportAttributeFlowsSummaryPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Metaverse Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Flow Direction
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 2
                // Source
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Sync Rule
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 5 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Precedence
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Sync Rule Scoping Condition
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Sync Rule Guid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 5 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 3
                // Metaverse Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule Name
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Flow Direction
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 4
                // Inbound Sync Rule Name
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Source
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Target Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Precedence
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 4 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 5 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 9 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Target Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 6 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 8 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Scoping Condition
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 7 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Target Connector Guid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 8 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Guid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 9 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector object import attribute flows summary data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorObjectImportAttributeFlowsSummary(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];
                var table2 = dataSet.Tables[1];
                var table3 = dataSet.Tables[2];
                var table4 = dataSet.Tables[3];

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var connectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant();
                    var metaverseAttributesXPath = Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorGuid + "' and direction = 'Inbound' " + Documenter.SyncRuleDisabledCondition + " and sourceObjectType = '" + this.currentDataSourceObjectType + "']/attribute-mappings/mapping/dest";
                    var metaverseAttributes = from inboundSyncRuleDestination in config.XPathSelectElements(metaverseAttributesXPath)
                                              let dest = (string)inboundSyncRuleDestination
                                              orderby dest
                                              select dest;
                    foreach (var metaverseAttribute in metaverseAttributes.Distinct())
                    {
                        Documenter.AddRow(table, new object[] { metaverseAttribute, "&#8592;" });

                        var inboundSyncRuleXPath = Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorGuid + "' and direction = 'Inbound' " + Documenter.SyncRuleDisabledCondition + " and sourceObjectType = '" + this.currentDataSourceObjectType + "' and ./attribute-mappings/mapping/dest = '" + metaverseAttribute + "']";
                        var inboundSyncRules = config.XPathSelectElements(inboundSyncRuleXPath);
                        inboundSyncRules = from syncRule in inboundSyncRules
                                           let inboundSyncRulePrecedence = (int)syncRule.Element("precedence")
                                           orderby inboundSyncRulePrecedence
                                           select syncRule;

                        var inboundSyncRuleRank = 0; // Used only for sorting, not displayed on the report
                        foreach (var inboundSyncRule in inboundSyncRules)
                        {
                            var inboundSyncRuleName = (string)inboundSyncRule.Element("name");
                            var inboundSyncRuleGuid = (string)inboundSyncRule.Element("id");
                            var targetObjectType = (string)inboundSyncRule.Element("targetObjectType");
                            ++inboundSyncRuleRank;

                            var transformation = inboundSyncRule.XPathSelectElement("./attribute-mappings/mapping[dest = '" + metaverseAttribute + "']");
                            var expression = (string)transformation.Element("expression");
                            var srcAttribute = (string)transformation.XPathSelectElement("src/attr");
                            var src = (string)transformation.XPathSelectElement("src");
                            var source = !string.IsNullOrEmpty(expression) ? expression : !string.IsNullOrEmpty(srcAttribute) ? srcAttribute : src;

                            var inboundSyncRuleScopingConditions = inboundSyncRule.XPathSelectElements("./synchronizationCriteria/conditions");
                            var inboundSyncRuleScopingConditionString = string.Empty;
                            var inboundSyncRuleScopingConditionsCount = inboundSyncRuleScopingConditions.Count();
                            if (inboundSyncRuleScopingConditionsCount != 0)
                            {
                                var conditionIndex = -1;
                                foreach (var condition in inboundSyncRuleScopingConditions)
                                {
                                    ++conditionIndex;
                                    var scopes = condition.Elements("scope");
                                    inboundSyncRuleScopingConditionString += scopes.Select(scope => string.Format(CultureInfo.InvariantCulture, "{0} {1} {2}", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"))).Aggregate((i, j) => i + " AND " + j);
                                    if (conditionIndex < inboundSyncRuleScopingConditionsCount - 1)
                                    {
                                        inboundSyncRuleScopingConditionString += " OR ";
                                    }
                                }
                            }

                            Documenter.AddRow(table2, new object[] { metaverseAttribute, source, inboundSyncRuleName, inboundSyncRuleRank, inboundSyncRuleScopingConditionString, inboundSyncRuleGuid });

                            if (table3.Select("MetaverseAttribute = '" + metaverseAttribute + "'").Count() == 0)
                            {
                                Documenter.AddRow(table3, new object[] { metaverseAttribute, inboundSyncRuleName, "&#8594;" });

                                // Outbound metaverse flows
                                var outboundSyncRules = config.XPathSelectElements(Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[sourceObjectType = '" + targetObjectType + "' and direction = 'Outbound' " + Documenter.SyncRuleDisabledCondition + " and (attribute-mappings/mapping/src/attr = '" + metaverseAttribute + "' or contains(attribute-mappings/mapping/expression, '[" + metaverseAttribute + "]'))]");
                                outboundSyncRules = from syncRule in outboundSyncRules
                                                    let outboundSyncRulePrecedence = (int)syncRule.Element("precedence")
                                                    orderby outboundSyncRulePrecedence
                                                    select syncRule;
                                var outboundSyncRuleRank = 0;
                                foreach (var outboundSyncRule in outboundSyncRules)
                                {
                                    var outboundSyncRuleName = (string)outboundSyncRule.Element("name");
                                    var outboundSyncRuleGuid = (string)outboundSyncRule.Element("id");
                                    var outboundConnectorGuid = ((string)outboundSyncRule.Element("connector") ?? string.Empty).ToUpperInvariant();
                                    var outboundConnectorName = (string)config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + outboundConnectorGuid + "']/name");
                                    if (string.IsNullOrEmpty(outboundConnectorName))
                                    {
                                        Logger.Instance.WriteWarning(string.Format(CultureInfo.InvariantCulture, "Unable to dereferece connector: '{0}'. The documentation of outbound flows will be skipped for this connector. PilotConfig: '{1}'.", outboundConnectorGuid, pilotConfig));
                                        continue;
                                    }

                                    ++outboundSyncRuleRank; // Used only for sorting, not displayed on the report

                                    var targetAttributeMapping = outboundSyncRule.XPathSelectElement("attribute-mappings/mapping[./src/attr = '" + metaverseAttribute + "' or contains(expression, '[" + metaverseAttribute + "]')]");
                                    var mappingExpression = (string)targetAttributeMapping.XPathSelectElement("./expression");
                                    var mappingSourceAttribute = (string)targetAttributeMapping.XPathSelectElement("./src/attr");
                                    var mappingSource = (string)targetAttributeMapping.XPathSelectElement("./src");
                                    var targetAttribute = (string)targetAttributeMapping.XPathSelectElement("./dest");
                                    var metaverseSource = (string)mappingExpression ?? (string)mappingSourceAttribute ?? (string)mappingSource ?? string.Empty;
                                    if (!string.IsNullOrEmpty(metaverseSource))
                                    {
                                        var outboundSyncRuleScopingConditions = outboundSyncRule.XPathSelectElements("./synchronizationCriteria/conditions");
                                        var outboundSyncRuleScopingCondition = string.Empty;
                                        var outboundSyncRuleScopingConditionCount = outboundSyncRuleScopingConditions.Count();
                                        if (outboundSyncRuleScopingConditionCount != 0)
                                        {
                                            var conditionIndex = -1;
                                            foreach (var condition in outboundSyncRuleScopingConditions)
                                            {
                                                ++conditionIndex;
                                                var scopes = condition.Elements("scope");
                                                outboundSyncRuleScopingCondition += scopes.Select(scope => string.Format(CultureInfo.InvariantCulture, "{0} {1} {2}", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"))).Aggregate((i, j) => i + " AND " + j);
                                                if (conditionIndex < outboundSyncRuleScopingConditionCount - 1)
                                                {
                                                    outboundSyncRuleScopingCondition += " OR ";
                                                }
                                            }
                                        }

                                        Documenter.AddRow(table4, new object[] { metaverseAttribute, inboundSyncRuleName, metaverseSource, targetAttribute, outboundSyncRuleRank, outboundSyncRuleName, outboundConnectorName, outboundSyncRuleScopingCondition, outboundConnectorGuid, outboundSyncRuleGuid });
                                    }
                                }
                            }
                        }
                    }
                }

                table.AcceptChanges();
                table2.AcceptChanges();
                table3.AcceptChanges();
                table4.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the connector object import attribute flows summary difference gram.
        /// </summary>
        protected void CreateConnectorObjectImportAttributeFlowsSummaryDiffgram()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = Documenter.GetDiffgram(this.PilotDataSet, this.ProductionDataSet);
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector object import attribute flows summary header table.
        /// </summary>
        /// <returns>The connector object import attribute flows summary header table.</returns>
        protected DataTable GetConnectorObjectImportAttributeFlowsSummaryHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Import Flows
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Import Flows" }, { "RowSpan", 1 }, { "ColSpan", 5 }, { "ColWidth", 0 } }).Values.Cast<object>().ToArray());

                // Export Flows
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Export Flows" }, { "RowSpan", 1 }, { "ColSpan", 7 }, { "ColWidth", 0 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Metaverse attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Metaverse Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 8 } }).Values.Cast<object>().ToArray());

                // Flow Direction
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "&#8592;" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 2 } }).Values.Cast<object>().ToArray());

                // Source
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Connector Source" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Inbound Sync Rule" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule Scoping Condition
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 4 }, { "ColumnName", "Inbound Sync Rule Scoping Condition" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Metaverse attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 5 }, { "ColumnName", "Metaverse Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 8 } }).Values.Cast<object>().ToArray());

                // Flow Direction
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 6 }, { "ColumnName", "&#8594;" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 2 } }).Values.Cast<object>().ToArray());

                // Source
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 7 }, { "ColumnName", "Source" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Target
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 8 }, { "ColumnName", "Target" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 9 }, { "ColumnName", "Outbound Sync Rule" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Target Connector
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 10 }, { "ColumnName", "Target Connector" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Scoping Condition
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 11 }, { "ColumnName", "Outbound Sync Rule Scoping Condition" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector object import attribute flows summary.
        /// </summary>
        protected void PrintCreateConnectorObjectImportAttributeFlowsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = this.GetConnectorObjectImportAttributeFlowsSummaryHeaderTable();
                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, HtmlTableSize.Huge);
                }
            }
            finally
            {
                ////this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Import Attribute Flows Summary

        #region Export Attribute Flows Summary

        /// <summary>
        /// Processes connector object export attribute flows summary
        /// </summary>
        protected void ProcessConnectorObjectExportAttributeFlowsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing End to End Export Attribute Flows Summary. This may take a few minutes...");

                this.CreateConnectorObjectExportAttributeFlowsSummaryDataSets();

                this.FillConnectorObjectExportAttributeFlowsSummary(true);
                this.FillConnectorObjectExportAttributeFlowsSummary(false);

                this.CreateConnectorObjectExportAttributeFlowsSummaryDiffgram();

                ////this.PrintCreateConnectorObjectExportAttributeFlowsSummary();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the connector object export attribute flows summary data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorObjectExportAttributeFlowsSummaryDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("DataSourceAttributes") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("DataSourceAttribute");
                var column2 = new DataColumn("FlowDirection");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("OutboundSyncRules") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("DataSourceAttribute"); // to be able to do cascading data relations
                var column22 = new DataColumn("Source");
                var column32 = new DataColumn("OutboundSyncRule");
                var column42 = new DataColumn("OutboundSyncRulePrecedence", typeof(int));
                var column52 = new DataColumn("OutboundSyncRuleScopingCondition");
                var column62 = new DataColumn("OutboundSyncRuleGuid");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.Columns.Add(column52);
                table2.Columns.Add(column62);
                table2.PrimaryKey = new[] { column12, column32 };

                var table3 = new DataTable("MetaverseAttributes") { Locale = CultureInfo.InvariantCulture };

                var column13 = new DataColumn("DataSourceAttribute");
                var column23 = new DataColumn("OutboundSyncRule");
                var column33 = new DataColumn("MetaverseAttribute");
                var column43 = new DataColumn("FlowDirection");

                table3.Columns.Add(column13);
                table3.Columns.Add(column23);
                table3.Columns.Add(column33);
                table3.Columns.Add(column43);
                table3.PrimaryKey = new[] { column13, column23, column33 };

                var table4 = new DataTable("MetaverseObjectTypeInboundFlows") { Locale = CultureInfo.InvariantCulture };

                var column14 = new DataColumn("DataSourceAttribute");
                var column24 = new DataColumn("OutboundSyncRule");
                var column34 = new DataColumn("MetaverseAttribute");
                var column44 = new DataColumn("Source");
                var column54 = new DataColumn("InboundSyncRulePrecedence", typeof(int));
                var column64 = new DataColumn("InboundSyncRule");
                var column74 = new DataColumn("SourceConnector");
                var column84 = new DataColumn("InboundSyncRuleScopingCondition");
                var column94 = new DataColumn("SourceConnectorGuid");
                var column104 = new DataColumn("InboundSyncRuleGuid");

                table4.Columns.Add(column14);
                table4.Columns.Add(column24);
                table4.Columns.Add(column34);
                table4.Columns.Add(column44);
                table4.Columns.Add(column54);
                table4.Columns.Add(column64);
                table4.Columns.Add(column74);
                table4.Columns.Add(column84);
                table4.Columns.Add(column94);
                table4.Columns.Add(column104);
                table4.PrimaryKey = new[] { column14, column34, column64, column74 }; // column24 is excluded as we'll insert only one row in the the parent table.

                this.PilotDataSet = new DataSet("ExportAttributeFlowsSummary") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);
                this.PilotDataSet.Tables.Add(table3);
                this.PilotDataSet.Tables.Add(table4);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);
                var dataRelation23 = new DataRelation("DataRelation23", new[] { column12, column32 }, new[] { column13, column23 }, false);
                var dataRelation34 = new DataRelation("DataRelation34", new[] { column13, column23, column33 }, new[] { column14, column24, column34 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);
                this.PilotDataSet.Relations.Add(dataRelation23);
                this.PilotDataSet.Relations.Add(dataRelation34);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorObjectExportAttributeFlowsSummaryPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector object export attribute flows summary print table.
        /// </summary>
        /// <returns>The connector object export attribute flows summary print table.</returns>
        protected DataTable GetConnectorObjectExportAttributeFlowsSummaryPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // DataSource Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Flow Direction
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 2
                // Source
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 5 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Precedence
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Scoping Condition
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Guid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 5 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 3
                // Outbound Sync Rule Name
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Metaverse Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Flow Direction
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 4
                // Outbound Sync Rule Name
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Source
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule Precedence
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 4 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 5 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 9 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Source Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 6 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 8 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule Scoping Condition
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 7 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Source Connector Guid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 8 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule Guid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 3 }, { "ColumnIndex", 9 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector object export attribute flows summary data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorObjectExportAttributeFlowsSummary(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];
                var table2 = dataSet.Tables[1];
                var table3 = dataSet.Tables[2];
                var table4 = dataSet.Tables[3];

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var connectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant();
                    var targetAttributesXPath = Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorGuid + "' and direction = 'Outbound' " + Documenter.SyncRuleDisabledCondition + " and targetObjectType = '" + this.currentDataSourceObjectType + "']/attribute-mappings/mapping/dest";
                    var targetAttributes = from outboundSyncRuleDestination in config.XPathSelectElements(targetAttributesXPath)
                                           let dest = (string)outboundSyncRuleDestination
                                           orderby dest
                                           select dest;

                    foreach (var targetAttribute in targetAttributes.Distinct())
                    {
                        Documenter.AddRow(table, new object[] { targetAttribute, "&#8592;" });

                        var outboundSyncRuleXPath = Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorGuid + "' and direction = 'Outbound' " + Documenter.SyncRuleDisabledCondition + " and targetObjectType = '" + this.currentDataSourceObjectType + "' and ./attribute-mappings/mapping/dest = '" + targetAttribute + "']";
                        var outboundSyncRules = config.XPathSelectElements(outboundSyncRuleXPath);
                        outboundSyncRules = from syncRule in outboundSyncRules
                                            let precedence = (int)syncRule.Element("precedence")
                                            orderby precedence
                                            select syncRule;

                        var outboundSyncRuleRank = 0; // Used only for sorting, not displayed on the report
                        foreach (var outboundSyncRule in outboundSyncRules)
                        {
                            var outboundSyncRuleName = (string)outboundSyncRule.Element("name");
                            var outboundSyncRuleGuid = (string)outboundSyncRule.Element("id");
                            var sourceObjectType = (string)outboundSyncRule.Element("sourceObjectType");
                            ++outboundSyncRuleRank;

                            var transformation = outboundSyncRule.XPathSelectElement("./attribute-mappings/mapping[dest = '" + targetAttribute + "']");
                            var expression = (string)transformation.Element("expression");
                            var srcAttribute = (string)transformation.XPathSelectElement("src/attr");
                            var src = (string)transformation.XPathSelectElement("src");
                            var source = !string.IsNullOrEmpty(expression) ? expression : !string.IsNullOrEmpty(srcAttribute) ? srcAttribute : src;

                            var outboundSyncRuleScopingConditions = outboundSyncRule.XPathSelectElements("./synchronizationCriteria/conditions");
                            var outboundSyncRuleScopingConditionString = string.Empty;
                            var outboundSyncRuleScopingConditionsCount = outboundSyncRuleScopingConditions.Count();
                            if (outboundSyncRuleScopingConditionsCount != 0)
                            {
                                var conditionIndex = -1;
                                foreach (var condition in outboundSyncRuleScopingConditions)
                                {
                                    ++conditionIndex;
                                    var scopes = condition.Elements("scope");
                                    outboundSyncRuleScopingConditionString += scopes.Select(scope => string.Format(CultureInfo.InvariantCulture, "{0} {1} {2}", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"))).Aggregate((i, j) => i + " AND " + j);
                                    if (conditionIndex < outboundSyncRuleScopingConditionsCount - 1)
                                    {
                                        outboundSyncRuleScopingConditionString += " OR ";
                                    }
                                }
                            }

                            Documenter.AddRow(table2, new object[] { targetAttribute, source, outboundSyncRuleName, outboundSyncRuleRank, outboundSyncRuleScopingConditionString, outboundSyncRuleGuid });

                            var metaverseAttributes = new List<string>();
                            if (!string.IsNullOrEmpty(expression))
                            {
                                foreach (Match match in Regex.Matches(expression, @"\[(.*?)\]"))
                                {
                                    metaverseAttributes.Add(match.Value.TrimStart('[').TrimEnd(']'));
                                }
                            }
                            else if (!string.IsNullOrEmpty(srcAttribute))
                            {
                                metaverseAttributes.Add(srcAttribute);
                            }

                            foreach (var metaverseAttribute in metaverseAttributes)
                            {
                                if (table3.Select("DataSourceAttribute = '" + targetAttribute + "' and MetaverseAttribute = '" + metaverseAttribute + "'").Count() == 0)
                                {
                                    Documenter.AddRow(table3, new object[] { targetAttribute, outboundSyncRuleName, metaverseAttribute, "&#8592;" });

                                    // Inbound metaverse flows
                                    var inboundSyncRules = config.XPathSelectElements(Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[targetObjectType = '" + sourceObjectType + "' and direction = 'Inbound' " + Documenter.SyncRuleDisabledCondition + " and (attribute-mappings/mapping/dest = '" + metaverseAttribute + "' or contains(attribute-mappings/mapping/expression, '[" + metaverseAttribute + "]'))]");
                                    inboundSyncRules = from syncRule in inboundSyncRules
                                                       let inboundSyncRulePrecedence = (int)syncRule.Element("precedence")
                                                       orderby inboundSyncRulePrecedence
                                                       select syncRule;

                                    var inboundSyncRuleRank = 0;
                                    foreach (var inboundSyncRule in inboundSyncRules)
                                    {
                                        var inboundSyncRuleName = (string)inboundSyncRule.Element("name");
                                        var inboundSyncRuleGuid = (string)inboundSyncRule.Element("id");
                                        var inboundConnectorGuid = ((string)inboundSyncRule.Element("connector") ?? string.Empty).ToUpperInvariant();
                                        var inboundConnectorName = (string)config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + inboundConnectorGuid + "']/name");
                                        if (string.IsNullOrEmpty(inboundConnectorName))
                                        {
                                            Logger.Instance.WriteWarning(string.Format(CultureInfo.InvariantCulture, "Unable to dereferece connector: '{0}'. The documentation of inbound flows will be skipped for this connector. PilotConfig: '{1}'.", inboundConnectorGuid, pilotConfig));
                                            continue;
                                        }

                                        ++inboundSyncRuleRank; // Used only for sorting, not displayed on the report

                                        var metaverseAttributeMapping = inboundSyncRule.XPathSelectElement("attribute-mappings/mapping[dest = '" + metaverseAttribute + "' or contains(expression, '[" + metaverseAttribute + "]')]");
                                        var mappingExpression = (string)metaverseAttributeMapping.XPathSelectElement("./expression");
                                        var mappingSourceAttribute = (string)metaverseAttributeMapping.XPathSelectElement("./src/attr");
                                        var mappingSource = (string)metaverseAttributeMapping.XPathSelectElement("./src");
                                        var metaverseSource = (string)mappingExpression ?? (string)mappingSourceAttribute ?? (string)mappingSource ?? string.Empty;
                                        if (!string.IsNullOrEmpty(metaverseSource))
                                        {
                                            var inboundSyncRuleScopingConditions = inboundSyncRule.XPathSelectElements("./synchronizationCriteria/conditions");
                                            var inboundSyncRuleScopingCondition = string.Empty;
                                            var inboundSyncRuleScopingConditionCount = inboundSyncRuleScopingConditions.Count();
                                            if (inboundSyncRuleScopingConditionCount != 0)
                                            {
                                                var conditionIndex = -1;
                                                foreach (var condition in inboundSyncRuleScopingConditions)
                                                {
                                                    ++conditionIndex;
                                                    var scopes = condition.Elements("scope");
                                                    inboundSyncRuleScopingCondition += scopes.Select(scope => string.Format(CultureInfo.InvariantCulture, "{0} {1} {2}", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"))).Aggregate((i, j) => i + " AND " + j);
                                                    if (conditionIndex < inboundSyncRuleScopingConditionCount - 1)
                                                    {
                                                        inboundSyncRuleScopingCondition += " OR ";
                                                    }
                                                }
                                            }

                                            Documenter.AddRow(table4, new object[] { targetAttribute, outboundSyncRuleName, metaverseAttribute, metaverseSource, inboundSyncRuleRank, inboundSyncRuleName, inboundConnectorName, inboundSyncRuleScopingCondition, inboundConnectorGuid, inboundSyncRuleGuid });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                table.AcceptChanges();
                table2.AcceptChanges();
                table3.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the connector object export attribute flows summary difference gram.
        /// </summary>
        protected void CreateConnectorObjectExportAttributeFlowsSummaryDiffgram()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = Documenter.GetDiffgram(this.PilotDataSet, this.ProductionDataSet);
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector object export attribute flows summary header table.
        /// </summary>
        /// <returns>The connector object export attribute flows summary header table.</returns>
        protected DataTable GetConnectorObjectExportAttributeFlowsSummaryHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Export Flows
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Export Flows" }, { "RowSpan", 1 }, { "ColSpan", 5 }, { "ColWidth", 0 } }).Values.Cast<object>().ToArray());

                // Import Flows
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Import Flows" }, { "RowSpan", 1 }, { "ColSpan", 6 }, { "ColWidth", 0 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // DataSource Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "CS Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Flow Direction
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "&#8592;" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 2 } }).Values.Cast<object>().ToArray());

                // Source
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Source" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Outbound Sync Rule" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Outbound Sync Rule Scoping Condition
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 4 }, { "ColumnName", "Outbound Sync Rule Scoping Condition" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 13 } }).Values.Cast<object>().ToArray());

                // Metaverse attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 5 }, { "ColumnName", "Metaverse Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Flow Direction
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 6 }, { "ColumnName", "&#8592;" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 2 } }).Values.Cast<object>().ToArray());

                // Source
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 7 }, { "ColumnName", "Source" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 8 }, { "ColumnName", "Inbound Sync Rule" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Source Connector
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 9 }, { "ColumnName", "Source Connector" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule Scoping Condition
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 10 }, { "ColumnName", "Inbound Sync Rule Scoping Condition" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 15 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector object export attribute flows summary.
        /// </summary>
        protected void PrintCreateConnectorObjectExportAttributeFlowsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = this.GetConnectorObjectExportAttributeFlowsSummaryHeaderTable();
                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, HtmlTableSize.Huge);
                }
            }
            finally
            {
                ////this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Export Attribute Flows Summary

        #endregion Attribute Flows Summary

        #region Provisioning Sync Rules

        /// <summary>
        /// Processes the connector provisioning synchronization rules.
        /// </summary>
        protected void ProcessConnectorProvisioningSyncRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var reportType = SyncRuleDocumenter.SyncRuleReportType.ProvisioningSection;
                this.WriteConnectorSyncRulesHeader(reportType);

                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Inbound, reportType);
                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Outbound, reportType);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Provisioning Sync Rules

        #region Sticky Join Sync Rules

        /// <summary>
        /// Processes the connector sticky join synchronization rules.
        /// </summary>
        protected void ProcessConnectorStickyJoinSyncRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var reportType = SyncRuleDocumenter.SyncRuleReportType.StickyJoinSection;
                this.WriteConnectorSyncRulesHeader(reportType);

                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Inbound, reportType);
                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Outbound, reportType);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sticky Join Rules

        #region Normal Join Sync Rules

        /// <summary>
        /// Processes the connector normal join sync rules.
        /// </summary>
        protected void ProcessConnectorNormalJoinSyncRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var reportType = SyncRuleDocumenter.SyncRuleReportType.ConditionalJoinSection;
                this.WriteConnectorSyncRulesHeader(reportType);

                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Inbound, reportType);
                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Outbound, reportType);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Normal Join Rules

        #region Sync Rules

        /// <summary>
        /// Processes the connector synchronization rules.
        /// </summary>
        protected void ProcessConnectorSyncRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var reportType = SyncRuleDocumenter.SyncRuleReportType.AllSections;
                this.WriteConnectorSyncRulesHeader(reportType);

                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Inbound, reportType);
                this.ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection.Outbound, reportType);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the connector synchronization rules.
        /// </summary>
        /// <param name="direction">The sync rule direction.</param>
        /// <param name="reportType">Type of the report.</param>
        protected void ProcessConnectorSyncRules(SyncRuleDocumenter.SyncRuleDirection direction, SyncRuleDocumenter.SyncRuleReportType reportType)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Direction: '{0}'. Sync Rule Report Type: '{1}'.", direction, reportType);

            try
            {
                var sectionTitle = ConnectorDocumenter.GetSyncRuleSectionTitle(reportType);

                this.WriteSectionHeader(direction.ToString(), 4, direction.ToString() + sectionTitle, this.ConnectorGuid);

                var sectionPrinted = false;
                var pilotConnector = this.PilotXml.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(true) + "/ma-data[name ='" + this.ConnectorName + "']");
                var productionConnector = this.ProductionXml.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(false) + "/ma-data[name ='" + this.ConnectorName + "']");

                var pilotConnectorGuid = pilotConnector != null ? ((string)pilotConnector.Element("id") ?? string.Empty).ToUpperInvariant() : string.Empty;
                var productionConnectorGuid = productionConnector != null ? ((string)productionConnector.Element("id") ?? string.Empty).ToUpperInvariant() : string.Empty;

                var pilotSyncRules = this.PilotXml.XPathSelectElements(ConnectorDocumenter.GetSyncRuleXPath(pilotConnectorGuid, direction, reportType, true));
                var productionSyncRules = this.ProductionXml.XPathSelectElements(ConnectorDocumenter.GetSyncRuleXPath(productionConnectorGuid, direction, reportType, false));

                // Sort by name
                pilotSyncRules = from syncRule in pilotSyncRules
                                 let name = (string)syncRule.Element("name")
                                 orderby name
                                 select syncRule;

                // Sort by name
                productionSyncRules = from syncRule in productionSyncRules
                                      let name = (string)syncRule.Element("name")
                                      orderby name
                                      select syncRule;

                sectionPrinted = pilotSyncRules.Count() != 0;

                foreach (var syncRule in pilotSyncRules)
                {
                    var syncRuleName = (string)syncRule.Element("name");
                    var syncRuleGuid = (string)syncRule.Element("id");
                    var configEnvironment = productionSyncRules.Any(productionSyncRule => (string)productionSyncRule.Element("name") == (string)syncRule.Element("name")) ? ConfigEnvironment.PilotAndProduction : ConfigEnvironment.PilotOnly;
                    var connectorDocumenter = new SyncRuleDocumenter(this.PilotXml, this.ProductionXml, syncRuleName, syncRuleGuid, this.ConnectorName, configEnvironment);
                    var report = connectorDocumenter.GetReport(reportType);
                    this.ReportWriter.Write(report.Item1);
                    this.ReportToCWriter.Write(report.Item2);
                }

                var pilotSyncRulesNames = from syncRule in pilotSyncRules
                                          select (string)syncRule.Element("name");

                productionSyncRules = productionSyncRules.Where(productionSyncRule => !pilotSyncRulesNames.Contains((string)productionSyncRule.Element("name")));

                sectionPrinted = sectionPrinted || productionSyncRules.Count() != 0;

                foreach (var syncRule in productionSyncRules)
                {
                    var syncRuleName = (string)syncRule.Element("name");
                    var syncRuleGuid = (string)syncRule.Element("id");
                    var connectorDocumenter = new SyncRuleDocumenter(this.PilotXml, this.ProductionXml, syncRuleName, syncRuleGuid, this.ConnectorName, ConfigEnvironment.ProductionOnly);
                    var report = connectorDocumenter.GetReport(reportType);
                    this.ReportWriter.Write(report.Item1);
                    this.ReportToCWriter.Write(report.Item2);
                }

                if (!sectionPrinted)
                {
                    this.WriteContentParagraph("There are no <b>" + direction.ToString() + " " + sectionTitle.Replace(" Summary", string.Empty) + "</b> configured.", Documenter.CanHide);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Sync Rule Direction: '{0}'. Sync Rule Report Type: '{1}'.", direction, reportType);
            }
        }

        #endregion Sync Rules

        #region Run Profiles

        /// <summary>
        /// Processes the connector run profiles.
        /// </summary>
        protected void ProcessConnectorRunProfiles()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Run Profiles.");

                var sectionTitle = "Run Profiles";

                this.WriteSectionHeader(sectionTitle, 3);

                var xpath = "/ma-data[name ='" + this.ConnectorName + "']/ma-run-data/run-configuration";

                var pilot = this.PilotXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(true) + xpath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(false) + xpath, Documenter.NamespaceManager);

                var connectorHasRunProfilesConfigured = false;

                // Sort by name
                var pilotRunProfiles = from runProfile in pilot
                                       let name = (string)runProfile.Element("name")
                                       orderby name
                                       select name;

                foreach (var runProfile in pilotRunProfiles)
                {
                    connectorHasRunProfilesConfigured = true;
                    this.ProcessConnectorRunProfile(runProfile);
                }

                // Sort by name
                var productionRunProfiles = from runProfile in production
                                            let name = (string)runProfile.Element("name")
                                            orderby name
                                            select name;

                productionRunProfiles = productionRunProfiles.Where(productionRunProfile => !pilotRunProfiles.Contains(productionRunProfile));

                foreach (var runProfile in productionRunProfiles)
                {
                    connectorHasRunProfilesConfigured = true;
                    this.ProcessConnectorRunProfile(runProfile);
                }

                if (!connectorHasRunProfilesConfigured)
                {
                    this.WriteContentParagraph("There are no run profiles configured.");
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the connector run profile.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        protected void ProcessConnectorRunProfile(string runProfileName)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'.", runProfileName);

            try
            {
                this.WriteSectionHeader("Run Profile: " + runProfileName, 4, runProfileName);

                this.CreateConnectorRunProfileDataSets();

                this.FillConnectorRunProfileDataSet(runProfileName, true);
                this.FillConnectorRunProfileDataSet(runProfileName, false);

                this.CreateConnectorRunProfileDiffgram();

                this.PrintConnectorRunProfile();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Run Profile Name: '{0}'.", runProfileName);
            }
        }

        /// <summary>
        /// Creates the connector run profile data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorRunProfileDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("ConnectorRunProfiles") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Step Number", typeof(int));
                var column2 = new DataColumn("Step Name");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("ConnectorRunProfileConfiguration") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Step Number", typeof(int));
                var column22 = new DataColumn("Setting");
                var column32 = new DataColumn("Configuration");
                var column42 = new DataColumn("Setting Number", typeof(int));

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.PrimaryKey = new[] { column12, column22 };

                this.PilotDataSet = new DataSet("ConnectorRunProfileConfigurations") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorRunProfilePrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector run profile print table.
        /// </summary>
        /// <returns>
        /// The connector run profile print table.
        /// </returns>
        protected DataTable GetConnectorRunProfilePrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Step Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Step Name
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Step Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Setting
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Configuration
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Setting Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector run profile data set.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected virtual void FillConnectorRunProfileDataSet(string runProfileName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'. Pilot Config: '{1}'.", runProfileName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];
                    var table2 = dataSet.Tables[1];

                    var runProfileSteps = connector.XPathSelectElements("ma-run-data/run-configuration[name = '" + runProfileName + "']/configuration/step");
                    var stepIndex = 0;
                    foreach (var runProfileStep in runProfileSteps)
                    {
                        ++stepIndex;
                        var runProfileStepType = ConnectorDocumenter.GetRunProfileStepType(runProfileStep.Element("step-type"));

                        Documenter.AddRow(table, new object[] { stepIndex, runProfileStepType });

                        var logFileName = (string)runProfileStep.Element("dropfile-name");
                        if (!string.IsNullOrEmpty(logFileName))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Log file", logFileName, 1 });
                        }

                        var numberOfObjects = (string)runProfileStep.XPathSelectElement("threshold/object");
                        if (!string.IsNullOrEmpty(numberOfObjects))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Number of objects", numberOfObjects, 2 });
                        }

                        var numberOfDeletions = (string)runProfileStep.XPathSelectElement("threshold/delete");
                        if (!string.IsNullOrEmpty(numberOfDeletions))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Number of deletions", numberOfDeletions, 3 });
                        }

                        var partitionId = ((string)runProfileStep.Element("partition") ?? string.Empty).ToUpperInvariant();
                        var partitionName = (string)connector.XPathSelectElement("ma-partition-data/partition[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + partitionId + "']/name");
                        Documenter.AddRow(table2, new object[] { stepIndex, "Partition", partitionName, 4 });

                        var inputFileName = (string)connector.XPathSelectElement("custom-data/run-config/input-file");
                        if (!string.IsNullOrEmpty(inputFileName))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Input file name", inputFileName, 5 });
                        }

                        var outputFileName = (string)connector.XPathSelectElement("custom-data/run-config/output-file");
                        if (!string.IsNullOrEmpty(outputFileName))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Output file name", outputFileName, 6 });
                        }
                    }

                    table.AcceptChanges();
                    table2.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Run Profile Name: '{0}'. Pilot Config: '{1}'.", runProfileName, pilotConfig);
            }
        }

        /// <summary>
        /// Creates the connector run profile diffgram.
        /// </summary>
        protected void CreateConnectorRunProfileDiffgram()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = Documenter.GetDiffgram(this.PilotDataSet, this.ProductionDataSet);
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the extensible2 run profile.
        /// </summary>
        protected void PrintConnectorRunProfile()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Step#", 5 }, { "Step Name", 35 }, { "Setting", 35 }, { "Configuration", 25 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Run Profiles
    }
}
