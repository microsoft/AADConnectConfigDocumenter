//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Extensible2ConnectorDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync ECMA2 Connector Configuration Documenter
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace AzureADConnectConfigDocumenter
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The Extensible2ConnectorDocumenter documents the configuration of any ECMA 2 connector.
    /// </summary>
    internal class Extensible2ConnectorDocumenter : ConnectorDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Extensible2ConnectorDocumenter"/> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The connector name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        public Extensible2ConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
            : base(pilotXml, productionXml, connectorName, configEnvironment)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportFileName = Documenter.GetTempFilePath(this.ConnectorName + ".tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath(this.ConnectorName + ".TOC.tmp.html");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Enumerator indicating whether the config element is present in pilot, production or both
        /// </summary>
        public enum ConfigParameterPage : int
        {
            /// <summary>
            /// The Connectivity page
            /// </summary>
            Connectivity = 0,

            /// <summary>
            /// The Global parameters page
            /// </summary>
            Global,

            /// <summary>
            /// The Partition properties page
            /// </summary>
            Partition,

            /// <summary>
            /// The Run Step properties page
            /// </summary>
            RunStep
        }

        /// <summary>
        /// Gets the Extensible2 connector configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.WriteConnectorReportHeader();

                this.ProcessConnectorProperties();

                this.ProcessExtensible2ExtensionInformation();
                this.ProcessExtensible2ConnectivityInformation();

                this.ProcessExtensible2ConnectorCapabilities();
                this.ProcessExtensible2GlobalParameters();
                this.ProcessConnectorProvisioningHierarchyConfiguration();

                this.ProcessExtensible2PartitionsAndHierarchiesConfiguration();

                this.ProcessConnectorSelectedObjectTypes();
                this.ProcessConnectorSelectedAttributes();
                this.ProcessExtensible2AnchorConfigurations();
                this.ProcessConnectorProvisioningSyncRules();
                this.ProcessConnectorStickyJoinSyncRules();
                this.ProcessConnectorNormalJoinSyncRules();
                this.ProcessConnectorSyncRules();
                this.ProcessConnectorRunProfiles();

                return this.GetReportTuple();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();

                // Clear Logger call context items
                Logger.ClearContextItem(ConnectorDocumenter.LoggerContextItemConnectorName);
                Logger.ClearContextItem(ConnectorDocumenter.LoggerContextItemConnectorGuid);
                Logger.ClearContextItem(ConnectorDocumenter.LoggerContextItemConnectorCategory);
            }
        }

        /// <summary>
        /// Gets the extensible2 config parameters data table.
        /// </summary>
        /// <param name="parameterDefinitions">The config parameter definition node for a configuration page.</param>
        /// <param name="parameterValues">he config parameter values node for the corresponding configuration page.</param>
        /// <returns>The config parameter values table.</returns>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected DataTable GetExtensible2ConfigParametersTable(IEnumerable<XElement> parameterDefinitions, IEnumerable<XElement> parameterValues)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var parametersTable = new DataTable("ConfigParametersTable") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("DisplayOrder", typeof(int));
                var column2 = new DataColumn("Setting");
                var column3 = new DataColumn("Configuration");
                var column4 = new DataColumn("Encrypted?");

                parametersTable.Columns.Add(column1);
                parametersTable.Columns.Add(column2);
                parametersTable.Columns.Add(column3);
                parametersTable.Columns.Add(column4);
                parametersTable.PrimaryKey = new[] { column1, column2 };

                if (parameterDefinitions != null && parameterValues != null)
                {
                    var parameterIndex = -1;
                    foreach (var parameterDefinition in parameterDefinitions)
                    {
                        ++parameterIndex;
                        var parameterName = (string)parameterDefinition.Element("name");
                        var parameterType = (string)parameterDefinition.Element("type") ?? string.Empty;
                        var parameterValue = string.Empty;

                        foreach (var parameter in parameterValues)
                        {
                            if ((string)parameter.Attribute("name") == parameterName)
                            {
                                parameterValue = (string)parameter;
                                break;
                            }
                        }

                        var encrypted = parameterType.StartsWith("encrypted", StringComparison.OrdinalIgnoreCase);

                        switch (parameterType.ToUpperInvariant())
                        {
                            case "FILE":
                                try
                                {
                                    parameterValue = Encoding.UTF8.GetString(Convert.FromBase64String(parameterValue));
                                }
                                catch (Exception e)
                                {
                                    Logger.Instance.WriteError(e.ToString());
                                }

                                break;
                            case "CHECKBOX":
                                parameterValue = parameterValue.Equals("1", StringComparison.OrdinalIgnoreCase) || parameterValue.Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";
                                break;
                        }

                        Documenter.AddRow(parametersTable, new object[] { parameterIndex, parameterName, encrypted ? "******" : parameterValue, encrypted ? "Yes" : "No" });
                    }

                    parametersTable.AcceptChanges();
                }

                return parametersTable;
            }
            finally
            {
                Logger.Instance.WriteMethodEntry();
            }
        }

        #region Extension Information

        /// <summary>
        /// Processes the extensible2 extension information.
        /// </summary>
        protected void ProcessExtensible2ExtensionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Extension Information.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Capability Name, 3 = Configuration

                this.FillExtensible2ExtensionInformationDataSet(true);
                this.FillExtensible2ExtensionInformationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintExtensible2ExtensionInformation();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the extensible2 extension information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillExtensible2ExtensionInformationDataSet(bool pilotConfig)
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

                    var connectorAssembly = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/filename");
                    var connectorAssemblyVersion = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/assembly-version");
                    var connectorCapabilityBits = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/capability-bits");

                    Documenter.AddRow(table, new object[] { 1, "Connector Assembly Name", connectorAssembly });
                    Documenter.AddRow(table, new object[] { 2, "Connector Assembly Version", connectorAssemblyVersion });
                    Documenter.AddRow(table, new object[] { 3, "Connector Capability Bits", connectorCapabilityBits });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the extensible2 extension information.
        /// </summary>
        protected void PrintExtensible2ExtensionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Connector Capabilities";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 30 }, { "Configuration", 70 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Extension Information

        #region Connectivity Information

        /// <summary>
        /// Processes the extensible2 connectivity information.
        /// </summary>
        protected void ProcessExtensible2ConnectivityInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connectivity Information");

                this.CreateSimpleOrderedSettingsDataSets(4); // 1= Setting Display Order, 2 = Setting, 3 = Configuration, 4 = Encrypted?

                this.FillExtensible2ConnectivityInformationDataSet(true);
                this.FillExtensible2ConnectivityInformationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintExtensible2ConnectivityInformation();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the extensible2 connectivity information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillExtensible2ConnectivityInformationDataSet(bool pilotConfig)
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

                    var parameterDefinitions = connector.XPathSelectElements("private-configuration/MAConfig/parameter-definitions/parameter[use = 'connectivity' and type != 'label' and type != 'divider']");
                    var parameterValues = connector.XPathSelectElements("private-configuration/MAConfig/parameter-values/parameter[@use = 'connectivity']");

                    var parameterTable = this.GetExtensible2ConfigParametersTable(parameterDefinitions, parameterValues);
                    foreach (DataRow row in parameterTable.Rows)
                    {
                        Documenter.AddRow(table, new object[] { row[0], row[1], row[2], row[3] });
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
        /// Prints the extensible2 connectivity information.
        /// </summary>
        protected void PrintExtensible2ConnectivityInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Connectivity Information";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 25 }, { "Configuration", 65 }, { "Encrypted?", 10 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Connectivity Information

        #region Connector Capabilities
        /// <summary>
        /// Processes the extensible2 connector capabilities.
        /// </summary>
        protected void ProcessExtensible2ConnectorCapabilities()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                // TODO: Capabilities
                ////Logger.Instance.WriteInfo("Processing Connector Capabilities");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Connector Capabilities

        #region Global Parameters

        /// <summary>
        /// Processes the extensible2 global parameters.
        /// </summary>
        protected void ProcessExtensible2GlobalParameters()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Global Parameters");

                this.CreateSimpleOrderedSettingsDataSets(4); // 1= Setting Display Order, 2 = Setting, 3 = Configuration, 4 = Encrypted?

                this.FillExtensible2GlobalParametersDataSet(true);
                this.FillExtensible2GlobalParametersDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintExtensible2GlobalParameters();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the extensible2 global parameters data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillExtensible2GlobalParametersDataSet(bool pilotConfig)
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

                    var parameterDefinitions = connector.XPathSelectElements("private-configuration/MAConfig/parameter-definitions/parameter[use = 'global' and type != 'label' and type != 'divider']");
                    var parameterValues = connector.XPathSelectElements("private-configuration/MAConfig/parameter-values/parameter[@use = 'global']");

                    var parameterTable = this.GetExtensible2ConfigParametersTable(parameterDefinitions, parameterValues);
                    foreach (DataRow row in parameterTable.Rows)
                    {
                        Documenter.AddRow(table, new object[] { row[0], row[1], row[2], row[3] });
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
        /// Prints the extensible2 global parameters.
        /// </summary>
        protected void PrintExtensible2GlobalParameters()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Global Parameters";

                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 25 }, { "Configuration", 65 }, { "Encrypted?", 10 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
                else
                {
                    this.WriteContentParagraph("There are no Global parameters configured.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Global Parameters

        #region Partitions and Hierarchies

        /// <summary>
        /// Processes the extensible2 partitions and hierarchies configuration.
        /// </summary>
        protected void ProcessExtensible2PartitionsAndHierarchiesConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Extensible2 Partitions and Hierarchies Configuration.");

                var sectionTitle = "Partitions and Hierarchies";

                this.WriteSectionHeader(sectionTitle, 3);

                var xpath = "/ma-data[name ='" + this.ConnectorName + "']" + "//ma-partition-data/partition[selected = 1]";

                var pilot = this.PilotXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(true) + xpath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(false) + xpath, Documenter.NamespaceManager);

                // Sort by name
                var pilotPartitions = from partition in pilot
                                      let name = (string)partition.Element("name")
                                      orderby name
                                      select name;

                foreach (var partition in pilotPartitions)
                {
                    this.ProcessExtensible2Partition(partition);
                }

                // Sort by name
                var productionPartitions = from partition in production
                                           let name = (string)partition.Element("name")
                                           orderby name
                                           select name;

                productionPartitions = productionPartitions.Where(productionPartition => !pilotPartitions.Contains(productionPartition));

                foreach (var partition in productionPartitions)
                {
                    this.ProcessExtensible2Partition(partition);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the extensible2 partition.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        protected void ProcessExtensible2Partition(string partitionName)
        {
            Logger.Instance.WriteMethodEntry("Partition: '{0}'.", partitionName);

            try
            {
                this.WriteSectionHeader("Partition: " + partitionName, 4, partitionName);

                // Partion Settings
                this.CreateSimpleOrderedSettingsDataSets(4); // 1= Setting Display Order, 2 = Setting, 3 = Configuration, 4 = Encrypted?

                this.FillExtensible2PartitionParametersDataSet(partitionName, true);
                this.FillExtensible2PartitionParametersDataSet(partitionName, false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintExtensible2PartitionParameters();

                // Container Include / Exclude settings
                this.ProcessConnectorPartitionContainers(partitionName);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Partition: '{0}'.", partitionName);
            }
        }

        #region Partition Parameters

        /// <summary>
        /// Fills the extensible2 partition parameters data set.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillExtensible2PartitionParametersDataSet(string partitionName, bool pilotConfig)
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

                    var partition = connector.XPathSelectElement("ma-partition-data/partition[selected = 1 and name = '" + partitionName + "']");

                    if (partition != null)
                    {
                        var parameterDefinitions = connector.XPathSelectElements("private-configuration/MAConfig/parameter-definitions/parameter[use = 'partition' and type != 'label' and type != 'divider']");
                        var parameterValues = partition.XPathSelectElements("custom-data/parameter-values/parameter");

                        var parameterTable = this.GetExtensible2ConfigParametersTable(parameterDefinitions, parameterValues);
                        foreach (DataRow row in parameterTable.Rows)
                        {
                            Documenter.AddRow(table, new object[] { row[0], row[1], row[2], row[3] });
                        }

                        table.AcceptChanges();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the extensible2 partition parameters.
        /// </summary>
        protected void PrintExtensible2PartitionParameters()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Partition Parameter", 25 }, { "Configuration", 65 }, { "Encrypted?", 10 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Partition Parameters

        #endregion Partitions and Hierarchies

        #region Achor Configuration

        /// <summary>
        /// Processes the extensible2 anchor configurations.
        /// </summary>
        protected void ProcessExtensible2AnchorConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Extensible2 Anchor Configuration.");

                this.CreateExtensible2AnchorConfigurationsDataSets();

                this.FillExtensible2AnchorConfigurationsDataSet(true);
                this.FillExtensible2AnchorConfigurationsDataSet(false);

                this.CreateExtensible2AnchorConfigurationsDiffgram();

                this.PrintExtensible2AnchorConfigurations();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the extensible2 anchor configurations data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateExtensible2AnchorConfigurationsDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("ConnectorObjectTypes") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Object Type");

                table.Columns.Add(column1);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("ObjectTypeAnchors") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Object Type");
                var column22 = new DataColumn("Anchor Attribute");
                var column32 = new DataColumn("Anchor Order", typeof(int));

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.PrimaryKey = new[] { column12, column22, column32 };

                this.PilotDataSet = new DataSet("ConnectorAnchorConfigurations") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetExtensible2AnchorConfigurationsPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the extensible2 anchor configurations print table.
        /// </summary>
        /// <returns>
        /// The extensible2 anchor configurations print table.
        /// </returns>
        protected DataTable GetExtensible2AnchorConfigurationsPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Anchor Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Anchor Order
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the extensible2 anchor configurations data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillExtensible2AnchorConfigurationsDataSet(bool pilotConfig)
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
                    var table2 = dataSet.Tables[1];

                    var objectTypes = connector.XPathSelectElements("private-configuration/MAConfig/importing/per-class-settings/class");

                    foreach (var objectType in objectTypes)
                    {
                        var objectName = (string)objectType.Element("name");
                        Documenter.AddRow(table, new object[] { objectName });

                        var anchor = objectType.Element("anchor");
                        if (anchor != null)
                        {
                            var anchorAttributes = anchor.Elements("attribute");

                            var attributeIndex = -1;
                            foreach (var anchorAttribute in anchorAttributes)
                            {
                                ++attributeIndex;
                                Documenter.AddRow(table2, new object[] { objectName, (string)anchorAttribute, attributeIndex });
                            }
                        }
                    }

                    table.AcceptChanges();
                    table2.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the extensible2 anchor configurations diffgram.
        /// </summary>
        protected void CreateExtensible2AnchorConfigurationsDiffgram()
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
        /// Prints the extensible2 anchor configurations.
        /// </summary>
        protected void PrintExtensible2AnchorConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Anchor Configuration";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Object Type", 30 }, { "Anchor Attribute", 70 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Achor Configuration

        #region Run Profiles

        /// <summary>
        /// Fills the extensible2 run profile data set.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected override void FillConnectorRunProfileDataSet(string runProfileName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'. Pilot Config: '{1}'.", runProfileName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    base.FillConnectorRunProfileDataSet(runProfileName, pilotConfig);

                    var table = dataSet.Tables[0];
                    var table2 = dataSet.Tables[1];

                    var runProfileSteps = connector.XPathSelectElements("ma-run-data/run-configuration[name = '" + runProfileName + "']/configuration/step");

                    var stepIndex = 0;
                    foreach (var runProfileStep in runProfileSteps)
                    {
                        ++stepIndex;
                        var batchSize = (string)runProfileStep.XPathSelectElement("custom-data/extensible2-step-data/batch-size");
                        if (!string.IsNullOrEmpty(batchSize))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Batch Size (objects)", batchSize, 1000 });
                        }

                        var timeout = (string)runProfileStep.XPathSelectElement("custom-data/extensible2-step-data/timeout");
                        if (!string.IsNullOrEmpty(timeout))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Timeout (seconds)", timeout, 1001 });
                        }

                        var parameterDefinitions = connector.XPathSelectElements("private-configuration/MAConfig/parameter-definitions/parameter[use = 'run-step' and type != 'label' and type != 'divider']");
                        var parameterValues = runProfileStep.XPathSelectElements("custom-data/parameter-values/parameter");

                        var parameterTable = this.GetExtensible2ConfigParametersTable(parameterDefinitions, parameterValues);
                        foreach (DataRow row in parameterTable.Rows)
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, row[1], row[2], 1000 + (int)row[0] });
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

        #endregion Run Profiles
    }
}
