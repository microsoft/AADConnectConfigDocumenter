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
    using System.Collections.Specialized;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Linq;
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
        /// <param name="productionOnly">If set to <c>true</c>, indicates the connector is present in production only.</param>
        public Extensible2ConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, bool productionOnly)
            : base(pilotXml, productionXml, connectorName, productionOnly)
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
        /// Gets the Extensible2 connector configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.WriteConnectorReportHeader();

                this.ProcessConnectorProperties();

                this.ProcessExtensible2ExtensionInformation();
                this.ProcessExtensible2ConnectivityInformation();

                if (!this.ConnectorSubType.StartsWith("Windows Azure Active Directory", StringComparison.OrdinalIgnoreCase))
                {
                    // TODO: Capabilities
                    this.ProcessExtensible2GlobalParameters();
                    this.ProcessConnectorProvisioningHierarchyConfiguration();
                }

                this.ProcessExtensible2PartitionsAndHierarchiesConfiguration();
                this.ProcessExtensible2AnchorConfigurations();

                this.ProcessConnectorSelectedObjectTypes();
                this.ProcessConnectorSelectedAttributes();
                this.ProcessConnectorProvisioningSyncRules();
                this.ProcessConnectorStickyJoinSyncRules();
                this.ProcessConnectorNormalJoinSyncRules();
                this.ProcessConnectorSyncRules();
                this.ProcessExtensible2RunProfiles();

                return base.GetReport();
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

        #region Extension Information

        /// <summary>
        /// Processes the extensible2 extension information.
        /// </summary>
        private void ProcessExtensible2ExtensionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Extension Information.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order, 2 = Capability Name, 3 = Configuration

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
        private void FillExtensible2ExtensionInformationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];

                    var importEnabled = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/import-enabled") == "1";
                    var exportEnabled = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/export-enabled") == "1";
                    var stepTypesSupported = importEnabled && exportEnabled ? "Import and Export" : importEnabled ? "Import Only" : exportEnabled ? "Export Only" : "??";

                    Documenter.AddRow(table, new object[] { 1, "Step Types supported", stepTypesSupported });

                    if (importEnabled)
                    {
                        var importMode = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/import-mode");
                        Documenter.AddRow(table, new object[] { 2, "Import Mode", importMode });
                    }

                    if (exportEnabled)
                    {
                        var exportMode = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/export-mode");
                        Documenter.AddRow(table, new object[] { 3, "Export Mode", exportMode });
                    }

                    var extensionDll = (string)connector.XPathSelectElement("private-configuration/MAConfig/extension-config/filename");
                    Documenter.AddRow(table, new object[] { 4, "Connected data source extension name", extensionDll });

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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintExtensible2ExtensionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Extension Information";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h3");

                #endregion section

                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table" + " " + this.GetCssVisibilityClass());
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                {
                    #region thead

                    this.ReportWriter.WriteBeginTag("thead");
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Setting");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Configuration");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
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
        private void ProcessExtensible2ConnectivityInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connectivity Information");

                this.CreateSimpleOrderedSettingsDataSets(4); // 1= Setting Display Order, 2 = Setting, 3 = Configuration, 4 = Encrypted?

                this.FillExtensible2ConnectivityInformationDataSet(true);
                this.FillExtensible2ConnectivityInformationDataSet(false);

                this.CreateSimpleSettingsDiffgram();

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
        private void FillExtensible2ConnectivityInformationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];

                    var parameters = connector.XPathSelectElements("private-configuration/MAConfig/parameter-values/parameter[@use='connectivity']");

                    for (var parameterIndex = 0; parameterIndex < parameters.Count(); ++parameterIndex)
                    {
                        var parameter = parameters.ElementAt(parameterIndex);
                        var encrypted = (string)parameter.Attribute("encrypted") == "1";
                        Documenter.AddRow(table, new object[] { parameterIndex, (string)parameter.Attribute("name"), encrypted ? "******" : (string)parameter, encrypted ? "Yes" : "No" });
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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintExtensible2ConnectivityInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Connectivity Information";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h3");

                #endregion section

                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table" + " " + this.GetCssVisibilityClass());
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                {
                    #region thead

                    this.ReportWriter.WriteBeginTag("thead");
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Setting");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Configuration");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Encrypted?");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Connectivity Information

        #region Global Parameters

        /// <summary>
        /// Processes the extensible2 global parameters.
        /// </summary>
        private void ProcessExtensible2GlobalParameters()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Global Parameters");

                this.CreateSimpleOrderedSettingsDataSets(4); // 1= Setting Display Order, 2 = Setting, 3 = Configuration, 4 = Encrypted?

                this.FillExtensible2GlobalParametersDataSet(true);
                this.FillExtensible2GlobalParametersDataSet(false);

                this.CreateSimpleSettingsDiffgram();

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
        private void FillExtensible2GlobalParametersDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];

                    var parameters = connector.XPathSelectElements("private-configuration/MAConfig/parameter-values/parameter[@use='global']");

                    for (var parameterIndex = 0; parameterIndex < parameters.Count(); ++parameterIndex)
                    {
                        var parameter = parameters.ElementAt(parameterIndex);
                        var encrypted = (string)parameter.Attribute("encrypted") == "1";
                        Documenter.AddRow(table, new object[] { parameterIndex, (string)parameter.Attribute("name"), encrypted ? "******" : (string)parameter, encrypted ? "Yes" : "No" });
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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintExtensible2GlobalParameters()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Global Parameters";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h3");

                #endregion section

                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table" + " " + this.GetCssVisibilityClass());
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                {
                    #region thead

                    this.ReportWriter.WriteBeginTag("thead");
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Setting");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Configuration");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Encrypted?");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessExtensible2PartitionsAndHierarchiesConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Extensible2 Partitions and Hierarchies Configuration.");

                var sectionTitle = "Partitions and Hierarchies";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h3");

                #endregion section

                // TODO: Extensible2 Partitions And Hierarchies
                this.ReportWriter.WriteLine();
                this.ReportWriter.WriteLine("<b>TO BE COMPLETED</b>");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Partitions and Hierarchies

        #region Achor Configuration

        /// <summary>
        /// Processes the extensible2 anchor configurations.
        /// </summary>
        private void ProcessExtensible2AnchorConfigurations()
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
        private void CreateExtensible2AnchorConfigurationsDataSets()
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
        private DataTable GetExtensible2AnchorConfigurationsPrintTable()
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
        private void FillExtensible2AnchorConfigurationsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

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

                            for (var attributeIndex = 0; attributeIndex < anchorAttributes.Count(); ++attributeIndex)
                            {
                                Documenter.AddRow(table2, new object[] { objectName, (string)anchorAttributes.ElementAt(attributeIndex), attributeIndex });
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
        private void CreateExtensible2AnchorConfigurationsDiffgram()
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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintExtensible2AnchorConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Anchors Configuration";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h3");

                #endregion section

                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table" + " " + this.GetCssVisibilityClass());
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                {
                    #region thead

                    this.ReportWriter.WriteBeginTag("thead");
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Object Type");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Anchor Attribute");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
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
        /// Processes the connector run profiles.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessExtensible2RunProfiles()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Run Profiles.");

                var sectionTitle = "Run Profiles";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h3");

                #endregion section

                var xpath = "//ma-data[name ='" + this.ConnectorName + "']/ma-run-data/run-configuration";

                var pilot = this.PilotXml.XPathSelectElements(xpath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(xpath, Documenter.NamespaceManager);

                // Sort by name
                var pilotRunProfiles = from runProfile in pilot
                                       let name = (string)runProfile.Element("name")
                                       orderby name
                                       select name;

                foreach (var runProfile in pilotRunProfiles)
                {
                    this.ProcessExtensible2RunProfile(runProfile);
                }

                // Sort by name
                var productionRunProfiles = from runProfile in production
                                            let name = (string)runProfile.Element("name")
                                            orderby name
                                            select name;

                productionRunProfiles = productionRunProfiles.Where(productionRunProfile => !pilotRunProfiles.Contains(productionRunProfile));

                foreach (var runProfile in productionRunProfiles)
                {
                    this.ProcessExtensible2RunProfile(runProfile);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the extensible2 run profile.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessExtensible2RunProfile(string runProfileName)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'.", runProfileName);

            try
            {
                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc4");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, runProfileName, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h4");
                Documenter.WriteBookmarkLocation(this.ReportWriter, runProfileName, "Run Profile: " + runProfileName, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h4");

                #endregion section

                this.CreateConnectorRunProfileDataSets();

                this.FillExtensible2RunProfileDataSet(runProfileName, true);
                this.FillExtensible2RunProfileDataSet(runProfileName, false);

                this.CreateConnectorRunProfileDiffgram();

                this.PrintConnectorRunProfile();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Run Profile Name: '{0}'.", runProfileName);
            }
        }

        /// <summary>
        /// Fills the extensible2 run profile data set.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillExtensible2RunProfileDataSet(string runProfileName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'. Pilot Config: '{1}'.", runProfileName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];
                    var table2 = dataSet.Tables[1];

                    var runProfileSteps = connector.XPathSelectElements("ma-run-data/run-configuration[name = '" + runProfileName + "']/configuration/step");

                    for (var stepIndex = 1; stepIndex <= runProfileSteps.Count(); ++stepIndex)
                    {
                        var runProfileStep = runProfileSteps.ElementAt(stepIndex - 1);

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

                        var pageSize = (string)runProfileStep.XPathSelectElement("custom-data/extensible2-step-data/batch-size");
                        if (!string.IsNullOrEmpty(pageSize))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Page Size (objects)", pageSize, 7 });
                        }

                        var timeout = (string)runProfileStep.XPathSelectElement("custom-data/extensible2-step-data/timeout");
                        if (!string.IsNullOrEmpty(timeout))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Timeout (seconds)", timeout, 8 });
                        }

                        var parameters = runProfileStep.XPathSelectElements("custom-data/parameter-values/parameter");

                        for (var parameterIndex = 0; parameterIndex < parameters.Count(); ++parameterIndex)
                        {
                            var parameter = parameters.ElementAt(parameterIndex);
                            Documenter.AddRow(table2, new object[] { stepIndex, (string)parameter.Attribute("name"), (string)parameter, 8 + parameterIndex });
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
