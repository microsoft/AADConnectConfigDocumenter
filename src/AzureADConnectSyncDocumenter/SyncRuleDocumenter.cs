//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="SyncRuleDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync Sync Rule Configuration Documenter
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
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The SyncRuleDocumenter documents the configuration of a Synchronization Rule
    /// </summary>
    internal class SyncRuleDocumenter : ConnectorDocumenter
    {
        /// <summary>
        /// The logger context item synchronize rule name
        /// </summary>
        private const string LoggerContextItemSyncRuleName = "Sync Rule Name";

        /// <summary>
        /// The logger context item synchronize rule unique identifier
        /// </summary>
        private const string LoggerContextItemSyncRuleGuid = "Sync Rule Guid";

        /// <summary>
        /// The logger context item synchronize rule report type
        /// </summary>
        private const string LoggerContextItemSyncRuleReportType = "Sync Rule Report Type";

        /// <summary>
        /// The synchronization rule direction
        /// </summary>
        private SyncRuleDirection syncRuleDirection;

        /// <summary>
        /// The synchronization rule report type
        /// </summary>
        private SyncRuleReportType syncRuleReportType;

        /// <summary>
        /// Indicates if the synchronization rule is inferred as a default sync rule
        /// </summary>
        private bool defaultSyncRule;

        /// <summary>
        /// Indicates if the default synchronization rule has to be visible all the time
        /// </summary>
        private bool defaultSyncRuleVisibility;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncRuleDocumenter" /> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="syncRuleName">Name of the synchronize rule.</param>
        /// <param name="syncRuleGuid">The synchronize rule unique identifier.</param>
        /// <param name="connectorName">The connector name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        public SyncRuleDocumenter(XElement pilotXml, XElement productionXml, string syncRuleName, string syncRuleGuid, string connectorName, ConfigEnvironment configEnvironment)
            : base(pilotXml, productionXml, connectorName, configEnvironment)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Name: '{0}'. Sync Rule Name: '{1}'.", syncRuleName, syncRuleGuid);

            try
            {
                this.SyncRuleName = syncRuleName;
                this.SyncRuleGuid = syncRuleGuid;
                this.defaultSyncRule = false;
                this.defaultSyncRuleVisibility = false;
                this.ReportFileName = Documenter.GetTempFilePath(this.SyncRuleGuid + ".tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath(this.SyncRuleGuid + ".TOC.tmp.html");

                // Set Logger call context items
                Logger.SetContextItem(SyncRuleDocumenter.LoggerContextItemSyncRuleName, this.SyncRuleName);
                Logger.SetContextItem(SyncRuleDocumenter.LoggerContextItemSyncRuleGuid, this.SyncRuleGuid);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Sync rule direction
        /// </summary>
        public enum SyncRuleDirection
        {
            /// <summary>
            /// The inbound direction
            /// </summary>
            Inbound,

            /// <summary>
            /// The outbound direction
            /// </summary>
            Outbound
        }

        /// <summary>
        /// Sync rule report type
        /// </summary>
        public enum SyncRuleReportType
        {
            /// <summary>
            /// All sections
            /// </summary>
            AllSections,

            /// <summary>
            /// The provisioning section report
            /// </summary>
            ProvisioningSection,

            /// <summary>
            /// The sticky join section report
            /// </summary>
            StickyJoinSection,

            /// <summary>
            /// The conditional join (i.e. scoped or explicit join) section report
            /// </summary>
            ConditionalJoinSection
        }

        /// <summary>
        /// Gets the name of the synchronize rule.
        /// </summary>
        /// <value>
        /// The name of the synchronize rule.
        /// </value>
        protected string SyncRuleName { get; private set; }

        /// <summary>
        /// Gets the synchronize rule unique identifier.
        /// </summary>
        /// <value>
        /// The synchronize rule unique identifier.
        /// </value>
        protected string SyncRuleGuid { get; private set; }

        /// <summary>
        /// Gets the sync rule configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                return this.GetReport(SyncRuleReportType.AllSections);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the report.
        /// </summary>
        /// <param name="reportType">Type of the report.</param>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        public Tuple<string, string> GetReport(SyncRuleReportType reportType)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Report Type: '{0}'.", reportType);

            // Set Logger call context items
            Logger.SetContextItem(SyncRuleDocumenter.LoggerContextItemSyncRuleReportType, reportType);

            try
            {
                this.syncRuleReportType = reportType;

                this.ProcessConnectorSyncRuleDescription();
                this.ProcessConnectorSyncRuleScopingFilter();
                this.ProcessConnectorSyncRuleJoinRules();

                if (reportType == SyncRuleReportType.AllSections)
                {
                    this.ProcessConnectorSyncRuleTransformations();
                }

                // for default rule it can be hidden if the only change is to the precedence number
                if (this.defaultSyncRule)
                {
                    var filterExpression = "[Column2] <> 'Precedence' AND [Column2] <> 'Tag' AND [" + Documenter.OldColumnPrefix + "Column3] <> [Column3]";
                    var defaultSyncRuleDescriptionChanged = this.DiffgramDataSets.Count > 0 && this.DiffgramDataSets[0].Tables[0].Select(filterExpression).Count() != 0;
                    var defaultSyncRuleScopingFilterChanged = this.DiffgramDataSets.Count > 1 && !(bool)this.DiffgramDataSets[1].ExtendedProperties[Documenter.CanHide];
                    var defaultSyncRuleJoinRulesChanged = this.DiffgramDataSets.Count > 2  && !(bool)this.DiffgramDataSets[2].ExtendedProperties[Documenter.CanHide];
                    var defaultSyncRuleTransformationsChanged = this.DiffgramDataSets.Count > 3 && !(bool)this.DiffgramDataSets[3].ExtendedProperties[Documenter.CanHide];

                    this.defaultSyncRuleVisibility = defaultSyncRuleDescriptionChanged || defaultSyncRuleScopingFilterChanged || defaultSyncRuleJoinRulesChanged || defaultSyncRuleTransformationsChanged;
                }

                var noHide = this.defaultSyncRule ? this.defaultSyncRuleVisibility != false : this.DiffgramDataSets.Any(dataSet => !(bool)dataSet.ExtendedProperties[Documenter.CanHide]);

                if (noHide)
                {
                    // Update HtmlTableRowVisibilityStatusColumn for all rows to NoHide 
                    foreach (var dataSet in this.DiffgramDataSets)
                    {
                        dataSet.ExtendedProperties[Documenter.CanHide] = false;
                        foreach (DataTable table in dataSet.Tables)
                        {
                            if (!table.TableName.Equals("PrintSettings", StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (DataRow row in table.Rows)
                                {
                                    row[Documenter.HtmlTableRowVisibilityStatusColumn] = Documenter.NoHide;
                                }
                            }
                        }

                        dataSet.AcceptChanges();
                    }
                }

                this.WriteSyncRuleReportHeader();

                try
                {
                    this.PrintConnectorSyncRuleDescription();
                    this.PrintConnectorSyncRuleScopingFilter();
                    this.PrintConnectorSyncRuleJoinRules();
                    if (reportType == SyncRuleReportType.AllSections)
                    {
                        this.PrintConnectorSyncRuleTransformations();
                        if (noHide)
                        {
                            this.CreateConnectorSyncRuleInstallationScript();
                        }
                    }
                }
                finally
                {
                    if (this.defaultSyncRule && this.defaultSyncRuleVisibility == false)
                    {
                        this.ReportWriter.WriteEndTag("div");
                        this.ReportToCWriter.WriteEndTag("div");
                    }

                    this.ResetDiffgram(); // reset the diffgram variables
                }

                return this.GetReportTuple();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();

                // Clear Logger call context items
                Logger.ClearContextItem(SyncRuleDocumenter.LoggerContextItemSyncRuleName);
                Logger.ClearContextItem(SyncRuleDocumenter.LoggerContextItemSyncRuleGuid);
                Logger.ClearContextItem(SyncRuleDocumenter.LoggerContextItemSyncRuleReportType);
            }
        }

        /// <summary>
        /// Writes the synchronize rule report header.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        private void WriteSyncRuleReportHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var bookmark = this.SyncRuleName;
                if (this.syncRuleReportType != SyncRuleReportType.AllSections)
                {
                    bookmark += this.syncRuleReportType.ToString();
                    this.ReportFileName = Documenter.GetTempFilePath(this.SyncRuleGuid + ".tmp.html");
                    this.ReportToCFileName = Documenter.GetTempFilePath(this.SyncRuleGuid + ".TOC.tmp.html");
                }

                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                if (this.defaultSyncRule && this.defaultSyncRuleVisibility == false)
                {
                    this.ReportWriter.WriteBeginTag("div");
                    this.ReportWriter.WriteAttribute("class", "DefaultRuleCanHide");
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    this.ReportToCWriter.WriteBeginTag("div");
                    this.ReportToCWriter.WriteAttribute("class", "DefaultRuleCanHide");
                    this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                }

                this.WriteSectionHeader(this.SyncRuleName, 5, bookmark, this.SyncRuleGuid);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the sync rule xpath.
        /// </summary>
        /// <param name="currentConnectorGuid">The current connector unique identifier.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        /// <returns>
        /// The sync rule xpath.
        /// </returns>
        private string GetSyncRuleXPath(string currentConnectorGuid, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Current Connector Guid: '{0}'. Pilot Config: '{1}'.", currentConnectorGuid, pilotConfig);

            var xpath = Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + currentConnectorGuid + "' and name = '" + this.SyncRuleName + "'";

            try
            {
                switch (this.syncRuleReportType)
                {
                    case SyncRuleReportType.ProvisioningSection:
                        xpath += " and linkType = 'Provision' " + Documenter.SyncRuleDisabledCondition;
                        break;
                    case SyncRuleReportType.StickyJoinSection:
                        xpath += " and linkType = 'StickyJoin' " + Documenter.SyncRuleDisabledCondition;
                        break;
                    case SyncRuleReportType.ConditionalJoinSection:
                        xpath += " and linkType = 'Join' " + Documenter.SyncRuleDisabledCondition;
                        break;
                }

                xpath += "]";

                return xpath;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Current Connector Guid: '{0}'. Pilot Config: '{1}'. XPath: '{2}'.", currentConnectorGuid, pilotConfig, xpath);
            }
        }

        #region Sync Rule Description

        /// <summary>
        /// Processes the connector synchronize rule description.
        /// </summary>
        private void ProcessConnectorSyncRuleDescription()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Sync Rule Description.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Field Name, 3 = Value

                this.FillConnectorSyncRuleDescriptionDataSet(true);
                this.FillConnectorSyncRuleDescriptionDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector synchronize rule description data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillConnectorSyncRuleDescriptionDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var currentConnectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant(); // This may be pilot or production GUID

                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid, pilotConfig);

                    var syncRule = config.XPathSelectElement(xpath);

                    if (syncRule != null)
                    {
                        var table = dataSet.Tables[0];

                        this.syncRuleDirection = (SyncRuleDirection)Enum.Parse(typeof(SyncRuleDirection), (string)syncRule.Element("direction"), true);
                        this.defaultSyncRule = ((string)syncRule.Element("immutable-tag") ?? string.Empty).StartsWith("Microsoft.", StringComparison.OrdinalIgnoreCase);

                        var setting = (string)syncRule.Element("name");
                        Documenter.AddRow(table, new object[] { 0, "Name", setting });

                        if (this.syncRuleReportType == SyncRuleReportType.AllSections)
                        {
                            setting = (string)syncRule.Element("description");
                            Documenter.AddRow(table, new object[] { 1, "Description", setting });
                        }

                        setting = (string)syncRule.Element("direction");
                        Documenter.AddRow(table, new object[] { 2, "Direction", setting });

                        setting = this.ConnectorName;
                        Documenter.AddRow(table, new object[] { 3, "Connected System", setting });

                        setting = this.syncRuleDirection == SyncRuleDirection.Inbound ? (string)syncRule.Element("sourceObjectType") : (string)syncRule.Element("targetObjectType");
                        Documenter.AddRow(table, new object[] { 4, "Connected System Object Type", setting });

                        setting = this.syncRuleDirection == SyncRuleDirection.Outbound ? (string)syncRule.Element("sourceObjectType") : (string)syncRule.Element("targetObjectType");
                        Documenter.AddRow(table, new object[] { 5, "Metaverse Object Type", setting });

                        setting = (string)syncRule.Element("linkType");
                        Documenter.AddRow(table, new object[] { 6, "Link Type", setting });

                        if (this.syncRuleReportType == SyncRuleReportType.AllSections)
                        {
                            setting = (string)syncRule.Element("precedence");
                            Documenter.AddRow(table, new object[] { 7, "Precedence", setting });

                            setting = (string)syncRule.Element("softDeleteExpiryInterval");
                            Documenter.AddRow(table, new object[] { 8, "Soft Delete Expiry Interval", setting });

                            setting = (string)syncRule.Element("immutable-tag");
                            Documenter.AddRow(table, new object[] { 9, "Tag", setting });

                            setting = (string)syncRule.Element("EnablePasswordSync");
                            if (!string.IsNullOrEmpty(setting))
                            {
                                Documenter.AddRow(table, new object[] { 10, "Enable Password Sync", setting.Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No" });
                            }
                        }

                        setting = (string)syncRule.Element("disabled");
                        if (!string.IsNullOrEmpty(setting))
                        {
                            Documenter.AddRow(table, new object[] { 11, "Disabled", setting.Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No" });
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
        /// Gets the connector synchronize rule description header table.
        /// </summary>
        /// <returns>The connector synchronize rule description header table.</returns>
        private DataTable GetConnectorSyncRuleDescriptionHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Description
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", this.GetConnectorSyncRuleDescriptionHeader() }, { "RowSpan", 1 }, { "ColSpan", 2 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Setting
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Setting" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 30 } }).Values.Cast<object>().ToArray());

                // Configuration
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Configuration" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 70 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector synchronize rule description column header.
        /// </summary>
        /// <returns>The connector synchronize rule description column header.</returns>
        private string GetConnectorSyncRuleDescriptionHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var descriptionHeader = "Description";

                switch (this.syncRuleReportType)
                {
                    case SyncRuleReportType.ProvisioningSection:
                        descriptionHeader += " (Provisioning Rules Summary)";
                        break;
                    case SyncRuleReportType.StickyJoinSection:
                        descriptionHeader += " (Sticky Join Rules Summary)";
                        break;
                    case SyncRuleReportType.ConditionalJoinSection:
                        descriptionHeader += " (Conditional Join Rules Summary)";
                        break;
                }

                return descriptionHeader;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector synchronize rule description.
        /// </summary>
        private void PrintConnectorSyncRuleDescription()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = this.GetConnectorSyncRuleDescriptionHeaderTable();

                this.DiffgramDataSet = this.DiffgramDataSets[0];
                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, this.syncRuleReportType == SyncRuleReportType.AllSections ? HtmlTableSize.Huge : HtmlTableSize.Standard);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sync Rule Description

        #region Sync Rule Scoping Filters

        /// <summary>
        /// Processes the connector synchronize rule scoping filters.
        /// </summary>
        private void ProcessConnectorSyncRuleScopingFilter()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Sync Rule Scoping Filter.");

                this.CreateSimpleOrderedSettingsDataSets(5, 5, false); // 1 = Display Order Control, 2 = Group #, 3 = Attribute, 4 = Operator, 5 = Value

                this.FillConnectorSyncRuleScopingFilterDataSet(true);
                this.FillConnectorSyncRuleScopingFilterDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector synchronize rule scoping filter data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillConnectorSyncRuleScopingFilterDataSet(bool pilotConfig)
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
                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid, pilotConfig) + "/synchronizationCriteria/conditions";

                    var syncRuleScopingConditions = config.XPathSelectElements(xpath);
                    var conditionCount = syncRuleScopingConditions.Count();

                    if (conditionCount == 0)
                    {
                        Documenter.AddRow(table, new object[] { 0, "-", "-", "-", "-" }, true);
                    }
                    else
                    {
                        var conditionIndex = -1;
                        foreach (var condition in syncRuleScopingConditions)
                        {
                            ++conditionIndex;
                            var scopes = condition.Elements("scope");
                            foreach (var scope in scopes)
                            {
                                var scopeAttribute = (string)scope.Element("csAttribute") ?? " ";
                                var scopeOperator = (string)scope.Element("csOperator") ?? " ";
                                var scopeValue = (string)scope.Element("csValue") ?? " ";
                                Documenter.AddRow(table, new object[] { conditionIndex + 1, conditionIndex + 1, scopeAttribute, scopeOperator, scopeValue });
                            }
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
        /// Gets the connector synchronize rule scoping filter header table.
        /// </summary>
        /// <returns>The connector synchronize rule scoping filter header table.</returns>
        private DataTable GetConnectorSyncRuleScopingFilterHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Scoping Filter
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Scoping Filter" }, { "RowSpan", 1 }, { "ColSpan", 4 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Group#
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Group#" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 30 } }).Values.Cast<object>().ToArray());

                // Operator
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Operator" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 30 } }).Values.Cast<object>().ToArray());

                // Value
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Value" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 30 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector synchronize rule scoping filter.
        /// </summary>
        private void PrintConnectorSyncRuleScopingFilter()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = this.GetConnectorSyncRuleScopingFilterHeaderTable();

                this.DiffgramDataSet = this.DiffgramDataSets[1];
                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, this.syncRuleReportType == SyncRuleReportType.AllSections ? HtmlTableSize.Huge : HtmlTableSize.Standard);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sync Rule Scoping Filters

        #region Sync Rule Join Rules

        /// <summary>
        /// Processes the connector synchronize rule join rules.
        /// </summary>
        private void ProcessConnectorSyncRuleJoinRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Sync Rule Join Rules");

                this.CreateSimpleOrderedSettingsDataSets(5, 5, false); // 1 = Display Order Control, 2 = Group #, 3 = Source Attribute, 4 = Target Attribute, 5 = Case Sensitive?

                this.FillConnectorSyncRuleJoinRulesDataSet(true);
                this.FillConnectorSyncRuleJoinRulesDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector synchronize rule join rules data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillConnectorSyncRuleJoinRulesDataSet(bool pilotConfig)
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
                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid, pilotConfig);
                    var syncRule = config.XPathSelectElement(xpath);
                    xpath += "/relationshipCriteria/conditions";
                    var syncRuleJoiningRules = config.XPathSelectElements(xpath);
                    var joinRuleCount = syncRuleJoiningRules.Count();

                    if (joinRuleCount == 0)
                    {
                        Documenter.AddRow(table, new object[] { 0, "-", "-", "-", "-" }, true);
                    }
                    else
                    {
                        this.syncRuleDirection = (SyncRuleDirection)Enum.Parse(typeof(SyncRuleDirection), (string)syncRule.Element("direction"), true);

                        var joinRuleIndex = -1;
                        foreach (var joinRule in syncRuleJoiningRules)
                        {
                            ++joinRuleIndex;
                            var conditions = joinRule.Elements("condition");
                            foreach (var condition in conditions)
                            {
                                var csAttribute = (string)condition.Element("csAttribute") ?? " ";
                                var mvAttribute = (string)condition.Element("ilmAttribute") ?? " ";
                                var caseSensitive = (string)condition.Element("caseSensitive") ?? " ";
                                if (this.syncRuleDirection == SyncRuleDirection.Inbound)
                                {
                                    Documenter.AddRow(table, new object[] { joinRuleIndex + 1, joinRuleIndex + 1, csAttribute, mvAttribute, caseSensitive });
                                }
                                else
                                {
                                    Documenter.AddRow(table, new object[] { joinRuleIndex + 1, joinRuleIndex + 1, mvAttribute, csAttribute, caseSensitive });
                                }
                            }
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
        /// Gets the connector synchronize rule join rules header table.
        /// </summary>
        /// <returns>The connector synchronize rule join rules header table.</returns>
        private DataTable GetConnectorSyncRuleJoinRulesHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Join Rules
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Join Rules" }, { "RowSpan", 1 }, { "ColSpan", 4 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Group#
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Group#" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Source Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Source Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 40 } }).Values.Cast<object>().ToArray());

                // Target Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Target Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 40 } }).Values.Cast<object>().ToArray());

                // Case Sensitive
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Case Sensitive" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector synchronize rule join rules.
        /// </summary>
        private void PrintConnectorSyncRuleJoinRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = this.GetConnectorSyncRuleJoinRulesHeaderTable();

                this.DiffgramDataSet = this.DiffgramDataSets[2];
                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, this.syncRuleReportType == SyncRuleReportType.AllSections ? HtmlTableSize.Huge : HtmlTableSize.Standard);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sync Rule Join Rules

        #region Sync Rule Transformations

        /// <summary>
        /// Processes the connector synchronize rule transformations.
        /// </summary>
        private void ProcessConnectorSyncRuleTransformations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.CreateSimpleSettingsDataSets(5); // 1 = Target Attribute, 2 = Source, 3 = Flow Type, 4 = Apply Once, 5 = Merge Type

                this.FillConnectorSyncRuleTransformationsDataSet(true);
                this.FillConnectorSyncRuleTransformationsDataSet(false);

                this.CreateSimpleSettingsDiffgram();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector synchronize rule transformations data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillConnectorSyncRuleTransformationsDataSet(bool pilotConfig)
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
                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid, pilotConfig) + "/attribute-mappings/mapping";

                    var transformations = from transformation in config.XPathSelectElements(xpath)
                                          let target = (string)transformation.Element("dest")
                                          orderby target
                                          select transformation;

                    if (!transformations.Any())
                    {
                        Documenter.AddRow(table, new object[] { "-", "-", "-", "-", "-" }, true);
                    }
                    else
                    {
                        foreach (var transformation in transformations)
                        {
                            var targetAttribute = (string)transformation.Element("dest");
                            var expression = (string)transformation.Element("expression");
                            var valueMergeType = (string)transformation.Element("valueMergeType");
                            var applyOnce = (string)transformation.Attribute("execute-once");

                            if (!string.IsNullOrEmpty(expression))
                            {
                                Documenter.AddRow(table, new object[] { targetAttribute, expression, "Expression", applyOnce, valueMergeType });
                            }
                            else
                            {
                                var sourceAttribute = (string)transformation.XPathSelectElement("src/attr");

                                if (!string.IsNullOrEmpty(sourceAttribute))
                                {
                                    Documenter.AddRow(table, new object[] { targetAttribute, sourceAttribute, "Direct", applyOnce, valueMergeType });
                                }
                                else
                                {
                                    var source = (string)transformation.Element("src");
                                    Documenter.AddRow(table, new object[] { targetAttribute, source, "Constant", applyOnce, valueMergeType });
                                }
                            }
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
        /// Gets the connector synchronize rule transformations header table.
        /// </summary>
        /// <returns>The connector synchronize rule transformations header table.</returns>
        private DataTable GetConnectorSyncRuleTransformationsHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Transformations
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Transformations" }, { "RowSpan", 1 }, { "ColSpan", 5 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Target Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", this.syncRuleDirection == SyncRuleDirection.Inbound ? "Target (MV) Attribute" : "Target (CS) Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 20 } }).Values.Cast<object>().ToArray());

                // Source
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Source" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 50 } }).Values.Cast<object>().ToArray());

                // Flow Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Flow Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Apply Once
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Apply Once" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Merge Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 4 }, { "ColumnName", "Merge Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector synchronize rule transformations.
        /// </summary>
        private void PrintConnectorSyncRuleTransformations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = this.GetConnectorSyncRuleTransformationsHeaderTable();

                this.DiffgramDataSet = this.DiffgramDataSets[3];
                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, this.syncRuleReportType == SyncRuleReportType.AllSections ? HtmlTableSize.Huge : HtmlTableSize.Standard);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sync Rule Transformations

        #region Sync Rule Script Generation
        /// <summary>
        /// Creates Connector Sync Rule Installation Script
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "Reviewed.")]
        private void CreateConnectorSyncRuleInstallationScript()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Creating installation script for sync rule. Name = '{0}'. Id = '{1}'.", this.SyncRuleName, this.SyncRuleGuid);

                var config = this.Environment != ConfigEnvironment.ProductionOnly ? this.PilotXml : this.ProductionXml;
                var syncRule = config.XPathSelectElement(Documenter.GetSynchronizationRuleXmlRootXPath(this.Environment != ConfigEnvironment.ProductionOnly) + "/synchronizationRule[id = '" + this.SyncRuleGuid + "']");
                var script = string.Empty;

                // if the sync rule is part of the default config (i.e. starts with tag "Microsoft.")
                // we'll igonre any changes, except for the supported change i.e. to the Disabled settings.
                // else for any custom rule we'll create the complete script
                var immutableTag = (string)syncRule.Element("immutable-tag");
                if (!string.IsNullOrEmpty(immutableTag) && immutableTag.StartsWith("Microsoft.", StringComparison.OrdinalIgnoreCase))
                {
                    if (this.Environment == ConfigEnvironment.ProductionOnly)
                    {
                        // This sync rule is present only in the Production config. Give a warning.
                        script = this.CreateMissingDefaultSyncRuleWarningScript(syncRule, true);
                    }
                    else
                    {
                        // This sync rule may be present only in the Pilot config OR in the Pilot as well as Production config.
                        var connectorProduction = this.ProductionXml.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(false) + "/ma-data[name ='" + this.ConnectorName + "']");
                        if (connectorProduction == null)
                        {
                            // This sync rule is present only in the Pilot config. Give a warning.
                            script = this.CreateMissingDefaultSyncRuleWarningScript(syncRule, false);
                        }
                        else
                        {
                            var connectorGuidProduction = ((string)connectorProduction.Element("id") ?? string.Empty).ToUpperInvariant();
                            var syncRuleProduction = this.ProductionXml.XPathSelectElement(this.GetSyncRuleXPath(connectorGuidProduction, false));
                            if (syncRuleProduction == null)
                            {
                                // This sync rule is present only in the Pilot config. Give a warning.
                                script = this.CreateMissingDefaultSyncRuleWarningScript(syncRule, false);
                            }
                            else
                            {
                                var disabledPilot = (string)syncRule.Element("disabled");
                                var disabledProduction = (string)syncRuleProduction.Element("disabled");

                                if (disabledPilot != disabledProduction)
                                {
                                    var disable = string.IsNullOrEmpty(disabledPilot) ? false : Convert.ToBoolean(disabledPilot, CultureInfo.InvariantCulture);

                                    script = this.CreateDefaultSyncRuleUpdateScript(syncRuleProduction, disable);
                                }
                                else
                                {
                                    var precedencePilot = (string)syncRule.Element("precedence");
                                    var precedenceProduction = (string)syncRuleProduction.Element("precedence");
                                    script = this.CreateDefaultSyncRuleUpdateWarningScript(syncRuleProduction, precedencePilot != precedenceProduction);
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (this.Environment == ConfigEnvironment.ProductionOnly)
                    {
                        script = this.CreateRemoveADSyncRuleScript(syncRule);
                    }
                    else
                    {
                        script = this.CreateNewADSyncRuleScript(syncRule);
                    }
                }

                #region div

                this.ReportWriter.WriteBeginTag("div");
                this.ReportWriter.WriteAttribute("class", "PowerShellScript");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportWriter.WriteLine(script);
                this.ReportWriter.WriteEndTag("div");

                #endregion div
            }
            catch (Exception e)
            {
                Logger.Instance.WriteError(e.ToString());

                try
                {
                    #region div

                    var syncRuleScript = new StringBuilder();

                    syncRuleScript.AppendFormat(CultureInfo.InvariantCulture, Documenter.GetEmbeddedScriptResource("PowerShellScriptSectionHeader.ps1"), this.ConnectorName, this.SyncRuleName);
                    syncRuleScript.AppendLine();
                    syncRuleScript.AppendLine();
                    syncRuleScript.AppendFormat("$connectorName = '{0}'", this.ConnectorName);
                    syncRuleScript.AppendLine();
                    syncRuleScript.AppendFormat("$syncRuleName = '{0}'", this.SyncRuleName);
                    syncRuleScript.AppendLine();
                    syncRuleScript.AppendLine("$errMsg = @'");
                    syncRuleScript.AppendLine(e.ToString());
                    syncRuleScript.AppendLine("'@");
                    syncRuleScript.AppendLine("Write-Error $errMsg");

                    this.ReportWriter.WriteBeginTag("div");
                    this.ReportWriter.WriteAttribute("class", "PowerShellScript");
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.WriteLine(syncRuleScript.ToString());
                    this.ReportWriter.WriteEndTag("div");

                    #endregion div
                }
                catch (Exception)
                {
                    // Do nothing. Already reported
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates a Warning Script if unsupported changes are made to a default sync rule
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <param name="productionConfig">If set to <c>true</c>, indicates the sync rule is missing from production config. Otherwise from pilot config.</param>
        /// <returns>The Synchronization Rule script.</returns>
        private string CreateMissingDefaultSyncRuleWarningScript(XElement syncRule, bool productionConfig)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                var syncRuleName = (string)syncRule.Element("name");

                syncRuleScript.AppendFormat(CultureInfo.InvariantCulture, Documenter.GetEmbeddedScriptResource("PowerShellScriptSectionHeader.ps1"), this.ConnectorName, syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$connectorName = '{0}'", this.ConnectorName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$syncRuleName = '{0}'", syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("Write-Warning(\"The sync rule '{0}' for the connector '{1}' only exists in the config supplied as the `\"" + (productionConfig ? "production" : "pilot") + "`\" config.\" -f $syncRuleName, $connectorName)");
                syncRuleScript.AppendLine("Write-Warning (\"This sync rule is inferred as a part of the out-of-box default rule set and will be skipped.\")");
                syncRuleScript.AppendLine("Write-Warning (\"This may be due to different versions or feature set selection between the production and pilot config.\")");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        /// <summary>
        /// Creates a Warning Script if unsupported changes are made to a default sync rule
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <param name="precedenceChange">Indicates if the precedence has been changed between the Pilot and Production.</param>
        /// <returns>The Synchronization Rule script.</returns>
        private string CreateDefaultSyncRuleUpdateWarningScript(XElement syncRule, bool precedenceChange)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                var syncRuleName = (string)syncRule.Element("name");

                syncRuleScript.AppendFormat(CultureInfo.InvariantCulture, Documenter.GetEmbeddedScriptResource("PowerShellScriptSectionHeader.ps1"), this.ConnectorName, syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$connectorName = '{0}'", this.ConnectorName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$syncRuleName = '{0}'", syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("Write-Warning(\"The sync rule '{0}' for the connector '{1}' has unsupported chanages detected.\" -f $syncRuleName, $connectorName)");
                syncRuleScript.AppendLine("Write-Warning (\"Only supported change to an out-of-box default rule is to make it `\"Disabled`\".\")");
                if (precedenceChange)
                {
                    syncRuleScript.AppendLine("Write-Warning (\"If only the precedence number is different for this out-of-box rule, this warning may be safely ignored.\")");
                }

                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        /// <summary>
        /// Creates a Update Script for supported changes made to a default sync rule
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <param name="disable">Indicates if the Synchronization Rule is to be disabled.</param>
        /// <returns>The Synchronization Rule script.</returns>
        private string CreateDefaultSyncRuleUpdateScript(XElement syncRule, bool disable)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                var syncRuleName = (string)syncRule.Element("name");

                syncRuleScript.AppendFormat(CultureInfo.InvariantCulture, Documenter.GetEmbeddedScriptResource("PowerShellScriptSectionHeader.ps1"), this.ConnectorName, syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$connectorName = '{0}'", this.ConnectorName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$syncRuleName = '{0}'", syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("Write-Host \"Processing Sync Rule '$syncRuleName' for Connector '$connectorName'\"");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("$connectorId = [string](Get-ADSyncConnector -Name $connectorName).Identifier");
                syncRuleScript.AppendLine("$syncRule = Get-ADSyncRule | Where { $_.Name -eq $syncRuleName -and $_.Connector -eq  $connectorId }");
                syncRuleScript.AppendLine("if ($syncRule.Count -eq 1)");
                syncRuleScript.AppendLine("{");
                syncRuleScript.AppendFormat("    $syncRule.Disabled =  ${0}", disable);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("    $syncRule | Add-ADSyncRule | Out-Null");
                syncRuleScript.AppendLine("}");
                syncRuleScript.AppendLine("elseif ($syncRule.Count -gt 1)");
                syncRuleScript.AppendLine("{");
                syncRuleScript.AppendLine("    Write-Error \"Error processing Sync Rule '$syncRuleName' for Connector '$connectorName'. More than one sync rules found.\"");
                syncRuleScript.AppendLine("}");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        /// <summary>
        /// Creates Remove-ADSyncRule Script
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <returns>The Synchronization Rule script.</returns>
        private string CreateRemoveADSyncRuleScript(XElement syncRule)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                var syncRuleName = (string)syncRule.Element("name");

                syncRuleScript.AppendFormat(CultureInfo.InvariantCulture, Documenter.GetEmbeddedScriptResource("PowerShellScriptSectionHeader.ps1"), this.ConnectorName, syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$connectorName = '{0}'", this.ConnectorName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$syncRuleName = '{0}'", syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("Write-Host \"Processing Sync Rule '$syncRuleName' for Connector '$connectorName'\"");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("$connectorId = [string](Get-ADSyncConnector -Name $connectorName).Identifier");
                syncRuleScript.AppendLine("$syncRule = Get-ADSyncRule | Where { $_.Name -eq $syncRuleName -and $_.Connector -eq  $connectorId }");
                syncRuleScript.AppendLine("if ($syncRule.Count -eq 1)");
                syncRuleScript.AppendLine("{");
                syncRuleScript.AppendLine("    $syncRule | Remove-ADSyncRule | Out-Null");
                syncRuleScript.AppendLine("}");
                syncRuleScript.AppendLine("elseif ($syncRule.Count -gt 1)");
                syncRuleScript.AppendLine("{");
                syncRuleScript.AppendLine("    Write-Error \"Error processing Sync Rule '$syncRuleName' for Connector '$connectorName'. More than one sync rules found.\"");
                syncRuleScript.AppendLine("}");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        /// <summary>
        /// Creates New-ADSyncRule Script
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <returns>The Synchronization Rule script.</returns>
        private string CreateNewADSyncRuleScript(XElement syncRule)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                var syncRuleName = (string)syncRule.Element("name");
                var syncRuleId = ((string)syncRule.Element("id")).TrimStart('{').TrimEnd('}');
                var syncRulePrecedence = (string)syncRule.Element("precedence");

                syncRuleScript.AppendFormat(CultureInfo.InvariantCulture, Documenter.GetEmbeddedScriptResource("PowerShellScriptSectionHeader.ps1"), this.ConnectorName, syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$connectorName = '{0}'", this.ConnectorName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$syncRuleName = '{0}'", syncRuleName);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$syncRuleId = '{0}'", syncRuleId);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendFormat("$syncRulePrecedence = {0}", syncRulePrecedence);
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("Write-Host \"Processing Sync Rule '$syncRuleName' for Connector '$connectorName'\"");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("$connectorId = [string](Get-ADSyncConnector -Name $connectorName).Identifier");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine("# First clean-up any existing old rule in case it was created using editor UI and hence has a different guid.");
                syncRuleScript.AppendLine("$syncRule = Get-ADSyncRule | Where { $_.Name -eq $syncRuleName -and $_.Connector -eq  $connectorId }");
                syncRuleScript.AppendLine("if ($syncRule.Count -eq 1)");
                syncRuleScript.AppendLine("{");
                syncRuleScript.AppendLine("    $syncRule | Remove-ADSyncRule | Out-Null");
                syncRuleScript.AppendLine("}");
                syncRuleScript.AppendLine("elseif ($syncRule.Count -gt 1)");
                syncRuleScript.AppendLine("{");
                syncRuleScript.AppendLine("    Write-Error \"Error processing Sync Rule '$syncRuleName' for Connector '$connectorName'. More than one sync rules found.\"");
                syncRuleScript.AppendLine("}");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();

                syncRuleScript.AppendLine("New-ADSyncRule `");

                syncRuleScript.AppendLine("-Name $syncRuleName `");

                syncRuleScript.AppendLine("-Identifier $syncRuleId `");

                syncRuleScript.AppendLine("-Description @'");
                syncRuleScript.AppendLine((string)syncRule.Element("description"));
                syncRuleScript.AppendLine("'@ `");

                syncRuleScript.AppendFormat("-Direction '{0}' `", (string)syncRule.Element("direction"));
                syncRuleScript.AppendLine();

                syncRuleScript.AppendLine("-Precedence $syncRulePrecedence `");

                syncRuleScript.AppendFormat("-PrecedenceAfter '{0}' `", (string)syncRule.Element("precedence-after"));
                syncRuleScript.AppendLine();

                syncRuleScript.AppendFormat("-PrecedenceBefore '{0}' `", (string)syncRule.Element("precedence-before"));
                syncRuleScript.AppendLine();

                syncRuleScript.AppendFormat("-SourceObjectType '{0}' `", (string)syncRule.Element("sourceObjectType"));
                syncRuleScript.AppendLine();

                syncRuleScript.AppendFormat("-TargetObjectType '{0}' `", (string)syncRule.Element("targetObjectType"));
                syncRuleScript.AppendLine();

                syncRuleScript.AppendLine("-Connector $connectorId `");

                syncRuleScript.AppendFormat("-LinkType '{0}' `", (string)syncRule.Element("linkType"));
                syncRuleScript.AppendLine();

                syncRuleScript.AppendFormat("-SoftDeleteExpiryInterval '{0}' `", (string)syncRule.Element("softDeleteExpiryInterval"));
                syncRuleScript.AppendLine();

                syncRuleScript.AppendFormat("-ImmutableTag '{0}' `", (string)syncRule.Element("immutable-tag"));
                syncRuleScript.AppendLine();

                var enablePasswordSync = (string)syncRule.Element("EnablePasswordSync");
                if (!string.IsNullOrEmpty(enablePasswordSync) && enablePasswordSync.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    syncRuleScript.AppendLine("-EnablePasswordSync `");
                }

                var disabled = (string)syncRule.Element("disabled");
                if (!string.IsNullOrEmpty(disabled) && disabled.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    syncRuleScript.AppendLine("-Disabled `");
                }

                syncRuleScript.AppendLine("-OutVariable syncRule | Out-Null");
                syncRuleScript.AppendLine();

                var addADSyncAttributeFlowMapping = this.CreateAddADSyncAttributeFlowMappingScript(syncRule);
                var addADSyncADSyncScopeConditionGroup = this.CreateAddADSyncADSyncScopeConditionGroupScript(syncRule);
                var addADSyncADSyncJoinConditionGroup = this.CreateAddADSyncADSyncJoinConditionGroupScript(syncRule);

                syncRuleScript.AppendLine(addADSyncAttributeFlowMapping);
                syncRuleScript.AppendLine(addADSyncADSyncScopeConditionGroup);
                syncRuleScript.AppendLine(addADSyncADSyncJoinConditionGroup);

                syncRuleScript.AppendLine("Add-ADSyncRule -SynchronizationRule $syncRule[0] | Out-Null");
                syncRuleScript.AppendLine();
                syncRuleScript.AppendLine();

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        /// <summary>
        /// Creates Add-ADSyncAttributeFlowMapping Script
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <returns>The Synchronization Rule Attribute Flow Mapping script.</returns>
        private string CreateAddADSyncAttributeFlowMappingScript(XElement syncRule)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                var transformations = from transformation in syncRule.XPathSelectElements("attribute-mappings/mapping")
                                      let target = (string)transformation.Element("dest")
                                      orderby target
                                      select transformation;

                foreach (var transformation in transformations)
                {
                    var targetAttribute = (string)transformation.Element("dest");
                    var expression = (string)transformation.Element("expression");
                    var valueMergeType = (string)transformation.Element("valueMergeType");
                    var applyOnce = (string)transformation.Attribute("execute-once");
                    var sourceAttributes = transformation.XPathSelectElements("src/attr");
                    string flowType = !string.IsNullOrEmpty(expression) ? "Expression" : sourceAttributes.Any() ? "Direct" : "Constant";

                    syncRuleScript.AppendLine("Add-ADSyncAttributeFlowMapping `");
                    syncRuleScript.AppendLine("-SynchronizationRule $syncRule[0] `");

                    var source = string.Empty;
                    switch (flowType)
                    {
                        case "Expression":
                            {
                                foreach (var sourceAttribute in sourceAttributes)
                                {
                                    source += "'" + (string)sourceAttribute + "',";
                                }

                                source = source.TrimEnd(',');
                            }

                            break;

                        case "Direct":
                        case "Constant":
                            {
                                source = "'" + (string)transformation.Element("src") + "'";
                            }

                            break;
                    }

                    // Source may be null e.g. RemoveDuplicates(Trim(ImportedValue("proxyAddresses")))
                    if (!string.IsNullOrEmpty(source))
                    {
                        syncRuleScript.AppendFormat("-Source @({0}) `", source);
                        syncRuleScript.AppendLine();
                    }

                    syncRuleScript.AppendFormat("-Destination '{0}' `", targetAttribute);
                    syncRuleScript.AppendLine();
                    syncRuleScript.AppendFormat("-FlowType '{0}' `", flowType);
                    syncRuleScript.AppendLine();
                    syncRuleScript.AppendFormat("-ValueMergeType '{0}' `", valueMergeType);
                    syncRuleScript.AppendLine();

                    if (!string.IsNullOrEmpty(applyOnce) && applyOnce.Equals("true", StringComparison.OrdinalIgnoreCase))
                    {
                        syncRuleScript.AppendLine("-ExecuteOnce `");
                    }

                    if (!string.IsNullOrEmpty(expression))
                    {
                        syncRuleScript.AppendFormat("-Expression '{0}' `", expression);
                        syncRuleScript.AppendLine();
                    }

                    syncRuleScript.AppendLine("-OutVariable syncRule | Out-Null");
                    syncRuleScript.AppendLine();
                }

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        /// <summary>
        /// Creates Add-ADSyncScopeConditionGroup Script
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <returns>The Synchronization Rule ScopeCondition Group script.</returns>
        private string CreateAddADSyncADSyncScopeConditionGroupScript(XElement syncRule)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                foreach (var condition in syncRule.XPathSelectElements("synchronizationCriteria/conditions"))
                {
                    var scopes = condition.Elements("scope");

                    var conditionVariables = "-ScopeConditions @(";
                    var scopeIndex = -1;
                    foreach (var scope in scopes)
                    {
                        ++scopeIndex;
                        var scopeAttribute = (string)scope.Element("csAttribute");
                        var scopeOperator = (string)scope.Element("csOperator");
                        var scopeValue = (string)scope.Element("csValue");

                        conditionVariables += "$condition" + scopeIndex + "[0],";

                        syncRuleScript.AppendLine("New-Object `");
                        syncRuleScript.AppendLine("-TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `");
                        syncRuleScript.AppendFormat("-ArgumentList '{0}', '{1}', '{2}' `", scopeAttribute, scopeValue, scopeOperator);
                        syncRuleScript.AppendLine();
                        syncRuleScript.AppendFormat("-OutVariable condition{0} | Out-Null", scopeIndex);
                        syncRuleScript.AppendLine();
                        syncRuleScript.AppendLine();
                    }

                    conditionVariables = conditionVariables.TrimEnd(',') + ") `";

                    syncRuleScript.AppendLine("Add-ADSyncScopeConditionGroup `");
                    syncRuleScript.AppendLine("-SynchronizationRule $syncRule[0] `");
                    syncRuleScript.AppendLine(conditionVariables);
                    syncRuleScript.AppendLine("-OutVariable syncRule | Out-Null");
                    syncRuleScript.AppendLine();
                }

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        /// <summary>
        /// Creates Add-ADSyncJoinConditionGroup Script
        /// </summary>
        /// <param name="syncRule">The Synchronization Rule to be scripted.</param>
        /// <returns>The Synchronization Rule JoinCondition Group script.</returns>
        private string CreateAddADSyncADSyncJoinConditionGroupScript(XElement syncRule)
        {
            Logger.Instance.WriteMethodEntry("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));

            var syncRuleScript = new StringBuilder();

            try
            {
                this.syncRuleDirection = (SyncRuleDirection)Enum.Parse(typeof(SyncRuleDirection), (string)syncRule.Element("direction"), true);

                foreach (var joinRule in syncRule.XPathSelectElements("relationshipCriteria/conditions"))
                {
                    var conditions = joinRule.Elements("condition");

                    var conditionVariables = "-JoinConditions @(";
                    var conditionIndex = -1;
                    foreach (var condition in conditions)
                    {
                        ++conditionIndex;
                        var csAttribute = (string)condition.Element("csAttribute");
                        var mvAttribute = (string)condition.Element("ilmAttribute");
                        var caseSensitive = (string)condition.Element("caseSensitive");
                        caseSensitive = string.IsNullOrEmpty(caseSensitive) || caseSensitive.Equals("False", StringComparison.OrdinalIgnoreCase) ? "$false" : "$true";

                        conditionVariables += "$condition" + conditionIndex + "[0],";

                        syncRuleScript.AppendLine("New-Object `");
                        syncRuleScript.AppendLine("-TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.JoinCondition' `");

                        syncRuleScript.AppendFormat("-ArgumentList '{0}', '{1}', {2} `", csAttribute, mvAttribute, caseSensitive);

                        syncRuleScript.AppendLine();
                        syncRuleScript.AppendFormat("-OutVariable condition{0} | Out-Null", conditionIndex);
                        syncRuleScript.AppendLine();
                        syncRuleScript.AppendLine();
                    }

                    conditionVariables = conditionVariables.TrimEnd(',') + ") `";

                    syncRuleScript.AppendLine("Add-ADSyncJoinConditionGroup `");
                    syncRuleScript.AppendLine("-SynchronizationRule $syncRule[0] `");
                    syncRuleScript.AppendLine(conditionVariables);
                    syncRuleScript.AppendLine("-OutVariable syncRule | Out-Null");
                    syncRuleScript.AppendLine();
                }

                return syncRuleScript.ToString();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("SycRule: '{0}', Id: '{1}'.", (string)syncRule.Element("name"), (string)syncRule.Element("id"));
            }
        }

        #endregion Sync Rule Script Generation
    }
}
