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
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Linq;
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
        /// The synchronize rule direction
        /// </summary>
        private SyncRuleDirection syncRuleDirection;

        /// <summary>
        /// The synchronize rule report type
        /// </summary>
        private SyncRuleReportType syncRuleReportType;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncRuleDocumenter" /> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="syncRuleName">Name of the synchronize rule.</param>
        /// <param name="syncRuleGuid">The synchronize rule unique identifier.</param>
        /// <param name="connectorName">The connector name.</param>
        /// <param name="productionOnly">if set to <c>true</c> [production only].</param>
        public SyncRuleDocumenter(XElement pilotXml, XElement productionXml, string syncRuleName, string syncRuleGuid, string connectorName, bool productionOnly)
            : base(pilotXml, productionXml, connectorName, productionOnly)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Name: '{0}'. Sync Rule Name: '{1}'.", syncRuleName, syncRuleGuid);

            try
            {
                this.SyncRuleName = syncRuleName;
                this.SyncRuleGuid = syncRuleGuid;
                this.ReportFileName = Documenter.GetTempFilePath(this.ConnectorName + "_" + this.SyncRuleName + "_" + this.SyncRuleGuid + ".tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath(this.ConnectorName + "_" + this.SyncRuleName + "_" + this.SyncRuleGuid + ".TOC.tmp.html");

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

                var noHide = this.DiffgramDataSets.Any(dataSet => !(bool)dataSet.ExtendedProperties[Documenter.CanHide]);

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
                
                try
                {
                    this.WriteSyncRuleReportHeader();
                    this.PrintConnectorSyncRuleDescription();
                    this.PrintConnectorSyncRuleScopingFilter();
                    this.PrintConnectorSyncRuleJoinRules();
                    if (reportType == SyncRuleReportType.AllSections)
                    {
                        this.PrintConnectorSyncRuleTransformations();
                    }
                }
                finally
                {
                    this.ResetDiffgram(); // reset the diffgram variables
                }

                return base.GetReport();
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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        protected void WriteSyncRuleReportHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var bookmark = this.SyncRuleName;
                if (this.syncRuleReportType != SyncRuleReportType.AllSections)
                {
                    bookmark += this.syncRuleReportType.ToString();
                    this.ReportFileName = Documenter.GetTempFilePath(this.syncRuleReportType + "." + this.ConnectorName + "_" + this.SyncRuleName + "_" + this.SyncRuleGuid + ".tmp.html");
                    this.ReportToCFileName = Documenter.GetTempFilePath(this.syncRuleReportType + "." + this.ConnectorName + "_" + this.SyncRuleName + "_" + this.SyncRuleGuid + ".TOC.tmp.html");
                }

                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc5" + " " + this.GetCssVisibilityClass());
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, bookmark, this.SyncRuleName, this.SyncRuleGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.WriteAttribute("class", " " + this.GetCssVisibilityClass());
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteBeginTag("h5");
                this.ReportWriter.WriteAttribute("class", this.GetCssVisibilityClass());
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteBookmarkLocation(this.ReportWriter, bookmark, this.SyncRuleName, this.SyncRuleGuid, "TOC");
                this.ReportWriter.WriteEndTag("h5");

                #endregion section
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
        /// <returns>
        /// The sync rule xpath.
        /// </returns>
        protected string GetSyncRuleXPath(string currentConnectorGuid)
        {
            Logger.Instance.WriteMethodEntry("Current Connector Guid: '{0}'.", currentConnectorGuid);

            var xpath = "//synchronizationRule[translate(connector, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + currentConnectorGuid + "' and name = '" + this.SyncRuleName + "'";

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
                Logger.Instance.WriteMethodExit("Current Connector Guid: '{0}'. XPath: '{1}'.", currentConnectorGuid, xpath);
            }
        }

        #region Sync Rule Description

        /// <summary>
        /// Processes the connector synchronize rule description.
        /// </summary>
        protected void ProcessConnectorSyncRuleDescription()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Sync Rule Description.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order, 2 = Field Name, 3 = Value

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
        protected void FillConnectorSyncRuleDescriptionDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var currentConnectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant(); // This may be pilot or production GUID

                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid);

                    var syncRule = config.XPathSelectElement(xpath);

                    if (syncRule != null)
                    {
                        var table = dataSet.Tables[0];

                        this.syncRuleDirection = (SyncRuleDirection)Enum.Parse(typeof(SyncRuleDirection), (string)syncRule.Element("direction"), true);

                        var setting = (string)syncRule.Element("name");
                        Documenter.AddRow(table, new object[] { 0, "Name", setting });

                        if (this.syncRuleReportType == SyncRuleReportType.AllSections)
                        {
                            setting = (string)syncRule.Element("description");
                            Documenter.AddRow(table, new object[] { 1, "Description", setting });
                        }

                        setting = (string)syncRule.Element("direction");
                        Documenter.AddRow(table, new object[] { 2, "Direction", setting });

                        if (this.syncRuleReportType == SyncRuleReportType.AllSections)
                        {
                            setting = this.ConnectorName;
                            Documenter.AddRow(table, new object[] { 3, "Connected System", setting });
                        }

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
        /// Prints the connector synchronize rule description.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintConnectorSyncRuleDescription()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
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
                            this.ReportWriter.WriteAttribute("colspan", "2");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Description");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteEndTag("tr");
                            this.ReportWriter.WriteLine();

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
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.DiffgramDataSet = this.DiffgramDataSets[0];
                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
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
        protected void ProcessConnectorSyncRuleScopingFilter()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Sync Rule Scoping Filter.");

                this.CreateSimpleOrderedSettingsDataSets(5, 5); // 1 = Display Order, 2 = Group #, 3 = Attribute, 4 = Operator, 5 = Value

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
        protected void FillConnectorSyncRuleScopingFilterDataSet(bool pilotConfig)
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
                    var currentConnectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant(); // This may be pilot or production GUID
                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid) + "/synchronizationCriteria/conditions";

                    var syncRuleScopingConditions = config.XPathSelectElements(xpath);
                    var conditionCount = syncRuleScopingConditions.Count();

                    if (conditionCount == 0)
                    {
                        Documenter.AddRow(table, new object[] { 0, "-", "-", "-", "-" }, true);
                    }
                    else
                    {
                        for (var conditionIndex = 0; conditionIndex < conditionCount; ++conditionIndex)
                        {
                            var condition = syncRuleScopingConditions.ElementAt(conditionIndex);
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
        /// Prints the connector synchronize rule scoping filter.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintConnectorSyncRuleScopingFilter()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
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
                            this.ReportWriter.WriteAttribute("colspan", "4");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Scoping Filter");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteEndTag("tr");
                            this.ReportWriter.WriteLine();

                            this.ReportWriter.WriteBeginTag("tr");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            {
                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Group#");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Attribute");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Operator");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Value");
                                this.ReportWriter.WriteEndTag("th");
                            }
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.DiffgramDataSet = this.DiffgramDataSets[1];
                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
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
        protected void ProcessConnectorSyncRuleJoinRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Sync Rule Join Rules");

                this.CreateSimpleOrderedSettingsDataSets(5, 5); // 1 = Display Order, 2 = Group #, 3 = Source Attribute, 4 = Target Attribute, 5 = Case Sensitive?

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
        protected void FillConnectorSyncRuleJoinRulesDataSet(bool pilotConfig)
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
                    var currentConnectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant(); // This may be pilot or production GUID
                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid);
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

                        for (var joinRuleIndex = 0; joinRuleIndex < joinRuleCount; ++joinRuleIndex)
                        {
                            var joinRule = syncRuleJoiningRules.ElementAt(joinRuleIndex);
                            var conditions = joinRule.Elements("condition");
                            foreach (var condition in conditions)
                            {
                                var scopeAttribute = (string)condition.Element("csAttribute") ?? " ";
                                var scopeOperator = (string)condition.Element("ilmAttribute") ?? " ";
                                var scopeValue = (string)condition.Element("caseSensitive") ?? " ";
                                if (this.syncRuleDirection == SyncRuleDirection.Inbound)
                                {
                                    Documenter.AddRow(table, new object[] { joinRuleIndex + 1, joinRuleIndex + 1, scopeAttribute, scopeOperator, scopeValue });
                                }
                                else
                                {
                                    Documenter.AddRow(table, new object[] { joinRuleIndex + 1, joinRuleIndex + 1, scopeOperator, scopeAttribute, scopeValue });
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
        /// Prints the connector synchronize rule join rules.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintConnectorSyncRuleJoinRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
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
                            this.ReportWriter.WriteAttribute("colspan", "4");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Join Rules");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteEndTag("tr");
                            this.ReportWriter.WriteLine();

                            this.ReportWriter.WriteBeginTag("tr");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            {
                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Group#");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Source Attribute");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Target Attribute");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Case Sensitive");
                                this.ReportWriter.WriteEndTag("th");
                            }
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.DiffgramDataSet = this.DiffgramDataSets[2];
                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
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
        protected void ProcessConnectorSyncRuleTransformations()
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
        protected void FillConnectorSyncRuleTransformationsDataSet(bool pilotConfig)
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
                    var currentConnectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant(); // This may be pilot or production GUID
                    var xpath = this.GetSyncRuleXPath(currentConnectorGuid) + "/attribute-mappings/mapping";

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
        /// Prints the connector synchronize rule transformations.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintConnectorSyncRuleTransformations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
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
                            this.ReportWriter.WriteAttribute("colspan", "5");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Transformations");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteEndTag("tr");
                            this.ReportWriter.WriteLine();

                            this.ReportWriter.WriteBeginTag("tr");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            {
                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write(this.syncRuleDirection == SyncRuleDirection.Inbound ? "Target (MV) Attribute" : "Target (CS) Attribute");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Source");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Flow Type");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Apply Once");
                                this.ReportWriter.WriteEndTag("th");

                                this.ReportWriter.WriteBeginTag("th");
                                this.ReportWriter.WriteAttribute("class", "column-th");
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.ReportWriter.Write("Merge Type");
                                this.ReportWriter.WriteEndTag("th");
                            }
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.DiffgramDataSet = this.DiffgramDataSets[3];
                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");

                #endregion table
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sync Rule Transformations
    }
}
