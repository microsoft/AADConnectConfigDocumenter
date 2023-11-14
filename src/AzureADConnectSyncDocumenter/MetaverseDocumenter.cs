//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MetaverseDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync Metaverse Configuration Documenter
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
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The MetaverseDocumenter documents the configuration of Metaverse.
    /// </summary>
    internal class MetaverseDocumenter : Documenter
    {
        /// <summary>
        /// The logger context item metaverse object type
        /// </summary>
        private const string LoggerContextItemMetaverseObjectType = "Metaverse Object Type";

        /// <summary>
        /// The logger context item metaverse attribute
        /// </summary>
        private const string LoggerContextItemMetaverseAttribute = "Metaverse Attribute";

        /// <summary>
        /// The object type currently being processed
        /// </summary>
        private string currentObjectType;

        /// <summary>
        /// Initializes a new instance of the <see cref="MetaverseDocumenter"/> class.
        /// </summary>
        /// <param name="pilotConfigXml">The pilot configuration XML.</param>
        /// <param name="productionConfigXml">The production configuration XML.</param>
        public MetaverseDocumenter(XElement pilotConfigXml, XElement productionConfigXml)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PilotXml = pilotConfigXml;
                this.ProductionXml = productionConfigXml;

                this.ReportFileName = Documenter.GetTempFilePath("MV.tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath("MV.TOC.tmp.html");
                this.SyncRuleChangesScriptFileName = Documenter.GetTempFilePath("MV.tmp.ps1");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the metaverse configuration report.
        /// </summary>
        /// <returns>The Tuple of configuration report and associated TOC</returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        public override Tuple<string, string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));
                this.SyncRuleChangesScriptWriter = new StreamWriter(this.SyncRuleChangesScriptFileName);

                var sectionTitle = "Metaverse Configuration";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 2);

                this.ProcessMetaverseObjectTypes();
                this.ProcessMetaverseObjectDeletionRules();

                return this.GetReportTuple();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Metaverse ObjectType

        /// <summary>
        /// Processes the metaverse object types.
        /// </summary>
        private void ProcessMetaverseObjectTypes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Metaverse Object Types";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3);

                var configSetting = ConfigurationManager.AppSettings["SuppressMetaverseObjectTypeConfigSection"];
                if (!string.IsNullOrEmpty(configSetting) && configSetting.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Instance.WriteWarning(string.Format(CultureInfo.InvariantCulture, "!!WARNING!! SuppressMetaverseObjectTypeConfigSection = '{0}'.", configSetting));
                    this.WriteContentParagraph("The documentation of this section is suppressed via config setting.", "Highlight");
                    return;
                }

                var pilot = this.PilotXml.XPathSelectElements(Documenter.GetMetaverseXmlRootXPath(true) + "/mv-data//dsml:class", Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(Documenter.GetMetaverseXmlRootXPath(false) + "/mv-data//dsml:class", Documenter.NamespaceManager);

                // Sort by name
                var pilotObjectTypes = from objectType in pilot
                                       let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                       orderby name
                                       select name;

                foreach (var objectType in pilotObjectTypes)
                {
                    this.currentObjectType = objectType;
                    this.ProcessMetaverseObjectType();
                }

                // Sort by name
                var productionObjectTypes = from objectType in production
                                            let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                            orderby name
                                            select name;

                productionObjectTypes = productionObjectTypes.Where(productionObjectType => !pilotObjectTypes.Contains(productionObjectType));

                foreach (var objectType in productionObjectTypes)
                {
                    this.currentObjectType = objectType;
                    this.ProcessMetaverseObjectType();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current metaverse object type.
        /// </summary>
        private void ProcessMetaverseObjectType()
        {
            Logger.Instance.WriteMethodEntry("Metaverse Object Type: {0}.", this.currentObjectType);

            try
            {
                // Set Logger call context items
                Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType, this.currentObjectType);

                Logger.Instance.WriteInfo("Processing Metaverse Object Type.");

                this.CreateMetaverseObjectTypeDataSets();

                this.FillMetaverseObjectTypeDataSet(true);
                this.FillMetaverseObjectTypeDataSet(false);

                this.CreateMetaverseObjectTypeDiffgram();

                this.PrintMetaverseObjectType();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType);
            }
        }

        /// <summary>
        /// Creates the metaverse object type data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        private void CreateMetaverseObjectTypeDataSets()
        {
            Logger.Instance.WriteMethodEntry();
            //columns and order
            //Precedence, Connector, InboundSyncRule, inboundFlowType, Source, inboundSyncRuleScopingConditionString, Attribute, Type, Multivalued, Indexed
            try
            {
                var table = new DataTable("customizedMetaverseObjectType") { Locale = CultureInfo.InvariantCulture };
                var column1 = new DataColumn("Precedence", typeof(int));
                var column2 = new DataColumn("Connector");
                var column3= new DataColumn("InboundSyncRule");
                var column4 = new DataColumn("inboundFlowType");
                var column5 = new DataColumn("Source");
                var column6 = new DataColumn("inboundSyncRuleScopingConditionString");
                var column7 = new DataColumn("metaverseAttribute");
                var column8 = new DataColumn("metaverseObjectType");
                var column9 = new DataColumn("Multivalued");
                var column10 = new DataColumn("Indexed");

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
                table.PrimaryKey = new[] { column1, column2, column3, column4, column5, column6, column7, column8, column9, column10 };

                this.PilotDataSet = new DataSet("MetaverseObjectType") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetMetaverseObjectTypePrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the metaverse object type print table.
        /// </summary>
        /// <returns>The metaverse object type print table.</returns>
        private DataTable GetMetaverseObjectTypePrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();
                //columns and order
                //Precedence, Connector, InboundSyncRule, inboundFlowType, Source, inboundSyncRuleScopingConditionString, metaverseAttribute, metaverseObjectType, Multivalued, Indexed

                // Table 1
                // Precedence
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 5 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // InboundSyncRule
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 6 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // inboundFlowType
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Source
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // inboundSyncRuleScopingConditionString
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 5 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // metaverseAttribute
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 6 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // metaverseObjectType
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 7 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Multi-valued
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 8 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Indexed
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 9 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the metaverse object type data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillMetaverseObjectTypeDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var attributes = config.XPathSelectElements(Documenter.GetMetaverseXmlRootXPath(pilotConfig) + "/mv-data//dsml:class[dsml:name = '" + this.currentObjectType + "' ]/dsml:attribute", Documenter.NamespaceManager);

                // Sort by name
                attributes = from attribute in attributes
                             let name = (string)attribute.Attribute("ref")
                             orderby name
                             select attribute;
                var attributeIndex = -1;

                string metaverseObjectType = this.currentObjectType;

                foreach (var attribute in attributes)
                {
                    ++attributeIndex;
                    var metaverseAttribute = ((string)attribute.Attribute("ref") ?? string.Empty).Trim('#');

                    // Set Logger call context items
                    Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute, metaverseAttribute);

                    Logger.Instance.WriteInfo("Processing Attribute Information.");

                    var attributeInfo = config.XPathSelectElement(Documenter.GetMetaverseXmlRootXPath(pilotConfig) + "/mv-data//dsml:attribute-type[dsml:name = '" + metaverseAttribute + "']", Documenter.NamespaceManager);

                    var attributeSyntax = (string)attributeInfo.Element(Documenter.DsmlNamespace + "syntax");

                    var row = table.NewRow();

                    string type = Documenter.GetAttributeType(attributeSyntax, (string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexable"));
                    string singleValue = ((string)attributeInfo.Attribute("single-value") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";
                    string indexed = ((string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexed") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";

                    Logger.Instance.WriteVerbose("Processed Attribute Information.");

                    // Fetch Sync Rules
                    var syncRules = config.XPathSelectElements(Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[targetObjectType = '" + this.currentObjectType + "' and direction = 'Inbound' " + Documenter.SyncRuleDisabledCondition + " and attribute-mappings/mapping/dest = '" + metaverseAttribute + "']");
                    syncRules = from syncRule in syncRules
                                let precedence = (int)syncRule.Element("precedence")
                                orderby precedence
                                select syncRule;

                    var syncRuleIndex = -1;
                    foreach (var syncRule in syncRules)
                    {
                        string inboundSyncRuleScopingConditionString = "";
                        ++syncRuleIndex;
                        int rank = syncRuleIndex + 1; // Care only about the precedence relative rank here than actual value
                        var connector = ((string)syncRule.Element("connector") ?? string.Empty).ToUpperInvariant();

                        var connectorName = (string)config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connector + "']/name");
                        if (string.IsNullOrEmpty(connectorName))
                        {
                            Logger.Instance.WriteWarning(string.Format(CultureInfo.InvariantCulture, "Unable to dereferece connector: '{0}'. The documentation of inbound flows will be skipped for this connector. PilotConfig: '{1}'.", connector, pilotConfig));
                            continue;
                        }

                        var inboundSyncRuleName = (string)syncRule.Element("name");

                        Logger.Instance.WriteVerbose("Processing Sync Rule Info for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, inboundSyncRuleName);

                        var mappingExpression = (string)syncRule.XPathSelectElement("attribute-mappings/mapping[dest = '" + metaverseAttribute + "']/expression");
                        var mappingSourceAttribute = (string)syncRule.XPathSelectElement("attribute-mappings/mapping[dest = '" + metaverseAttribute + "']/src/attr");
                        var mappingSource = (string)syncRule.XPathSelectElement("attribute-mappings/mapping[dest = '" + metaverseAttribute + "']/src");
                        string inboundExpression = (string)mappingExpression ?? (string)mappingSourceAttribute ?? (string)mappingSource ?? "??";
                        string inboundFlowType = !string.IsNullOrEmpty(mappingExpression) ? "Expression" : "Direct";


                        Logger.Instance.WriteVerbose("Processed Sync Rule Info for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, inboundSyncRuleName);

                        // Fetch Sync Rule Scoping Conditions
                        var inboundSyncRuleScopingConditions = syncRule.XPathSelectElements("./synchronizationCriteria/conditions");
                        var inboundSyncRuleScopingConditionCount = inboundSyncRuleScopingConditions.Count();
                        var conditionIndex = -1;
                        foreach (var condition in inboundSyncRuleScopingConditions)
                        {
                            ++conditionIndex;
                            Logger.Instance.WriteVerbose("Processing Sync Rule Scope for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, inboundSyncRuleName);

                            var scopes = condition.Elements("scope");
                            // Putting all scopes into inboundSyncRuleScopingConditionString for single row in table
                            //foreach (var scope in scopes)
                            //{
                            //    string csAttribute = (string)scope.Element("csAttribute");
                            //    string csOperator = (string)scope.Element("csOperator");
                            //    string csValue = (string)scope.Element("csValue");
                            //}
                            inboundSyncRuleScopingConditionString += scopes.Select(scope => string.Format(CultureInfo.InvariantCulture, "{0} {1} {2}", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"))).Aggregate((x, j) => x + "<br><b>AND</b><br>" + j);
                            if (conditionIndex < inboundSyncRuleScopingConditionCount - 1)
                            {
                                inboundSyncRuleScopingConditionString += "<br><b>OR</b><br>";
                            }

                            Logger.Instance.WriteVerbose("Processed Sync Rule Scope for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, inboundSyncRuleName);
                        }
                        //columns and order
                        //Precedence, Connector, InboundSyncRule, inboundFlowType, Source, inboundSyncRuleScopingConditionString, metaverseAttribute, metaverseObjectType, Multivalued, Indexed
                        Documenter.AddRow(table, new object[] { rank, connectorName, inboundSyncRuleName, inboundFlowType, inboundExpression, inboundSyncRuleScopingConditionString, metaverseAttribute, metaverseObjectType, singleValue, indexed });
                    }
                }

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute);
            }
        }

        /// <summary>
        /// Creates the metaverse object type difference gram.
        /// </summary>
        private void CreateMetaverseObjectTypeDiffgram()
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
        /// Gets the metaverse object type header table.
        /// </summary>
        /// <returns>The metaverse object type header table.</returns>
        private DataTable GetMetaverseObjectTypeHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();
                //columns and order
                //Precedence, Connector, InboundSyncRule, inboundFlowType, Source, inboundSyncRuleScopingConditionString, metaverseAttribute, metaverseObjectType, Multivalued, Indexed

                // Column widths need to match the column order left to right of the final table view.  Not specific to column name
                // Updated column order
                // Rank 4
                // Connector 10
                // Inbound Sync Rule 10
                // inboundFlowType 7
                // Source 20
                // inboundSyncRuleScopingConditionString 18
                // Attribute 10
                // Type 4
                // Multi-valued 4
                // Indexed 4

                // Header Row 1
                // Precedence Display - Rank or Manual or Equal
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Rank" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 4 } }.Values.Cast<object>().ToArray());

                // Connector
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Connector" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }.Values.Cast<object>().ToArray());

                // Inbound Sync Rule
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 2 }, { "ColumnName", "Inbound Sync Rule" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }.Values.Cast<object>().ToArray());

                // inboundFlowType
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 3 }, { "ColumnName", "Inbound Flow Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 7 } }.Values.Cast<object>().ToArray());

                // Source
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 4 }, { "ColumnName", "Source" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 23 } }.Values.Cast<object>().ToArray());

                // Inbound Sync Rule Scoping Condition String
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 5 }, { "ColumnName", "Scoping Condition" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 21 } }.Values.Cast<object>().ToArray());

                // Attribute
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 6 }, { "ColumnName", "Metaverse Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 13 } }.Values.Cast<object>().ToArray());

                // Type
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 7 }, { "ColumnName", "Metaverse Object Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 4 } }.Values.Cast<object>().ToArray());

                // Multivalued
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 8 }, { "ColumnName", "Multi-valued" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 4 } }.Values.Cast<object>().ToArray());

                // Indexed
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 9 }, { "ColumnName", "Indexed" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 4 } }.Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the type of the metaverse object.
        /// </summary>
        private void PrintMetaverseObjectType()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = this.currentObjectType;

                this.WriteSectionHeader(sectionTitle, 4);

                this.ReportWriter.WriteBeginTag("div");
                this.ReportWriter.WriteAttribute("class", "EndToEndFlowsSummary");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);

                var headerTable = this.GetMetaverseObjectTypeHeaderTable();
                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, HtmlTableSize.Huge);
            }
            finally
            {
                this.ReportWriter.WriteEndTag("div");

                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Metaverse ObjectType

        #region Metaverse Object Deletion Rules

        /// <summary>
        /// Processes the metaverse object deletion rules.
        /// </summary>
        private void ProcessMetaverseObjectDeletionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Metaverse Object Deletion Rules Summary");

                var pilot = this.PilotXml.XPathSelectElements(Documenter.GetMetaverseXmlRootXPath(true) + "/mv-data//dsml:class", Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(Documenter.GetMetaverseXmlRootXPath(false) + "/mv-data//dsml:class", Documenter.NamespaceManager);

                // Sort by name
                var pilotObjectTypes = from objectType in pilot
                                       let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                       orderby name
                                       select name;

                foreach (var objectType in pilotObjectTypes)
                {
                    this.ProcessMetaverseObjectDeletionRule(objectType);
                }

                // Sort by name
                var productionObjectTypes = from objectType in production
                                            let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                            orderby name
                                            select name;

                productionObjectTypes = productionObjectTypes.Where(productionObjectType => !pilotObjectTypes.Contains(productionObjectType));

                foreach (var objectType in productionObjectTypes)
                {
                    this.ProcessMetaverseObjectDeletionRule(objectType);
                }

                this.PrintMetaverseObjectDeletionRules();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the metaverse object deletion rule.
        /// </summary>
        /// <param name="objectType">Name of the object.</param>
        private void ProcessMetaverseObjectDeletionRule(string objectType)
        {
            Logger.Instance.WriteMethodEntry("Metaverse Object Type: '{0}'.", objectType);

            try
            {
                // Set Logger call context items
                Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType, objectType);

                Logger.Instance.WriteInfo("Processing Metaverse Object Deletion Rules.");

                this.CreateMetaverseObjectDeletionRuleDataSets();

                this.FillMetaverseObjectDeletionRuleDataSet(objectType, true);
                this.FillMetaverseObjectDeletionRuleDataSet(objectType, false);

                this.CreateMetaverseObjectDeletionRuleDiffgram();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType);
            }
        }

        /// <summary>
        /// Creates the metaverse object deletion rule data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        private void CreateMetaverseObjectDeletionRuleDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("MetaverseObjectTypes") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Object Type");

                table.Columns.Add(column1);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("MetaverseObjectTypeConnectors") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Object Type");
                var column22 = new DataColumn("Connector");
                var column32 = new DataColumn("ConnectorGuid");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.PrimaryKey = new[] { column12, column22 };

                var table3 = new DataTable("MetaverseObjectTypeDeletionRules") { Locale = CultureInfo.InvariantCulture };

                var column13 = new DataColumn("Object Type");
                var column23 = new DataColumn("Connector");
                var column33 = new DataColumn("Inbound Sync Rule");
                var column43 = new DataColumn("Sync Rule Link Type");
                var column53 = new DataColumn("ConnectorGuid");
                var column63 = new DataColumn("SyncRuleGuid");

                table3.Columns.Add(column13);
                table3.Columns.Add(column23);
                table3.Columns.Add(column33);
                table3.Columns.Add(column43);
                table3.Columns.Add(column53);
                table3.Columns.Add(column63);
                table3.PrimaryKey = new[] { column13, column23, column33 };

                this.PilotDataSet = new DataSet("MetaverseObjectDeletionRules") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);
                this.PilotDataSet.Tables.Add(table3);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);
                var dataRelation23 = new DataRelation("DataRelation23", new[] { column12, column22 }, new[] { column13, column23 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);
                this.PilotDataSet.Relations.Add(dataRelation23);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetMetaverseObjectDeletionRulePrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the metaverse object deletion rule print table.
        /// </summary>
        /// <returns>The metaverse object deletion rule print table.</returns>
        private DataTable GetMetaverseObjectDeletionRulePrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Object Type
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Table 2
                // Object Type
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 2 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }.Values.Cast<object>().ToArray());

                // Table 3
                // Object Type
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Sync Rule
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 5 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // Sync Rule Type
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }.Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 4 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }.Values.Cast<object>().ToArray());

                // SyncRuleGuid
                printTable.Rows.Add(new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 5 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }.Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the metaverse object deletion rule data set.
        /// </summary>
        /// <param name="objectType">Type of the object.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillMetaverseObjectDeletionRuleDataSet(string objectType, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];
                var table2 = dataSet.Tables[1];
                var table3 = dataSet.Tables[2];

                Documenter.AddRow(table, new object[] { objectType });

                var deletionRules = config.XPathSelectElements(Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[direction = 'Inbound' " + Documenter.SyncRuleDisabledCondition + " and (linkType = 'Provision' or  linkType = 'StickyJoin') and targetObjectType = '" + objectType + "']");

                var deletionRuleIndex = -1;
                foreach (var deletionRule in deletionRules)
                {
                    ++deletionRuleIndex;
                    var deletionRuleName = (string)deletionRule.Element("name");
                    var deletionRuleGuid = (string)deletionRule.Element("id");
                    var deletionRuleLinkType = (string)deletionRule.Element("linkType");
                    var connectorGuid = (string)deletionRule.Element("connector");
                    var connectorName = (from connectorData in config.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data")
                                        where ((string)connectorData.Element("id") ?? string.Empty).Equals(connectorGuid, StringComparison.OrdinalIgnoreCase)
                                        select (string)connectorData.Element("name")).FirstOrDefault();
                    if (string.IsNullOrEmpty(connectorName))
                    {
                        Logger.Instance.WriteWarning(string.Format(CultureInfo.InvariantCulture, "Unable to dereferece connector: '{0}'. The documentation of deletion rules will be skipped for this connector. PilotConfig: '{1}'.", connectorGuid, pilotConfig));
                        continue;
                    }

                    var row2 = table2.NewRow();

                    row2[0] = objectType;
                    row2[1] = connectorName;
                    row2[2] = connectorGuid;

                    // Expected that this is likely to result in Contraint violation
                    // so make an explict check than unnesessary error getting logged.
                    if (!table2.Rows.Contains(new[] { row2[0], row2[1] }))
                    {
                        Documenter.AddRow(table2, row2);
                    }

                    var row3 = table3.NewRow();

                    row3[0] = objectType;
                    row3[1] = connectorName;
                    row3[2] = deletionRuleName;
                    row3[3] = deletionRuleLinkType;
                    row3[4] = connectorGuid;
                    row3[5] = deletionRuleGuid;

                    Documenter.AddRow(table3, row3);
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
        /// Creates the metaverse object deletion rule difference gram.
        /// </summary>
        private void CreateMetaverseObjectDeletionRuleDiffgram()
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
        /// Gets the metaverse object deletion rules header table.
        /// </summary>
        /// <returns>The metaverse object deletion rules header table.</returns>
        private DataTable GetMetaverseObjectDeletionRulesHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Object Type
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Object Type" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 20 } }.Values.Cast<object>().ToArray());

                // Deletion Rules
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Deletion Rules" }, { "RowSpan", 1 }, { "ColSpan", 3 } }.Values.Cast<object>().ToArray());

                // Header Row 2
                // Connector
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Connector" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 25 } }.Values.Cast<object>().ToArray());

                // Synchronization Rule
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Synchronization Rule" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 40 } }.Values.Cast<object>().ToArray());

                // Link Type
                headerTable.Rows.Add(new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Link Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 15 } }.Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the metaverse object deletion rule.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintMetaverseObjectDeletionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Metaverse Object Deletion Rules Summary";
                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSets.Count != 0)
                {
                    var headerTable = this.GetMetaverseObjectDeletionRulesHeaderTable();

                    #region table

                    this.ReportWriter.WriteBeginTag("table");
                    this.ReportWriter.WriteAttribute("class", HtmlTableSize.Standard.ToString() + " " + this.GetCssVisibilityClass());
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        #region thead

                        this.WriteTableHeader(headerTable);

                        #endregion thead
                    }

                    #region rows

                    foreach (var dataSet in this.DiffgramDataSets)
                    {
                        this.DiffgramDataSet = dataSet;
                        this.WriteRows(dataSet.Tables[0].Rows);
                    }

                    #endregion rows

                    this.ReportWriter.WriteEndTag("table");

                    #endregion table
                }
                else
                {
                    this.WriteContentParagraph("There are no metaverse object deletion rules configured.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Metaverse Object Deletion Rules
    }
}
