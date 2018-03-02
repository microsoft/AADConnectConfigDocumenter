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
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

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

            try
            {
                var table = new DataTable("MetaverseObjectType") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Attribute");
                var column2 = new DataColumn("Type");
                var column3 = new DataColumn("Multi-valued");
                var column4 = new DataColumn("Indexed");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("MetaverseObjectTypePrecedence") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Attribute");
                var column22 = new DataColumn("Precedence", typeof(int));
                var column32 = new DataColumn("Connector");
                var column42 = new DataColumn("Inbound Sync Rule");
                var column52 = new DataColumn("Source");
                var column62 = new DataColumn("ConnectorGuid");
                var column72 = new DataColumn("SyncRuleGuid");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.Columns.Add(column52);
                table2.Columns.Add(column62);
                table2.Columns.Add(column72);
                table2.PrimaryKey = new[] { column12, column32, column42 };

                var table3 = new DataTable("SyncRuleScopingCondition") { Locale = CultureInfo.InvariantCulture };

                var column13 = new DataColumn("Attribute");
                var column23 = new DataColumn("Connector");
                var column33 = new DataColumn("Inbound Sync Rule");
                var column43 = new DataColumn("Group#");
                var column53 = new DataColumn("Scope#");
                var column63 = new DataColumn("CS Attribute");
                var column73 = new DataColumn("Operator");
                var column83 = new DataColumn("Value");
                var column93 = new DataColumn("ConnectorGuid");
                var column103 = new DataColumn("SyncRuleGuid");

                table3.Columns.Add(column13);
                table3.Columns.Add(column23);
                table3.Columns.Add(column33);
                table3.Columns.Add(column43);
                table3.Columns.Add(column53);
                table3.Columns.Add(column63);
                table3.Columns.Add(column73);
                table3.Columns.Add(column83);
                table3.Columns.Add(column93);
                table3.Columns.Add(column103);
                table3.PrimaryKey = new[] { column13, column23, column33, column43, column53, column63, column73 };

                this.PilotDataSet = new DataSet("MetaverseObjectType") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);
                this.PilotDataSet.Tables.Add(table3);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);
                var dataRelation23 = new DataRelation("DataRelation23", new[] { column12, column32, column42 }, new[] { column13, column23, column33 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);
                this.PilotDataSet.Relations.Add(dataRelation23);

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

                // Table 1
                // Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Multi-valued
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Indexed
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Precedence
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 5 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 6 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Source
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 5 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // SyncRuleGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 6 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 3
                // Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Group#
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 3 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Scope#
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 4 }, { "Hidden", true }, { "SortOrder", 4 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // CS Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 5 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Operator
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 6 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Value
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 7 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 8 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // SyncRuleGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 9 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

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
                var table2 = dataSet.Tables[1];
                var table3 = dataSet.Tables[2];

                var attributes = config.XPathSelectElements(Documenter.GetMetaverseXmlRootXPath(pilotConfig) + "/mv-data//dsml:class[dsml:name = '" + this.currentObjectType + "' ]/dsml:attribute", Documenter.NamespaceManager);

                // Sort by name
                attributes = from attribute in attributes
                             let name = (string)attribute.Attribute("ref")
                             orderby name
                             select attribute;
                var attributeIndex = -1;
                foreach (var attribute in attributes)
                {
                    ++attributeIndex;
                    var attributeName = ((string)attribute.Attribute("ref") ?? string.Empty).Trim('#');

                    // Set Logger call context items
                    Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute, attributeName);

                    Logger.Instance.WriteInfo("Processing Attribute Information.");

                    var attributeInfo = config.XPathSelectElement(Documenter.GetMetaverseXmlRootXPath(pilotConfig) + "/mv-data//dsml:attribute-type[dsml:name = '" + attributeName + "']", Documenter.NamespaceManager);

                    var attributeSyntax = (string)attributeInfo.Element(Documenter.DsmlNamespace + "syntax");

                    var row = table.NewRow();

                    row[0] = attributeName;
                    row[1] = Documenter.GetAttributeType(attributeSyntax, (string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexable"));
                    row[2] = ((string)attributeInfo.Attribute("single-value") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";
                    row[3] = ((string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexed") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";

                    Documenter.AddRow(table, row);

                    Logger.Instance.WriteVerbose("Processed Attribute Information.");

                    // Fetch Sync Rules
                    var syncRules = config.XPathSelectElements(Documenter.GetSynchronizationRuleXmlRootXPath(pilotConfig) + "/synchronizationRule[targetObjectType = '" + this.currentObjectType + "' and direction = 'Inbound' " + Documenter.SyncRuleDisabledCondition + " and attribute-mappings/mapping/dest = '" + attributeName + "']");
                    syncRules = from syncRule in syncRules
                                let precedence = (int)syncRule.Element("precedence")
                                orderby precedence
                                select syncRule;

                    var syncRuleIndex = -1;
                    foreach (var syncRule in syncRules)
                    {
                        ++syncRuleIndex;
                        var row2 = table2.NewRow();
                        row2[0] = attributeName;
                        row2[1] = syncRuleIndex + 1; // Care only about the precedence relative rank here than actual value
                        var connector = ((string)syncRule.Element("connector") ?? string.Empty).ToUpperInvariant();

                        var connectorName = (string)config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connector + "']/name");
                        if (string.IsNullOrEmpty(connectorName))
                        {
                            Logger.Instance.WriteWarning(string.Format(CultureInfo.InvariantCulture, "Unable to dereferece connector: '{0}'. The documentation of inbound flows will be skipped for this connector. PilotConfig: '{1}'.", connector, pilotConfig));
                            continue;
                        }

                        var syncRuleName = (string)syncRule.Element("name");
                        row2[2] = connectorName;
                        row2[3] = syncRuleName;

                        Logger.Instance.WriteVerbose("Processing Sync Rule Info for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, syncRuleName);

                        var mappingExpression = syncRule.XPathSelectElement("attribute-mappings/mapping[dest = '" + attributeName + "']/expression");
                        var mappingSourceAttribute = syncRule.XPathSelectElement("attribute-mappings/mapping[dest = '" + attributeName + "']/src/attr");
                        var mappingSource = syncRule.XPathSelectElement("attribute-mappings/mapping[dest = '" + attributeName + "']/src");

                        row2[4] = (string)mappingExpression ?? (string)mappingSourceAttribute ?? (string)mappingSource ?? "??";
                        row2[5] = connector;
                        row2[6] = (string)syncRule.Element("id");

                        Documenter.AddRow(table2, row2);

                        Logger.Instance.WriteVerbose("Processed Sync Rule Info for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, syncRuleName);

                        // Fetch Sync Rule Scoping Conditions
                        var conditions = syncRule.XPathSelectElements("synchronizationCriteria/conditions");
                        var conditionIndex = -1;
                        foreach (var condition in conditions)
                        {
                            ++conditionIndex;

                            Logger.Instance.WriteVerbose("Processing Sync Rule Scope for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, syncRuleName);

                            var scopes = condition.Elements("scope");
                            var scopeIndex = -1;
                            foreach (var scope in scopes)
                            {
                                ++scopeIndex;
                                var row3 = table3.NewRow();
                                row3[0] = attributeName;
                                row3[1] = connectorName;
                                row3[2] = syncRuleName;
                                row3[3] = conditionIndex;
                                row3[4] = scopeIndex;
                                row3[5] = (string)scope.Element("csAttribute");
                                row3[6] = (string)scope.Element("csOperator");
                                row3[7] = (string)scope.Element("csValue");
                                row3[8] = connector;
                                row3[9] = (string)syncRule.Element("id");

                                Documenter.AddRow(table3, row3);
                            }

                            Logger.Instance.WriteVerbose("Processed Sync Rule Scope for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, syncRuleName);
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

                // Header Row 1
                // Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Attribute" }, { "RowSpan", 3 }, { "ColSpan", 1 }, { "ColWidth", 13 } }).Values.Cast<object>().ToArray());

                // Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Type" }, { "RowSpan", 3 }, { "ColSpan", 1 }, { "ColWidth", 7 } }).Values.Cast<object>().ToArray());

                // Multi-valued
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 2 }, { "ColumnName", "Multi-valued" }, { "RowSpan", 3 }, { "ColSpan", 1 }, { "ColWidth", 4 } }).Values.Cast<object>().ToArray());

                // Indexed
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 3 }, { "ColumnName", "Indexed" }, { "RowSpan", 3 }, { "ColSpan", 1 }, { "ColWidth", 4 } }).Values.Cast<object>().ToArray());

                // Precedence
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 4 }, { "ColumnName", "Precedence" }, { "RowSpan", 1 }, { "ColSpan", 7 }, { "ColWidth", 0 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Precedence Display - Rank or Manual or Equal
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Rank" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 4 } }).Values.Cast<object>().ToArray());

                // Connector
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Connector" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 15 } }).Values.Cast<object>().ToArray());

                // Inbound Sync Rule
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Inbound Sync Rule" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 15 } }).Values.Cast<object>().ToArray());

                // Source
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Source" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 15 } }).Values.Cast<object>().ToArray());

                // Scoping Condition
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 4 }, { "ColumnName", "Scoping Condition" }, { "RowSpan", 1 }, { "ColSpan", 3 }, { "ColWidth", 0 } }).Values.Cast<object>().ToArray());

                // Header Row 3
                // CS Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 2 }, { "ColumnIndex", 0 }, { "ColumnName", "CS Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Operator
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 2 }, { "ColumnIndex", 1 }, { "ColumnName", "Operator" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 8 } }).Values.Cast<object>().ToArray());

                // Value
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 2 }, { "ColumnIndex", 2 }, { "ColumnName", "Value" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 5 } }).Values.Cast<object>().ToArray());

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
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 2 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Table 3
                // Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Sync Rule
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 5 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Sync Rule Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 4 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // SyncRuleGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 5 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

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
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Object Type" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 20 } }).Values.Cast<object>().ToArray());

                // Deletion Rules
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Deletion Rules" }, { "RowSpan", 1 }, { "ColSpan", 3 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Connector
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Connector" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 25 } }).Values.Cast<object>().ToArray());

                // Synchronization Rule
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Synchronization Rule" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 40 } }).Values.Cast<object>().ToArray());

                // Link Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Link Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 15 } }).Values.Cast<object>().ToArray());

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
