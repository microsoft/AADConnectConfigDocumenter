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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
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

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc2");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, null, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h2");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, null, "TOC");
                this.ReportWriter.WriteEndTag("h2");

                #endregion section

                this.ProcessMetaverseObjectTypes();
                this.ProcessMetaverseObjectDeletionRules();

                this.ReportWriter.Close();
                this.ReportToCWriter.Close();

                string report;
                string toc;

                using (var reportReader = new StreamReader(this.ReportFileName))
                {
                    report = reportReader.ReadToEnd();
                    using (var tocReader = new StreamReader(this.ReportToCFileName))
                    {
                        toc = tocReader.ReadToEnd();
                    }
                }

                return new Tuple<string, string>(report, toc);
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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessMetaverseObjectTypes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Metaverse Object Types";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                const string XPath = "//mv-data//dsml:class";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, null, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, null, "TOC");
                this.ReportWriter.WriteEndTag("h3");

                #endregion section

                var pilot = this.PilotXml.XPathSelectElements(XPath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(XPath, Documenter.NamespaceManager);

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
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
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

                this.CreateMetaverseObjectTypeDiffGram();

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

                var attributes = config.XPathSelectElements("//mv-data//dsml:class[dsml:name = '" + this.currentObjectType + "' ]/dsml:attribute", Documenter.NamespaceManager);

                // Sort by name
                attributes = from attribute in attributes
                             let name = (string)attribute.Attribute("ref")
                             orderby name
                             select attribute;

                for (var attributeIndex = 0; attributeIndex < attributes.Count(); ++attributeIndex)
                {
                    var attribute = attributes.ElementAt(attributeIndex);
                    var attributeName = ((string)attribute.Attribute("ref") ?? string.Empty).Trim('#');

                    // Set Logger call context items
                    Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute, attributeName);

                    Logger.Instance.WriteInfo("Processing Attribute Information.");

                    var attributeInfo = config.XPathSelectElement("//mv-data//dsml:attribute-type[dsml:name = '" + attributeName + "']", Documenter.NamespaceManager);

                    var attributeSyntax = (string)attributeInfo.Element(Documenter.DsmlNamespace + "syntax");

                    var row = table.NewRow();

                    row[0] = attributeName;
                    row[1] = Documenter.GetAttributeType(attributeSyntax, (string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexable"));
                    row[2] = ((string)attributeInfo.Attribute("single-value") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";
                    row[3] = ((string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexed") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";

                    Documenter.AddRow(table, row);

                    Logger.Instance.WriteVerbose("Processed Attribute Information.");

                    // Fetch Sync Rules
                    var syncRules = config.XPathSelectElements("//synchronizationRule[targetObjectType = '" + this.currentObjectType + "'and direction = 'Inbound' and attribute-mappings/mapping/dest = '" + attributeName + "']");
                    syncRules = from syncRule in syncRules
                                let precedence = (int)syncRule.Element("precedence")
                                orderby precedence
                                select syncRule;

                    for (var syncRuleIndex = 0; syncRuleIndex < syncRules.Count(); ++syncRuleIndex)
                    {
                        var syncRule = syncRules.ElementAt(syncRuleIndex);
                        var row2 = table2.NewRow();
                        row2[0] = attributeName;
                        row2[1] = syncRuleIndex + 1; // Care only about the precedence relative rank here than actual value
                        var connector = ((string)syncRule.Element("connector") ?? string.Empty).ToUpperInvariant();

                        var connectorName = (string)config.XPathSelectElement("//Connectors/ma-data[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connector + "']/name");
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
                        for (var conditionIndex = 0; conditionIndex < conditions.Count(); ++conditionIndex)
                        {
                            Logger.Instance.WriteVerbose("Processing Sync Rule Scope for Connector: '{0}'. Sync Rule: '{1}'.", connectorName, syncRuleName);

                            var condition = conditions.ElementAt(conditionIndex);
                            var scopes = condition.Elements("scope");
                            for (var scopeIndex = 0; scopeIndex < scopes.Count(); ++scopeIndex)
                            {
                                var scope = scopes.ElementAt(scopeIndex);
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
        private void CreateMetaverseObjectTypeDiffGram()
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
        /// Prints the type of the metaverse object.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintMetaverseObjectType()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = this.currentObjectType;

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc4" + " " + this.GetCssVisibilityClass());
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, null, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.WriteAttribute("class", this.GetCssVisibilityClass());
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteBeginTag("h4");
                this.ReportWriter.WriteAttribute("class", this.GetCssVisibilityClass());
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, null, "TOC");
                this.ReportWriter.WriteEndTag("h4");

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
                        #region head row

                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "3");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Attribute");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "3");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Type");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "3");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Multi-valued");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "3");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Indexed");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("colspan", "7");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Precedence");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();

                        #endregion head row

                        #region head row

                        this.ReportWriter.WriteBeginTag("tr");
                        {
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "2");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Rank");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "2");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Connector");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "2");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Inbound Sync Rule");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "2");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Source");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("colspan", "3");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Scoping Condition");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();

                        #endregion head row

                        #region head row

                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("CS Attribute");
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

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();

                        #endregion head row
                    }

                    this.ReportWriter.WriteEndTag("thead");

                    #endregion thead
                }

                #region rows

                this.WriteRows(this.DiffgramDataSet.Tables[0].Rows);

                #endregion rows

                this.ReportWriter.WriteEndTag("table");
                this.ReportWriter.WriteLine();
                this.ReportWriter.Flush();

                #endregion table
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Metaverse ObjectType

        #region Metaverse Object Deletion Rules

        /// <summary>
        /// Processes the metaverse object deletion rules.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessMetaverseObjectDeletionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Metaverse Object Deletion Rules Summary");

                const string XPath = "//mv-data//dsml:class";
                var pilot = this.PilotXml.XPathSelectElements(XPath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(XPath, Documenter.NamespaceManager);

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

                this.CreateMetaverseObjectDeletionRuleDiffGram();
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

                var deletionRules = config.XPathSelectElements("//SynchronizationRules/synchronizationRule[direction = 'Inbound' and (linkType = 'Provision' or  linkType = 'StickyJoin') and targetObjectType = '" + objectType + "']");

                for (var deletionRuleIndex = 0; deletionRuleIndex < deletionRules.Count(); ++deletionRuleIndex)
                {
                    var deletionRule = deletionRules.ElementAt(deletionRuleIndex);
                    var deletionRuleName = (string)deletionRule.Element("name");
                    var deletionRuleGuid = (string)deletionRule.Element("id");
                    var deletionRuleLinkType = (string)deletionRule.Element("linkType");
                    var connectorGuid = (string)deletionRule.Element("connector");
                    var connectorName = from connectorData in config.XPathSelectElements("//ma-data")
                                        where ((string)connectorData.Element("id") ?? string.Empty).Equals(connectorGuid, StringComparison.OrdinalIgnoreCase)
                                        select (string)connectorData.Element("name");

                    var row2 = table2.NewRow();

                    row2[0] = objectType;
                    row2[1] = connectorName.FirstOrDefault();
                    row2[2] = connectorGuid;

                    // Expected that this is likely to result in Contraint violation
                    // so make an explict check than unnesessary error getting logged.
                    if (!table2.Rows.Contains(new[] { row2[0], row2[1] }))
                    {
                        Documenter.AddRow(table2, row2);
                    }

                    var row3 = table3.NewRow();

                    row3[0] = objectType;
                    row3[1] = connectorName.FirstOrDefault();
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
        private void CreateMetaverseObjectDeletionRuleDiffGram()
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
        /// Prints the metaverse object deletion rule.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintMetaverseObjectDeletionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Metaverse Object Deletion Rules Summary";

                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc3");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, null, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h3");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, null, "TOC");
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
                        #region head row

                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("rowspan", "2");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Object Type");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.WriteAttribute("colspan", "3");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Deletion Rules");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();

                        #endregion head row

                        #region head row

                        this.ReportWriter.WriteBeginTag("tr");
                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        {
                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Connector");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Synchronization Rule");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Link Type");
                            this.ReportWriter.WriteEndTag("th");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();

                        #endregion head row
                    }

                    this.ReportWriter.WriteEndTag("thead");

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
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Metaverse Object Deletion Rules
    }
}
