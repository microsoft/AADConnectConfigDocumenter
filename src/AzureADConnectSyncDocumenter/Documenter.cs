//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Documenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Base class for Azure AD Connect Sync Configuration Documenter
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace AzureADConnectConfigDocumenter
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Web.UI;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The abstract base class for the documenters of all other types of AAD Connect sync configuration items.
    /// </summary>
    public abstract class Documenter
    {
        /// <summary>
        /// The name of the column which stores the row state of the row in a diffgram.
        /// </summary>
        public const string RowStateColumn = "ROW-STATE";

        /// <summary>
        /// The name of the column which stores the display visibility status of the row in a diffgram.
        /// </summary>
        public const string HtmlTableRowVisibilityStatusColumn = "HTML-TABLE-ROW-VISIBILITY-STATUS";

        /// <summary>
        /// The literal string to indicate visibility can be hidden
        /// </summary>
        public const string CanHide = "CanHide";

        /// <summary>
        /// The literal string to indicate visibility cannot be hidden
        /// </summary>
        public const string NoHide = "NoHide";

        /// <summary>
        /// The prefix used for the name of the columns of old data row in a diffgram 
        /// </summary>
        public const string OldColumnPrefix = "OLD-";

        /// <summary>
        /// The lower case letters
        /// </summary>
        public const string LowercaseLetters = "abcdefghijklmnopqrstuvwxyz";

        /// <summary>
        /// The upper case letters
        /// </summary>
        public const string UppercaseLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        /// <summary>
        /// The row error text added to identify a row that was to an empty table so that it prints better
        /// Such row is not printed if it appears as a deleted row in the diffgram
        /// </summary>
        public const string VanityRow = "~~VANITY_ROW~~";

        /// <summary>
        /// The maximum number of sortable columns in a table
        /// </summary>
        public const int MaxSortableColumns = 10;

        /// <summary>
        /// The XPath condition for disabled synchronization rules
        /// </summary>
        public const string SyncRuleDisabledCondition = " and (count(disabled) = 0 or (disabled != 'True' and disabled != 'true' and disabled != '1')) ";

        /// <summary>
        /// The namespace manager
        /// </summary>
        private static XmlNamespaceManager namespaceManager = new XmlNamespaceManager(new NameTable());

        /// <summary>
        /// Initializes static members of the <see cref="Documenter"/> class.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1810:InitializeReferenceTypeStaticFieldsInline", Justification = "Reviewed.")]
        static Documenter()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Documenter.namespaceManager.AddNamespace("dsml", Documenter.DsmlNamespace.NamespaceName);
                Documenter.namespaceManager.AddNamespace("ms-dsml", Documenter.MmsDsmlNamespace.NamespaceName);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Documenter"/> class.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataSet.")]
        protected Documenter()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                // Set Logger call context items
                Logger.FlushContextItems();
                this.DiffgramDataSet = new DataSet() { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSets = new List<DataSet>();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="Documenter"/> class.
        /// </summary>
        ~Documenter()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (File.Exists(this.ReportFileName))
                {
                    File.Delete(this.ReportFileName);
                }

                if (File.Exists(this.ReportToCFileName))
                {
                    File.Delete(this.ReportToCFileName);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Enumerator indicating whether the config element is present in pilot, production or both
        /// </summary>
        public enum ConfigEnvironment : int
        {
            /// <summary>
            /// The config element exists in Pilot as well as Production
            /// </summary>
            PilotAndProduction = 0,

            /// <summary>
            /// The config element exists in only in the Pilot
            /// </summary>
            PilotOnly,

            /// <summary>
            /// The config element exists in only in the Production
            /// </summary>
            ProductionOnly
        }

        /// <summary>
        /// Enumerator indicating the HTML table size
        /// </summary>
        public enum HtmlTableSize : int
        {
            /// <summary>
            /// The standard / default html table size (50%)
            /// </summary>
            Standard = 0,

            /// <summary>
            /// The large html table size (75%)
            /// </summary>
            Large,

            /// <summary>
            /// The extra large table size (95%)
            /// </summary>
            Huge
        }

        /// <summary>
        /// Gets the report folder.
        /// </summary>
        /// <value>
        /// The report folder.
        /// </value>
        public static string ReportFolder
        {
            get
            {
                var rootDirectory = Directory.GetCurrentDirectory().TrimEnd('\\');
                var reportFolder = rootDirectory + @"\Report";
                if (!Directory.Exists(reportFolder))
                {
                    Directory.CreateDirectory(reportFolder);
                }

                return reportFolder;
            }
        }

        /// <summary>
        /// Gets the namespace manager.
        /// </summary>
        /// <value>
        /// The namespace manager.
        /// </value>
        protected static XmlNamespaceManager NamespaceManager
        {
            get { return Documenter.namespaceManager; }
        }

        /// <summary>
        /// Gets the DSML namespace.
        /// </summary>
        /// <value>
        /// The DSML namespace.
        /// </value>
        protected static XNamespace DsmlNamespace
        {
            get { return "http://www.dsml.org/DSML"; }
        }

        /// <summary>
        /// Gets the MMS DSML namespace.
        /// </summary>
        /// <value>
        /// The MMS DSML namespace.
        /// </value>
        protected static XNamespace MmsDsmlNamespace
        {
            get { return "http://www.microsoft.com/MMS/DSML"; }
        }

        /// <summary>
        /// Gets or sets the consolidated pilot XML.
        /// </summary>
        /// <value>
        /// The pilot XML.
        /// </value>
        protected XElement PilotXml { get; set; }

        /// <summary>
        /// Gets or sets the consolidated production XML.
        /// </summary>
        /// <value>
        /// The production XML.
        /// </value>
        protected XElement ProductionXml { get; set; }

        /// <summary>
        /// Gets or sets the pilot data set.
        /// </summary>
        /// <value>
        /// The pilot data set.
        /// </value>
        protected DataSet PilotDataSet { get; set; }

        /// <summary>
        /// Gets or sets a value indicating which environment the config element is present.
        /// </summary>
        protected ConfigEnvironment Environment { get; set; }

        /// <summary>
        /// Gets or sets the production data set.
        /// </summary>
        /// <value>
        /// The production data set.
        /// </value>
        protected DataSet ProductionDataSet { get; set; }

        /// <summary>
        /// Gets or sets the diffgram data set.
        /// </summary>
        /// <value>
        /// The diffgram data set.
        /// </value>
        protected DataSet DiffgramDataSet { get; set; }

        /// <summary>
        /// Gets or sets the diffgram data sets of a single section.
        /// </summary>
        /// <value>
        /// The diffgram data sets of a single section.
        /// </value>
        protected List<DataSet> DiffgramDataSets { get; set; }

        /// <summary>
        /// Gets or sets the main report writer.
        /// </summary>
        /// <value>
        /// The main report writer.
        /// </value>
        protected XhtmlTextWriter ReportWriter { get; set; }

        /// <summary>
        /// Gets or sets the report Table of Content writer.
        /// </summary>
        /// <value>
        /// The report Table of Content writer.
        /// </value>
        protected XhtmlTextWriter ReportToCWriter { get; set; }

        /// <summary>
        /// Gets or sets the name of the report file.
        /// </summary>
        /// <value>
        /// The name of the report file.
        /// </value>
        protected string ReportFileName { get; set; }

        /// <summary>
        /// Gets or sets the name of the report to c file.
        /// </summary>
        /// <value>
        /// The name of the report to c file.
        /// </value>
        protected string ReportToCFileName { get; set; }

        /// <summary>
        /// Gets the temporary file path.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>
        /// The temporary file path.
        /// </returns>
        public static string GetTempFilePath(string fileName)
        {
            fileName = Path.GetInvalidFileNameChars().Aggregate(fileName + "_" + Guid.NewGuid().ToString("N"), (current, c) => current.Replace(c, '-'));

            return Path.GetTempPath() + @"\" + fileName;
        }

        /// <summary>
        /// Gets the report file's base name.
        /// </summary>
        /// <param name="pilotConfigRelativePath">The relative path of the pilot / test configuration directory.</param>
        /// <param name="productionConfigRelativePath">The relative path of the production configuration directory.</param>
        /// <returns>
        /// The report file's base name.
        /// </returns>
        public static string GetReportFileBaseName(string pilotConfigRelativePath, string productionConfigRelativePath)
        {
            return (pilotConfigRelativePath ?? string.Empty).Replace(@"\", "_") + "_AppliedTo_" + (productionConfigRelativePath ?? string.Empty).Replace(@"\", "_");
        }

        /// <summary>
        /// Gets the type of the attribute from it's DSML syntax.
        /// </summary>
        /// <param name="syntax">The DSML syntax of the attribute.</param>
        /// <param name="indexable">If <c>"true"</c>, the attribute is indexable.</param>
        /// <returns>
        /// The type of the attribute.
        /// </returns>
        public static string GetAttributeType(string syntax, string indexable)
        {
            Logger.Instance.WriteMethodEntry("Syntax: '{0}'. Indexable: '{1}'.", syntax, indexable);

            var attributeType = syntax;

            try
            {
                var attributeSuffix = (indexable == "true") ? " (indexable)" : " (non-indexable)";

                attributeType = Documenter.GetAttributeType(syntax) + attributeSuffix;

                return attributeType;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Syntax: '{0}'. Indexable: '{1}'. Attribute Type: '{2}'.", syntax, indexable, attributeType);
            }
        }

        /// <summary>
        /// Gets the type of the attribute from it's DSML syntax.
        /// </summary>
        /// <param name="syntax">The DSML syntax of the attribute.</param>
        /// <returns>
        /// Returns the type of the attribute.
        /// </returns>
        public static string GetAttributeType(string syntax)
        {
            Logger.Instance.WriteMethodEntry("Syntax: '{0}'.", syntax);

            var attributeType = syntax;

            try
            {
                switch (syntax)
                {
                    case "1.3.6.1.4.1.1466.115.121.1.27":
                    case "1.2.840.113556.1.4.906":
                        attributeType = "Number";
                        break;
                    case "1.3.6.1.4.1.1466.115.121.1.7":
                        attributeType = "Boolean";
                        break;
                    case "1.3.6.1.4.1.1466.115.121.1.5":
                    case "1.3.6.1.4.1.1466.115.121.1.40":
                        attributeType = "Binary";
                        break;
                    case "1.3.6.1.4.1.1466.115.121.1.15":
                    case "1.2.840.113556.1.4.1221":
                    case "1.2.840.113556.1.4.905":
                    case "1.3.6.1.4.1.1466.115.121.1.44":
                    case "1.2.840.113556.1.4.1362":
                        attributeType = "String";
                        break;
                    case "1.3.6.1.4.1.1466.115.121.1.12":
                        attributeType = "Reference (DN)";
                        break;
                    case "1.3.6.1.4.1.1466.115.121.1.24":
                    case "1.3.6.1.4.1.1466.115.121.1.53":
                        attributeType = "DateTime";
                        break;
                }

                return attributeType;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Syntax: '{0}'. Attribute Type: '{1}'.", syntax, attributeType);
            }
        }

        /// <summary>
        /// Gets the metaverse configuration report.
        /// </summary>
        /// <returns>The Tuple of configuration report and associated TOC</returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "The method performs a time-consuming operation.")]
        public abstract Tuple<string, string> GetReport();

        /// <summary>
        /// Gets the XPath for "mv-data" node
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        /// <returns>Returns "mv-data" XPath</returns>
        protected static string GetMetaverseXmlRootXPath(bool pilotConfig)
        {
            return (pilotConfig ? "/Pilot" : "/Production") + "/GlobalSettings/child::*";
        }

        /// <summary>
        /// Gets the XPath for "ma-data" node
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        /// <returns>Returns "ma-data" XPath</returns>
        protected static string GetConnectorXmlRootXPath(bool pilotConfig)
        {
            return (pilotConfig ? "/Pilot" : "/Production") + "/Connectors";
        }

        /// <summary>
        /// Gets the XPath for "synchronizationRule" node
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        /// <returns>Returns "synchronizationRule" XPath</returns>
        protected static string GetSynchronizationRuleXmlRootXPath(bool pilotConfig)
        {
            return (pilotConfig ? "/Pilot" : "/Production") + "/SynchronizationRules";
        }

        /// <summary>
        /// Gets the Bookmark code for the specified bookmark text.
        /// </summary>
        /// <param name="text">The Bookmark text.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        /// <returns>
        /// The Bookmark code for the given bookmark text.
        /// </returns>
        protected static string GetBookmarkCode(string text, string sectionGuid)
        {
            Logger.Instance.WriteMethodEntry("Bookmark Text: '{0}'. Section Guid: '{1}'.", text, sectionGuid);

            var bookmarkCode = string.Empty;

            try
            {
                // MS Word does not like "-" in the bookmarks
                bookmarkCode = (sectionGuid + text).ToUpperInvariant().GetHashCode().ToString(CultureInfo.InvariantCulture).Replace("-", "_");

                return bookmarkCode;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Bookmark Text: '{0}'. Section Guid: '{1}'. Bookmark Code: '{2}'.", text, sectionGuid, bookmarkCode);
            }
        }

        /// <summary>
        /// Writes the jump to bookmark location.
        /// </summary>
        /// <param name="htmlWriter">The htmlWriter.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="displayText">The display text.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        /// <param name="cellClass">The cell class.</param>
        protected static void WriteJumpToBookmarkLocation(HtmlTextWriter htmlWriter, string bookmark, string displayText, string sectionGuid, string cellClass)
        {
            Logger.Instance.WriteMethodEntry("Bookmark Code: '{0}'. Bookmark Text: '{1}'. Section Guid: '{2}'. Cell Class: '{3}'.", bookmark, displayText, sectionGuid, cellClass);

            try
            {
                if (htmlWriter == null)
                {
                    throw new ArgumentNullException("htmlWriter");
                }

                htmlWriter.WriteBeginTag("a");
                htmlWriter.WriteAttribute("class", cellClass);
                htmlWriter.WriteAttribute("href", "#" + Documenter.GetBookmarkCode(bookmark, sectionGuid));
                htmlWriter.Write(HtmlTextWriter.TagRightChar);
                htmlWriter.Write(displayText);
                htmlWriter.WriteEndTag("a");
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Bookmark Code: '{0}'. Bookmark Text: '{1}'. Section Guid: '{2}'. Cell Class: '{3}'.", bookmark, displayText, sectionGuid, cellClass);
            }
        }

        /// <summary>
        /// Gets the diffgram.
        /// </summary>
        /// <param name="pilotDataSet">The pilot data set.</param>
        /// <param name="productionDataSet">The production data set.</param>
        /// <returns>
        /// An <see cref="DataSet"/> object representing the diffgram of the two data sets.
        /// </returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataSet.")]
        protected static DataSet GetDiffgram(DataSet pilotDataSet, DataSet productionDataSet)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing changes. This may take a few minutes...");

                if (pilotDataSet == null)
                {
                    throw new ArgumentNullException("pilotDataSet");
                }

                if (productionDataSet == null)
                {
                    throw new ArgumentNullException("productionDataSet");
                }

                var diffgramDataSet = new DataSet(pilotDataSet.DataSetName) { Locale = CultureInfo.InvariantCulture };

                var printTable = pilotDataSet.Tables["PrintSettings"];

                for (var i = 0; i < pilotDataSet.Tables.Count; ++i)
                {
                    if (pilotDataSet.Tables[i].TableName == "PrintSettings")
                    {
                        continue;
                    }

                    var columnsIgnored = printTable.Select("ChangeIgnored = true AND TableIndex = " + i).Select(row => (int)row["ColumnIndex"]).ToArray();
                    var diffgramTable = Documenter.GetDiffgram(pilotDataSet.Tables[i], productionDataSet.Tables[i], columnsIgnored);

                    diffgramDataSet.Tables.Add(diffgramTable);
                }

                diffgramDataSet.Tables.Add(printTable.Copy());
                diffgramDataSet = Documenter.SortDataSet(diffgramDataSet);

                // set up data relations
                foreach (DataRelation dataRelation in pilotDataSet.Relations)
                {
                    var parentColumns = new List<DataColumn>();
                    foreach (var parentColumn in dataRelation.ParentColumns)
                    {
                        var tableName = parentColumn.Table.TableName;
                        var columnIndex = parentColumn.Ordinal;
                        parentColumns.Add(diffgramDataSet.Tables[tableName].Columns[columnIndex]);
                    }

                    var childColumns = new List<DataColumn>();
                    foreach (var childColumn in dataRelation.ChildColumns)
                    {
                        var tableName = childColumn.Table.TableName;
                        var columnIndex = childColumn.Ordinal;
                        childColumns.Add(diffgramDataSet.Tables[tableName].Columns[columnIndex]);
                    }

                    var dataRelationClone = new DataRelation(dataRelation.RelationName, parentColumns.ToArray(), childColumns.ToArray(), false);
                    diffgramDataSet.Relations.Add(dataRelationClone);
                }

                Documenter.SetHtmlTableRowVisibilityStatus(diffgramDataSet);

                return diffgramDataSet;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Adds the visibility status column to each table of the diffgram data set.
        /// </summary>
        /// <param name="diffgramDataSet">The diffgram data set.</param>
        protected static void AddRowVisibilityStatusColumn(DataSet diffgramDataSet)
        {
            if (diffgramDataSet == null)
            {
                throw new ArgumentNullException("diffgramDataSet");
            }

            diffgramDataSet.ExtendedProperties.Add(Documenter.CanHide, true);

            // Loop all ignore last table which is PrintSetting table
            for (var i = 0; i < diffgramDataSet.Tables.Count - 1; ++i)
            {
                try
                {
                    diffgramDataSet.Tables[i].Columns.Add(Documenter.HtmlTableRowVisibilityStatusColumn);
                }
                catch (DuplicateNameException)
                {
                    // just ignore
                }
            }

            diffgramDataSet.AcceptChanges();
        }

        /// <summary>
        /// Sets the visibility status of each row of the html table based on the modification status of the row or it's children.
        /// </summary>
        /// <param name="diffgramDataSet">The diffgram data set.</param>
        protected static void SetHtmlTableRowVisibilityStatus(DataSet diffgramDataSet)
        {
            Documenter.AddRowVisibilityStatusColumn(diffgramDataSet);

            // Loop all ignore last table which is PrintSetting table
            for (var i = 0; i < diffgramDataSet.Tables.Count - 1; ++i)
            {
                foreach (DataRow row in diffgramDataSet.Tables[i].Rows)
                {
                    if (i == 0)
                    {
                        if (!Documenter.IsCumulativeRowStateChanged(row, i))
                        {
                            row[Documenter.HtmlTableRowVisibilityStatusColumn] = Documenter.CanHide;
                        }
                        else
                        {
                            diffgramDataSet.ExtendedProperties[Documenter.CanHide] = false;
                        }
                    }
                    else if (!Documenter.IsCumulativeRowStateChanged(row, i))
                    {
                        try
                        {
                            var dataRelationName = string.Format(CultureInfo.InvariantCulture, "DataRelation{0}{1}", i, i + 1);
                            var parentRow = row.GetParentRow(dataRelationName);
                            row[Documenter.HtmlTableRowVisibilityStatusColumn] = parentRow[Documenter.HtmlTableRowVisibilityStatusColumn];
                        }
                        catch (DataException e)
                        {
                            var rowString = (row != null) ? string.Join("|", row.ItemArray) : string.Empty;
                            var errorMsg = e.Message + " Data Row: " + rowString + e.StackTrace;
                            Logger.Instance.WriteError(errorMsg);
                        }
                    }
                }

                diffgramDataSet.Tables[i].AcceptChanges();
            }
        }

        /// <summary>
        /// Gets the diffgram.
        /// </summary>
        /// <param name="pilotTable">The pilot table.</param>
        /// <param name="productionTable">The production table.</param>
        /// <param name="columnsIgnored">The columns ignored when calculating diffgram.</param>
        /// <returns>
        /// An <see cref="DataTable" /> object representing the diffgram of the two tables.
        /// </returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected static DataTable GetDiffgram(DataTable pilotTable, DataTable productionTable, int[] columnsIgnored)
        {
            Logger.Instance.WriteMethodEntry();

            DataTable diffgramTable = CreateDiffgramTable(pilotTable);

            try
            {
                if (columnsIgnored == null)
                {
                    columnsIgnored = new int[] { -1 };
                }

                // Unchanged rows in pilotTable: pilotRow PrimaryKey matches productionRow PrimaryKey AND pilotRow also matches productionRow:
                var unchangedPilotRows = from pilotRow in pilotTable.AsEnumerable()
                                         from productionRow in productionTable.AsEnumerable()
                                         where pilotTable.PrimaryKey.Aggregate(true, (match, keyColumn) => match && pilotRow[keyColumn].Equals(productionRow[keyColumn.Ordinal]))
                                               && pilotRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)).SequenceEqual(productionRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)))
                                         select pilotRow;

                // Unchanged rows in productionTable: pilotRow PrimaryKey matches productionRow PrimaryKey AND pilotRow also matches productionRow:
                var unchangedProductionRows = from pilotRow in pilotTable.AsEnumerable()
                                              from productionRow in productionTable.AsEnumerable()
                                              where pilotTable.PrimaryKey.Aggregate(true, (match, keyColumn) => match && pilotRow[keyColumn].Equals(productionRow[keyColumn.Ordinal]))
                                                    && pilotRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)).SequenceEqual(productionRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)))
                                              select productionRow;

                // Modified rows in pilotTable : pilotRow PrimaryKey matches productionRow PrimaryKey BUT pilotRow does not match productionRow:
                var modifiedPilotRows = from pilotRow in pilotTable.AsEnumerable()
                                        from productionRow in productionTable.AsEnumerable()
                                        where pilotTable.PrimaryKey.Aggregate(true, (match, keyColumn) => match && pilotRow[keyColumn].Equals(productionRow[keyColumn.Ordinal]))
                                              && !pilotRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)).SequenceEqual(productionRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)))
                                        select pilotRow;

                // Modified rows in productionTable : pilotRow PrimaryKey matches productionRow PrimaryKey BUT pilotRow does not match productionRow:
                var modifiedProductionRows = from pilotRow in pilotTable.AsEnumerable()
                                             from productionRow in productionTable.AsEnumerable()
                                             where pilotTable.PrimaryKey.Aggregate(true, (match, keyColumn) => match && pilotRow[keyColumn].Equals(productionRow[keyColumn.Ordinal]))
                                                 && !pilotRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)).SequenceEqual(productionRow.ItemArray.Where((item, index) => !columnsIgnored.Contains(index)))
                                             select productionRow;

                // Added rows : rows in pilotTable not modified AND not unchanged
                var addedRows = pilotTable.AsEnumerable().Except(modifiedPilotRows, DataRowComparer.Default).Except(unchangedPilotRows, DataRowComparer.Default);

                // Deleted rows : rows in productionTable not modified AND not unchanged
                var deletedRows = productionTable.AsEnumerable().Except(modifiedProductionRows, DataRowComparer.Default).Except(unchangedProductionRows, DataRowComparer.Default);

                // Populate unchanged rows
                foreach (var row in unchangedPilotRows)
                {
                    var newRow = diffgramTable.NewRow();
                    newRow[Documenter.RowStateColumn] = DataRowState.Unchanged;
                    foreach (DataColumn column in pilotTable.Columns)
                    {
                        newRow[column.ColumnName] = row[column.ColumnName];
                    }

                    Documenter.AddRow(diffgramTable, newRow);
                }

                // Populate modified rows
                foreach (var row in modifiedPilotRows)
                {
                    // Match the unmodified version of the row via the PrimaryKey
                    ////var matchInProductionTable = modifiedProductionRows.Where(mondifiedProductionRow => productionTable.PrimaryKey.Aggregate(true, (match, keyColumn) => match && mondifiedProductionRow[keyColumn].Equals(row[keyColumn.Ordinal]))).First();
                    // revised query for perf improvements
                    var matchInProductionTable = (from productionRow in productionTable.AsEnumerable()
                                                 where productionTable.PrimaryKey.Aggregate(true, (match, keyColumn) => match && productionRow[keyColumn].Equals(row[keyColumn.Ordinal]))
                                                 select productionRow).First();
                    var newRow = diffgramTable.NewRow();
                    newRow[Documenter.RowStateColumn] = DataRowState.Modified;

                    // Set the row with the original values
                    foreach (DataColumn column in pilotTable.Columns)
                    {
                        if (!pilotTable.PrimaryKey.Contains(column))
                        {
                            newRow[Documenter.OldColumnPrefix + column.ColumnName] = matchInProductionTable[column.ColumnName];
                        }
                    }

                    // Set the modified values
                    foreach (DataColumn column in pilotTable.Columns)
                    {
                        newRow[column.ColumnName] = row[column.ColumnName];
                    }

                    Documenter.AddRow(diffgramTable, newRow);
                }

                // Populate added rows
                foreach (var row in addedRows)
                {
                    var newRow = diffgramTable.NewRow();
                    newRow[Documenter.RowStateColumn] = DataRowState.Added;
                    foreach (DataColumn column in pilotTable.Columns)
                    {
                        newRow[column.ColumnName] = row[column.ColumnName];
                    }

                    Documenter.AddRow(diffgramTable, newRow);
                }

                // Populate deleted rows
                foreach (var row in deletedRows)
                {
                    if (row.RowError == Documenter.VanityRow)
                    {
                        break;
                    }

                    var newRow = diffgramTable.NewRow();
                    newRow[Documenter.RowStateColumn] = DataRowState.Deleted;
                    foreach (DataColumn column in pilotTable.Columns)
                    {
                        if (!pilotTable.PrimaryKey.Contains(column))
                        {
                            newRow[column.ColumnName] = row[column.ColumnName]; // save the value in the orginal column as well in case it's used for Sorting table
                            newRow[Documenter.OldColumnPrefix + column.ColumnName] = row[column.ColumnName];
                        }
                        else
                        {
                            newRow[column.ColumnName] = row[column.ColumnName];
                        }
                    }

                    Documenter.AddRow(diffgramTable, newRow);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }

            return diffgramTable;
        }

        /// <summary>
        /// Creates the diffgram table structure based on input table.
        /// </summary>
        /// <param name="dataTable">The input table to base the diffgram table on.</param>
        /// <returns>The diffgram table structure based on input table.</returns>
        protected static DataTable CreateDiffgramTable(DataTable dataTable)
        {
            if (dataTable == null)
            {
                throw new ArgumentNullException("dataTable");
            }

            var diffgramTable = dataTable.Clone();
            diffgramTable.Columns.Add(Documenter.RowStateColumn);
            foreach (DataColumn column in dataTable.Columns)
            {
                if (!dataTable.PrimaryKey.Contains(column))
                {
                    diffgramTable.Columns.Add(Documenter.OldColumnPrefix + column.ColumnName);
                }
            }

            return diffgramTable;
        }

        /// <summary>
        /// Gets the print settings table for a configuration section.
        /// </summary>
        /// <returns>
        /// The print settings table for a configuration section.
        /// </returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "The method performs a time-consuming operation.")]
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected static DataTable GetPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = new DataTable("PrintSettings") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("TableIndex", typeof(int));
                var column2 = new DataColumn("ColumnIndex", typeof(int));
                var column3 = new DataColumn("Hidden", typeof(bool));
                var column4 = new DataColumn("SortOrder", typeof(int));
                var column5 = new DataColumn("BookmarkIndex", typeof(int));
                var column6 = new DataColumn("JumpToBookmarkIndex", typeof(int));
                var column7 = new DataColumn("ChangeIgnored", typeof(bool));

                printTable.Columns.Add(column1);
                printTable.Columns.Add(column2);
                printTable.Columns.Add(column3);
                printTable.Columns.Add(column4);
                printTable.Columns.Add(column5);
                printTable.Columns.Add(column6);
                printTable.Columns.Add(column7);
                printTable.PrimaryKey = new[] { column1, column2 };

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the header table for a configuration section.
        /// </summary>
        /// <returns>
        /// The header table for a configuration section.
        /// </returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "The method performs a time-consuming operation.")]
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected static DataTable GetHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = new DataTable("HeaderTable") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("RowIndex", typeof(int));
                var column2 = new DataColumn("ColumnIndex", typeof(int));
                var column3 = new DataColumn("ColumnName", typeof(string));
                var column4 = new DataColumn("RowSpan", typeof(int));
                var column5 = new DataColumn("ColSpan", typeof(int));
                var column6 = new DataColumn("ColWidth", typeof(int)); // polulate it as zero for any colspan rows and populate the width only the child rows

                headerTable.Columns.Add(column1);
                headerTable.Columns.Add(column2);
                headerTable.Columns.Add(column3);
                headerTable.Columns.Add(column4);
                headerTable.Columns.Add(column5);
                headerTable.Columns.Add(column6);
                headerTable.PrimaryKey = new[] { column1, column2 };

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the simple settings header table with multiple columns.
        /// </summary>
        /// <param name="columnNames">The column names of the table header.</param>
        /// <returns>
        /// The simple settings header table with multiple columns.
        /// </returns>
        protected static DataTable GetSimpleSettingsHeaderTable(OrderedDictionary columnNames)
        {
            return Documenter.GetSimpleSettingsHeaderTable(null, columnNames);
        }

        /// <summary>
        /// Gets the simple settings header table with first header row having single column and second header row with multiple columns.
        /// </summary>
        /// <param name="columnName">The column name of the first row table header.</param>
        /// <param name="columnNames">The column names of the second row table header.</param>
        /// <returns>
        /// The simple settings header table.
        /// </returns>
        protected static DataTable GetSimpleSettingsHeaderTable(string columnName, OrderedDictionary columnNames)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnName);

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                if (!string.IsNullOrEmpty(columnName) && columnNames != null && columnNames.Count != 0)
                {
                    headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", columnName }, { "RowSpan", 1 }, { "ColSpan", columnNames.Count } }).Values.Cast<object>().ToArray());

                    for (var i = 0; i < columnNames.Count; ++i)
                    {
                        headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", i }, { "ColumnName", columnNames.Cast<DictionaryEntry>().ElementAt(i).Key }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", columnNames[i] } }).Values.Cast<object>().ToArray());
                    }
                }
                else if (string.IsNullOrEmpty(columnName) && columnNames != null && columnNames.Count != 0)
                {
                    for (var i = 0; i < columnNames.Count; ++i)
                    {
                        headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", i }, { "ColumnName", columnNames.Cast<DictionaryEntry>().ElementAt(i).Key }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", columnNames[i] } }).Values.Cast<object>().ToArray());
                    }
                }
                else if (!string.IsNullOrEmpty(columnName))
                {
                    headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", columnName }, { "RowSpan", 1 }, { "ColSpan", 2 } }).Values.Cast<object>().ToArray());
                }

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnName);
            }
        }

        /// <summary>
        /// Sorts a diffgram dataset.
        /// </summary>
        /// <param name="diffgramDataSet">The diffgram dataset.</param>
        /// <returns>
        /// The sorted diffgram dataset.
        /// </returns>
        protected static DataSet SortDataSet(DataSet diffgramDataSet)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (diffgramDataSet == null)
                {
                    throw new ArgumentNullException("diffgramDataSet");
                }

                var sortedDiffgramDataSet = diffgramDataSet.Clone();

                var printTable = diffgramDataSet.Tables["PrintSettings"];
                sortedDiffgramDataSet.Tables["PrintSettings"].Merge(printTable);

                for (var i = 0; i < diffgramDataSet.Tables.Count; ++i)
                {
                    if (diffgramDataSet.Tables[i].TableName == "PrintSettings")
                    {
                        continue;
                    }

                    var table = diffgramDataSet.Tables[i];
                    var sortColumns = printTable.Select("SortOrder <> -1 AND TableIndex = " + i, "SortOrder").Select(row => (int)row["ColumnIndex"]).ToArray();

                    var tableClone = table.Clone();

                    IEnumerable<DataRow> rows;

                    switch (sortColumns.Length)
                    {
                        case 0:
                            rows = table.Rows.Cast<DataRow>();
                            break;
                        case 1:
                            rows = table.Rows.Cast<DataRow>().OrderBy(row => row[sortColumns[0]]);
                            break;
                        case 2:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]]
                                   select row;
                            break;
                        case 3:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]]
                                   select row;
                            break;
                        case 4:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]], row[sortColumns[3]]
                                   select row;
                            break;
                        case 5:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]], row[sortColumns[3]], row[sortColumns[4]]
                                   select row;
                            break;
                        case 6:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]], row[sortColumns[3]], row[sortColumns[4]], row[sortColumns[5]]
                                   select row;
                            break;
                        case 7:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]], row[sortColumns[3]], row[sortColumns[4]], row[sortColumns[5]], row[sortColumns[6]]
                                   select row;
                            break;
                        case 8:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]], row[sortColumns[3]], row[sortColumns[4]], row[sortColumns[5]], row[sortColumns[6]], row[sortColumns[7]]
                                   select row;
                            break;
                        case 9:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]], row[sortColumns[3]], row[sortColumns[4]], row[sortColumns[5]], row[sortColumns[6]], row[sortColumns[7]], row[sortColumns[8]]
                                   select row;
                            break;
                        case 10:
                            rows = from row in table.Rows.Cast<DataRow>()
                                   orderby row[sortColumns[0]], row[sortColumns[1]], row[sortColumns[2]], row[sortColumns[3]], row[sortColumns[4]], row[sortColumns[5]], row[sortColumns[6]], row[sortColumns[7]], row[sortColumns[8]], row[sortColumns[9]]
                                   select row;
                            break;
                        default:
                            throw new ArgumentOutOfRangeException("columns", string.Format(CultureInfo.InvariantCulture, "Columns length must be between {0} and {1}.", 1, Documenter.MaxSortableColumns));
                    }

                    foreach (var row in rows)
                    {
                        tableClone.ImportRow(row);
                    }

                    tableClone.AcceptChanges();

                    sortedDiffgramDataSet.Tables[i].Merge(tableClone);
                }

                sortedDiffgramDataSet.AcceptChanges();

                return sortedDiffgramDataSet;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Adds the row to the table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="row">The row.</param>
        protected static void AddRow(DataTable table, object row)
        {
            AddRow(table, row, false);
        }

        /// <summary>
        /// Adds the row to the table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="row">The row.</param>
        /// <param name="vanityRow">if set to <c>true</c>, the row is not printed if it appears as a deleted row in the diffgram.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily", Justification = "Reviewed.")]
        protected static void AddRow(DataTable table, object row, bool vanityRow)
        {
            if (table == null)
            {
                throw new ArgumentNullException("table");
            }

            if (row == null)
            {
                throw new ArgumentNullException("row");
            }

            try
            {
                var dataRow = row as DataRow;
                if (dataRow != null)
                {
                    table.Rows.Add(dataRow);
                }
                else
                {
                    var values = row as object[];
                    if (values != null)
                    {
                        dataRow = table.Rows.Add(values);
                    }
                    else
                    {
                        throw new ArgumentException("Parameter must be a DataRow or object[].", "row");
                    }
                }

                if (vanityRow)
                {
                    dataRow.RowError = Documenter.VanityRow;
                }

                dataRow.AcceptChanges();
            }
            catch (DataException e)
            {
                var dataRow = row as DataRow;
                var values = (dataRow != null) ? dataRow.ItemArray : row as object[];
                var rowString = (values != null) ? string.Join("|", values) : string.Empty;
                var errorMsg = e.Message + " Data Row: " + rowString + e.StackTrace;
                Logger.Instance.WriteError(errorMsg);
            }
        }

        /// <summary>
        /// Gets the embedded script resource
        /// </summary>
        /// <param name="resourceName">Name of the embedded script resource</param>
        /// <returns>The script resource</returns>
        protected static string GetEmbeddedScriptResource(string resourceName)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var qualifiedResource = "AzureADConnectConfigDocumenter.Scripts." + resourceName;
                using (StreamReader reader = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream(qualifiedResource)))
                {
                    return reader.ReadToEnd();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Determines whether the specified row or any child rows have changed.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="currentTableIndex">Index of the current table.</param>
        /// <returns>True if the specified row or any child rows have changed. Otherwise false.</returns>
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "currentTableIndex+1", Justification = "Reviewed.")]
        protected static bool IsCumulativeRowStateChanged(DataRow row, int currentTableIndex)
        {
            Logger.Instance.WriteMethodEntry("Current Table Index: '{0}'.", currentTableIndex);

            var rowStateChanged = false;

            try
            {
                if (row == null)
                {
                    throw new ArgumentNullException("row");
                }

                var rowState = (string)row[Documenter.RowStateColumn];

                if (!rowState.Equals(DataRowState.Unchanged.ToString()))
                {
                    rowStateChanged = true;
                    return rowStateChanged;
                }

                var childTableIndex = currentTableIndex + 1;
                var dataRelationName = string.Format(CultureInfo.InvariantCulture, "DataRelation{0}{1}", currentTableIndex + 1, childTableIndex + 1);
                var childRows = row.GetChildRows(dataRelationName);
                var childRowsCount = childRows.Count();
                for (var i = 0; i < childRowsCount; ++i)
                {
                    var childRowState = (string)childRows[i][Documenter.RowStateColumn];

                    if (!childRowState.Equals(DataRowState.Unchanged.ToString()))
                    {
                        rowStateChanged = true;
                        return rowStateChanged;
                    }

                    if (Documenter.IsCumulativeRowStateChanged(childRows[i], currentTableIndex + 1))
                    {
                        rowStateChanged = true;
                        return rowStateChanged;
                    }
                }

                return rowStateChanged;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Current Table Index: '{0}'. RowStateChanged: '{1}'.", currentTableIndex, rowStateChanged);
            }
        }

        /// <summary>
        /// Gets the configuration report tuple.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "Reviewed.")]
        protected Tuple<string, string> GetReportTuple()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
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

        /// <summary>
        /// Writes the report.
        /// </summary>
        /// <param name="reportHeader">The report header.</param>
        /// <param name="reportHtml">The report html.</param>
        /// <param name="tocHtml">The TOC html.</param>
        /// <param name="pilotConfigPath">The pilot config path.</param>
        /// <param name="productionConfigPath">The production config path.</param>
        /// <param name="configReportFilePath">The config report file path.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        protected void WriteReport(string reportHeader, string reportHtml, string tocHtml, string pilotConfigPath, string productionConfigPath, string configReportFilePath)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(configReportFilePath));

                this.ReportWriter.WriteFullBeginTag("html");

                this.WriteReportHeader();

                this.ReportWriter.WriteFullBeginTag("body");

                this.ReportWriter.WriteFullBeginTag("h1");
                this.ReportWriter.Write(reportHeader);
                this.ReportWriter.WriteEndTag("h1");
                this.ReportWriter.WriteLine();

                this.WriteDocumenterInfo();

                var syncVersionXPath = "/mv-data//parameter-values/parameter[@name = 'Microsoft.Synchronize.ServerConfigurationVersion']";
                var syncVersionPilot = (string)this.PilotXml.XPathSelectElement(Documenter.GetMetaverseXmlRootXPath(true) + syncVersionXPath);
                var syncVersionProduction = (string)this.ProductionXml.XPathSelectElement(Documenter.GetMetaverseXmlRootXPath(false) + syncVersionXPath);
                var cellClass = (syncVersionPilot == syncVersionProduction ? DataRowState.Unchanged : DataRowState.Modified).ToString();

                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Target / Pilot Config ");
                {
                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", cellClass);
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write("(" + syncVersionPilot + "):");
                    this.ReportWriter.WriteEndTag("span");
                }

                this.ReportWriter.WriteEndTag("strong");
                {
                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", DataRowState.Unchanged.ToString());
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write(pilotConfigPath);
                    this.ReportWriter.WriteEndTag("span");

                    this.WriteBreakTag();
                }

                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Reference / Production Config ");
                {
                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", cellClass);
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write("(" + syncVersionProduction + "):");
                    this.ReportWriter.WriteEndTag("span");
                }

                this.ReportWriter.WriteEndTag("strong");
                {
                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", DataRowState.Unchanged.ToString());
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write(productionConfigPath);
                    this.ReportWriter.WriteEndTag("span");

                    this.WriteBreakTag();
                }

                this.ReportWriter.WriteLine();

                this.ReportWriter.WriteFullBeginTag("h1");
                this.ReportWriter.Write("Table of Contents");
                this.ReportWriter.WriteEndTag("h1");
                this.ReportWriter.WriteLine();

                this.ReportWriter.WriteLine(tocHtml);
                this.ReportWriter.WriteLine(reportHtml);

                this.ReportWriter.WriteEndTag("body");
                this.ReportWriter.WriteEndTag("html");
            }
            finally
            {
                this.ReportWriter.Close();
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Writes the report header.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void WriteReportHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter.WriteFullBeginTag("head");

                #region meta

                this.ReportWriter.WriteBeginTag("meta");
                this.ReportWriter.WriteAttribute("http-equiv", "Content-Type");
                this.ReportWriter.WriteAttribute("content", "text/html; charset=UTF-8");
                this.ReportWriter.WriteLine(XhtmlTextWriter.SelfClosingTagEnd);

                #endregion meta

                #region style

                ////this.ReportWriter.WriteBeginTag("link");
                ////this.ReportWriter.WriteAttribute("rel", "stylesheet");
                ////this.ReportWriter.WriteAttribute("type", "text/css");
                ////this.ReportWriter.WriteAttribute("href", "documenter.css");
                ////this.ReportWriter.WriteLine(XhtmlTextWriter.SelfClosingTagEnd);

                this.ReportWriter.WriteBeginTag("style");
                this.ReportWriter.WriteAttribute("type", "text/css");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportWriter.Write(GetEmbeddedScriptResource("Documenter.css"));
                this.ReportWriter.WriteEndTag("style");
                this.ReportWriter.WriteLine();

                #endregion style

                #region script

                this.ReportWriter.WriteFullBeginTag("script");
                this.ReportWriter.WriteLine();
                this.ReportWriter.WriteLine(GetEmbeddedScriptResource("Documenter.js"));
                this.ReportWriter.WriteEndTag("script");
                this.ReportWriter.WriteLine();

                #endregion script

                this.ReportWriter.WriteFullBeginTag("title");
                this.ReportWriter.Write("AAD Connect Config Documenter Report");
                this.ReportWriter.WriteEndTag("title");

                this.ReportWriter.WriteEndTag("head");
                this.ReportWriter.WriteLine();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Writes the documenter information.
        /// </summary>
        protected void WriteDocumenterInfo()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Only Show Changes:");
                this.ReportWriter.WriteEndTag("strong");

                this.ReportWriter.WriteBeginTag("input");
                this.ReportWriter.WriteAttribute("type", "checkbox");
                this.ReportWriter.WriteAttribute("id", "OnlyShowChanges");
                this.ReportWriter.WriteAttribute("disabled", null);
                this.ReportWriter.WriteAttribute("onclick", "ToggleVisibility();");
                this.ReportWriter.WriteLine(HtmlTextWriter.SelfClosingTagEnd);

                this.ReportWriter.WriteBeginTag("a");
                this.ReportWriter.WriteAttribute("style", "display: none;");
                this.ReportWriter.WriteAttribute("href", "#");
                this.ReportWriter.WriteAttribute("id", "DownloadLink");
                this.ReportWriter.WriteAttribute("onclick", "return DownloadScript();");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportWriter.Write("Download Sync Rule Changes Script");
                this.ReportWriter.WriteEndTag("a");

                this.WriteBreakTag();

                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Hide Default Sync Rules:");
                this.ReportWriter.WriteEndTag("strong");

                this.ReportWriter.WriteBeginTag("input");
                this.ReportWriter.WriteAttribute("type", "checkbox");
                this.ReportWriter.WriteAttribute("id", "HideDefaultSyncRules");
                this.ReportWriter.WriteAttribute("disabled", null);
                this.ReportWriter.WriteAttribute("onclick", "ToggleDefaultRuleVisibility();");
                this.ReportWriter.WriteLine(HtmlTextWriter.SelfClosingTagEnd);

                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Hide End-to-end Summary Flows:");
                this.ReportWriter.WriteEndTag("strong");

                this.ReportWriter.WriteBeginTag("input");
                this.ReportWriter.WriteAttribute("type", "checkbox");
                this.ReportWriter.WriteAttribute("id", "HideEndToEndFlowsSummary");
                this.ReportWriter.WriteAttribute("disabled", null);
                this.ReportWriter.WriteAttribute("onclick", "ToggleEndToEndFlowsSummaryVisibility();");
                this.ReportWriter.WriteLine(HtmlTextWriter.SelfClosingTagEnd);

                this.WriteBreakTag();

                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Legend:");
                this.ReportWriter.WriteEndTag("strong");

                {
                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", DataRowState.Added.ToString());
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write("Create ");
                    this.ReportWriter.WriteEndTag("span");

                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", DataRowState.Modified.ToString());
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write("Update ");
                    this.ReportWriter.WriteEndTag("span");

                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", DataRowState.Deleted.ToString());
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write("Delete ");
                    this.ReportWriter.WriteEndTag("span");

                    this.WriteBreakTag();
                }

                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Documenter Version:");
                this.ReportWriter.WriteEndTag("strong");

                {
                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", DataRowState.Unchanged.ToString());
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write(VersionInfo.Version);
                    this.ReportWriter.WriteEndTag("span");

                    this.WriteBreakTag();
                }

                this.ReportWriter.WriteFullBeginTag("strong");
                this.ReportWriter.Write("Report Date:");
                this.ReportWriter.WriteEndTag("strong");

                {
                    this.ReportWriter.WriteBeginTag("span");
                    this.ReportWriter.WriteAttribute("class", DataRowState.Unchanged.ToString());
                    this.ReportWriter.WriteLine(HtmlTextWriter.TagRightChar);
                    this.ReportWriter.Write(DateTime.Now.ToString(CultureInfo.CurrentCulture));
                    this.ReportWriter.WriteEndTag("span");

                    this.WriteBreakTag();
                }

                this.ReportWriter.WriteBeginTag("div");
                this.ReportWriter.WriteAttribute("class", "PowerShellScript");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportWriter.WriteLine(Documenter.GetEmbeddedScriptResource("PowerShellScriptHeader.ps1"));
                this.ReportWriter.WriteLine();
                this.ReportWriter.WriteEndTag("div");

                this.ReportWriter.WriteLine();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Writes the bookmark location.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        /// <param name="anchorClass">The anchor style class.</param>
        protected void WriteBookmarkLocation(string bookmark, string sectionGuid, string anchorClass)
        {
            this.WriteBookmarkLocation(bookmark, bookmark, sectionGuid, anchorClass);
        }

        /// <summary>
        /// Writes the bookmark location.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="displayText">The display text.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        /// <param name="anchorClass">The anchor style class.</param>
        protected void WriteBookmarkLocation(string bookmark, string displayText, string sectionGuid, string anchorClass)
        {
            Logger.Instance.WriteMethodEntry("Bookmark Code: '{0}'. Bookmark Text: '{1}'. Section Guid: '{2}'. Anchor Class: '{3}'.", bookmark, displayText, sectionGuid, anchorClass);

            try
            {
                this.ReportWriter.WriteBeginTag("a");
                this.ReportWriter.WriteAttribute("class", anchorClass);
                this.ReportWriter.WriteAttribute("name", Documenter.GetBookmarkCode(bookmark, sectionGuid));
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportWriter.Write(displayText);
                this.ReportWriter.WriteEndTag("a");
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Bookmark Code: '{0}'. Bookmark Text: '{1}'. Section Guid: '{2}'. Anchor Class: '{3}'.", bookmark, displayText, sectionGuid, anchorClass);
            }
        }

        /// <summary>
        /// Writes the jump to bookmark location.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        /// <param name="cellClass">The cell class.</param>
        protected void WriteJumpToBookmarkLocation(string bookmark, string sectionGuid, string cellClass)
        {
            this.WriteJumpToBookmarkLocation(bookmark, bookmark, sectionGuid, cellClass);
        }

        /// <summary>
        /// Writes the jump to bookmark location in TOC.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        protected void WriteJumpToBookmarkLocationInTOC(string bookmark, string sectionGuid)
        {
            this.WriteJumpToBookmarkLocationInTOC(bookmark, bookmark, sectionGuid);
        }

        /// <summary>
        /// Writes the jump to bookmark location.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="displayText">The display text.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        /// <param name="cellClass">The cell class.</param>
        protected void WriteJumpToBookmarkLocation(string bookmark, string displayText, string sectionGuid, string cellClass)
        {
            Documenter.WriteJumpToBookmarkLocation(this.ReportWriter, bookmark, displayText, sectionGuid, cellClass);
        }

        /// <summary>
        /// Writes the jump to bookmark location.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="displayText">The display text.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        protected void WriteJumpToBookmarkLocationInTOC(string bookmark, string displayText, string sectionGuid)
        {
            Logger.Instance.WriteMethodEntry("Bookmark Code: '{0}'. Bookmark Text: '{1}'. Section Guid: '{2}'.", bookmark, displayText, sectionGuid);

            try
            {
                this.ReportToCWriter.WriteBeginTag("a");
                this.ReportToCWriter.WriteAttribute("class", this.GetAnchorClassForTOC());
                this.ReportToCWriter.WriteAttribute("href", "#" + Documenter.GetBookmarkCode(bookmark, sectionGuid));
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportToCWriter.Write(displayText);
                this.ReportToCWriter.WriteEndTag("a");
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Bookmark Code: '{0}'. Bookmark Text: '{1}'. Section Guid: '{2}'.", bookmark, displayText, sectionGuid);
            }
        }

        /// <summary>
        /// Gets the anchor style class for TOC
        /// </summary>
        /// <returns>The anchor style class for TOC</returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "Reviewed.")]
        protected string GetAnchorClassForTOC()
        {
            var anchorClass = "toc";

            return this.Environment == ConfigEnvironment.ProductionOnly ? anchorClass + "-Deleted" : this.Environment == ConfigEnvironment.PilotOnly ? anchorClass + "-Added" : anchorClass;
        }

        /// <summary>
        /// Resets the diffgram and prepares the variable for new report section.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataSet.")]
        protected void ResetDiffgram()
        {
            this.DiffgramDataSet = new DataSet() { Locale = CultureInfo.InvariantCulture };
            this.DiffgramDataSets = new List<DataSet>();
        }

        /// <summary>
        /// Gets the CSS visibility class.
        /// </summary>
        /// <returns>The CSS visibility class, either CanHide or empty string</returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "Reviewed.")]
        protected virtual string GetCssVisibilityClass()
        {
            return this.DiffgramDataSets.Count == 0 || this.DiffgramDataSets.Any(dataSet => !(bool)dataSet.ExtendedProperties[Documenter.CanHide]) ? string.Empty : Documenter.CanHide;
        }

        /// <summary>
        /// Writes the rows.
        /// </summary>
        /// <param name="rows">The rows.</param>
        protected void WriteRows(DataRowCollection rows)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (rows == null)
                {
                    throw new ArgumentNullException("rows");
                }

                int currentTableIndex = 0;
                int currentCellIndex = 0;
                this.WriteRows(rows.Cast<DataRow>().ToArray(), currentTableIndex, ref currentCellIndex);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Writes the rows.
        /// </summary>
        /// <param name="rows">The rows.</param>
        /// <param name="currentTableIndex">Index of the current table.</param>
        /// <param name="currentCellIndex">Index of the current cell.</param>
        [SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "2#", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "currentTableIndex+1", Justification = "Reviewed.")]
        protected void WriteRows(DataRow[] rows, int currentTableIndex, ref int currentCellIndex)
        {
            Logger.Instance.WriteMethodEntry("Current Table Index: '{0}'. Current Cell Index: '{1}'.", currentTableIndex, currentCellIndex);

            try
            {
                if (rows == null)
                {
                    throw new ArgumentNullException("rows");
                }
                else if (rows.Count() == 0)
                {
                    return;
                }

                var printTable = this.DiffgramDataSet.Tables["PrintSettings"];
                var maxCellCount = printTable.Select("Hidden = false").Count();

                var rowCount = rows.Length;

                for (var i = 0; i < rowCount; ++i)
                {
                    var row = rows[i];
                    var cellClass = (string)row[Documenter.RowStateColumn];

                    if (currentCellIndex == 0)
                    {
                        // Start the new row
                        this.ReportWriter.WriteBeginTag("tr");
                        if (row[Documenter.HtmlTableRowVisibilityStatusColumn] as string == Documenter.CanHide)
                        {
                            this.ReportWriter.WriteAttribute("class", cellClass + " " + Documenter.CanHide);
                        }
                        else
                        {
                            this.ReportWriter.WriteAttribute("class", cellClass);
                        }

                        this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    }

                    currentCellIndex = printTable.Select("Hidden = false AND TableIndex < " + currentTableIndex).Count();

                    var rowSpan = this.GetRowSpan(row, currentTableIndex) + 1;

                    var printColumns = printTable.Select("Hidden = false AND TableIndex = " + currentTableIndex).Select(rowX => rowX["ColumnIndex"]);
                    foreach (var column in row.Table.Columns.Cast<DataColumn>().Where(column => printColumns.Contains(column.Ordinal)))
                    {
                        this.WriteCell(row, column, rowSpan, currentTableIndex);
                        ++currentCellIndex;
                    }

                    // child rows
                    var childTableIndex = currentTableIndex + 1;
                    var dataRelationName = string.Format(CultureInfo.InvariantCulture, "DataRelation{0}{1}", childTableIndex, childTableIndex + 1);
                    var childRows = row.GetChildRows(dataRelationName);
                    var childRowsCount = childRows.Count();

                    if (childRowsCount == 0)
                    {
                        // complete the row if required
                        for (; currentCellIndex < maxCellCount; ++currentCellIndex)
                        {
                            this.ReportWriter.WriteBeginTag("td");
                            this.ReportWriter.WriteAttribute("class", cellClass);
                            this.ReportWriter.WriteAttribute("rowspan", "1");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("-");
                            this.ReportWriter.WriteEndTag("td");
                        }

                        this.ReportWriter.WriteEndTag("tr");
                        this.ReportWriter.WriteLine();
                        currentCellIndex = 0; // reset the current cell index
                    }
                    else
                    {
                        currentCellIndex = currentCellIndex == maxCellCount ? 0 : currentCellIndex;
                        this.WriteRows(childRows, childTableIndex, ref currentCellIndex);
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Current Table Index: '{0}'. Current Cell Index: '{1}'.", currentTableIndex, currentCellIndex);
            }
        }

        /// <summary>
        /// Gets the row span for the cells of the current row by counting all the predecessor rows.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="currentTableIndex">Index of the current table.</param>
        /// <returns>The number of all predecessor rows</returns>
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "currentTableIndex+1", Justification = "Reviewed.")]
        protected int GetRowSpan(DataRow row, int currentTableIndex)
        {
            Logger.Instance.WriteMethodEntry("Current Table Index: '{0}'.", currentTableIndex);

            try
            {
                if (row == null)
                {
                    throw new ArgumentNullException("row");
                }

                var rowSpan = 0;
                var childTableIndex = currentTableIndex + 1;
                var dataRelationName = string.Format(CultureInfo.InvariantCulture, "DataRelation{0}{1}", childTableIndex, childTableIndex + 1);
                var childRows = row.GetChildRows(dataRelationName);
                var childRowsCount = childRows.Count();
                for (var i = 0; i < childRowsCount; ++i)
                {
                    if (i > 0)
                    {
                        ++rowSpan;
                    }

                    rowSpan += this.GetRowSpan(childRows[i], currentTableIndex + 1);
                }

                return rowSpan;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Current Table Index: '{0}'.", currentTableIndex);
            }
        }

        /// <summary>
        /// Writes the cell.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="rowSpan">The row span.</param>
        /// <param name="tableIndex">Index of the table.</param>
        protected void WriteCell(DataRow row, DataColumn column, int rowSpan, int tableIndex)
        {
            Logger.Instance.WriteMethodEntry("Row Span: '{0}'. Table Index: '{1}'.", rowSpan, tableIndex);

            try
            {
                if (row == null)
                {
                    throw new ArgumentNullException("row");
                }

                if (column == null)
                {
                    throw new ArgumentNullException("column");
                }

                var printTable = this.DiffgramDataSet.Tables["PrintSettings"];
                var rowFilter = string.Format(CultureInfo.InvariantCulture, "TableIndex = {0} AND ColumnIndex = {1} AND (BookmarkIndex <> -1 OR JumpToBookmarkIndex <> -1)", tableIndex, column.Ordinal);
                var printRow = printTable.Select(rowFilter).FirstOrDefault();
                var bookmarkIndex = printRow != null && (int)printRow["BookmarkIndex"] != -1 ? printRow["BookmarkIndex"] : null;
                var jumpToBookmarkIndex = printRow != null && (int)printRow["JumpToBookmarkIndex"] != -1 ? printRow["JumpToBookmarkIndex"] : null;

                var cellClass = (string)row[Documenter.RowStateColumn];

                this.ReportWriter.WriteBeginTag("td");
                this.ReportWriter.WriteAttribute("class", cellClass);
                this.ReportWriter.WriteAttribute("rowspan", rowSpan.ToString(CultureInfo.InvariantCulture));
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);

                if (column.Table.PrimaryKey.Contains(column))
                {
                    var text = Convert.ToString(row[column.ColumnName], CultureInfo.InvariantCulture);
                    this.WriteCellText(text, (int?)bookmarkIndex, (int?)jumpToBookmarkIndex, row, cellClass);
                }
                else
                {
                    var rowState = (DataRowState)Enum.Parse(typeof(DataRowState), row[Documenter.RowStateColumn].ToString());

                    switch (rowState)
                    {
                        case DataRowState.Modified:
                            {
                                var oldText = Convert.ToString(row[Documenter.OldColumnPrefix + column.ColumnName], CultureInfo.InvariantCulture);
                                var text = Convert.ToString(row[column.ColumnName], CultureInfo.InvariantCulture);

                                if (oldText != text)
                                {
                                    cellClass = DataRowState.Deleted.ToString();
                                    this.ReportWriter.WriteBeginTag("span");
                                    this.ReportWriter.WriteAttribute("class", cellClass);
                                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                    this.WriteCellText(oldText, (int?)bookmarkIndex, (int?)jumpToBookmarkIndex, row, cellClass);
                                    this.ReportWriter.WriteEndTag("span");
                                }

                                this.ReportWriter.WriteBeginTag("span");
                                this.ReportWriter.WriteAttribute("class", DataRowState.Modified.ToString());
                                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                                this.WriteCellText(text, (int?)bookmarkIndex, (int?)jumpToBookmarkIndex, row, cellClass);
                                this.ReportWriter.WriteEndTag("span");
                                break;
                            }

                        case DataRowState.Deleted:
                            {
                                cellClass = DataRowState.Deleted.ToString();
                                var text = Convert.ToString(row[Documenter.OldColumnPrefix + column.ColumnName], CultureInfo.InvariantCulture);
                                this.WriteCellText(text, (int?)bookmarkIndex, (int?)jumpToBookmarkIndex, row, cellClass);
                                break;
                            }

                        default:
                            {
                                var text = Convert.ToString(row[column.ColumnName], CultureInfo.InvariantCulture);
                                this.WriteCellText(text, (int?)bookmarkIndex, (int?)jumpToBookmarkIndex, row, cellClass);
                                break;
                            }
                    }
                }

                this.ReportWriter.WriteEndTag("td");
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Row Span: '{0}'. Table Index: '{1}'.", rowSpan, tableIndex);
            }
        }

        /// <summary>
        /// Write the cell text or bookmark markup
        /// </summary>
        /// <param name="cellText">The cell text.</param>
        /// <param name="bookmarkIndex">The bookmark column index.</param>
        /// <param name="jumpToBookmarkIndex">The jump to bookmark column index.</param>
        /// <param name="row">The row.</param>
        /// <param name="cellClass">The cell class.</param>
        protected void WriteCellText(string cellText, int? bookmarkIndex, int? jumpToBookmarkIndex, DataRow row, string cellClass)
        {
            if (row == null)
            {
                return;
            }

            var bookmark = bookmarkIndex != null ? Convert.ToString(row[(int)bookmarkIndex], CultureInfo.InvariantCulture) : null;
            var jumpToBookmark = jumpToBookmarkIndex != null ? Convert.ToString(row[(int)jumpToBookmarkIndex], CultureInfo.InvariantCulture) : null;

            if (bookmark != null)
            {
                this.WriteBookmarkLocation(cellText, bookmark, cellClass);
            }
            else if (jumpToBookmark != null)
            {
                if (!string.IsNullOrEmpty(jumpToBookmark))
                {
                    this.WriteJumpToBookmarkLocation(cellText, jumpToBookmark, cellClass);
                }
                else
                {
                    this.ReportWriter.Write(cellText);
                }
            }
            else
            {
                this.ReportWriter.Write(cellText);
            }
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        protected virtual void WriteSectionHeader(string title, int level)
        {
            this.WriteSectionHeader(title, level, null);
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        protected void WriteSectionHeader(string title, int level, string sectionGuid)
        {
            this.WriteSectionHeader(title, level, title, sectionGuid);
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        protected void WriteSectionHeader(string title, int level, string bookmark, string sectionGuid)
        {
            this.WriteToCEntry(title, level, bookmark, sectionGuid);

            this.ReportWriter.WriteBeginTag("h" + level);
            this.ReportWriter.WriteAttribute("class", this.GetCssVisibilityClass());
            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);

            this.WriteBookmarkLocation(bookmark, title, sectionGuid, this.GetAnchorClassForTOC());
            this.ReportWriter.WriteEndTag("h" + level);
            this.ReportWriter.WriteLine();
        }

        /// <summary>
        /// Writes the ToC Entry
        /// </summary>
        /// <param name="entryText">The ToC entry text</param>
        /// <param name="level">The ToC item level</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="sectionGuid">The section unique identifier.</param>
        protected void WriteToCEntry(string entryText, int level, string bookmark, string sectionGuid)
        {
            this.ReportToCWriter.WriteBeginTag("span");
            this.ReportToCWriter.WriteAttribute("class", "toc" + level + " " + this.GetCssVisibilityClass());
            this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
            this.WriteJumpToBookmarkLocationInTOC(bookmark, entryText, sectionGuid);
            this.ReportToCWriter.WriteEndTag("span");
            this.ReportToCWriter.WriteBeginTag("br");
            this.ReportToCWriter.WriteAttribute("class", this.GetCssVisibilityClass());
            this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
            this.ReportToCWriter.WriteLine();
        }

        /// <summary>
        /// Writes the table header cell
        /// </summary>
        /// <param name="cellText">The cell text</param>
        /// <param name="rowSpan">The row span of the cell</param>
        /// <param name="columnSpan">The column span of the cell</param>
        protected void WriteTableHeaderCell(string cellText, int rowSpan, int columnSpan)
        {
            this.ReportWriter.WriteBeginTag("th");
            this.ReportWriter.WriteAttribute("class", "column-th");
            this.ReportWriter.WriteAttribute("rowspan", rowSpan.ToString(CultureInfo.InvariantCulture));
            this.ReportWriter.WriteAttribute("colspan", columnSpan.ToString(CultureInfo.InvariantCulture));
            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
            this.ReportWriter.Write(cellText);
            this.ReportWriter.WriteEndTag("th");
        }

        /// <summary>
        /// Writes the section data table
        /// </summary>
        /// <param name="dataTable">The section data table</param>
        /// <param name="headerTable">The section data header table</param>
        /// <param name="tableSize">The size (width) of the HTML table</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void WriteTable(DataTable dataTable, DataTable headerTable, HtmlTableSize tableSize = HtmlTableSize.Standard)
        {
            #region table

            this.ReportWriter.WriteBeginTag("table");
            this.ReportWriter.WriteAttribute("class", tableSize.ToString() + " " + this.GetCssVisibilityClass());
            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
            {
                this.WriteTableHeader(headerTable);
            }

            #region rows

            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                this.WriteRows(dataTable.Rows);
            }

            #endregion rows

            this.ReportWriter.WriteEndTag("table");
            this.ReportWriter.WriteLine();
            this.ReportWriter.Flush();

            #endregion table
        }

        /// <summary>
        /// Writes the section data table header
        /// </summary>
        /// <param name="headerTable">The section data header table</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void WriteTableHeader(DataTable headerTable)
        {
            if (headerTable != null && headerTable.Rows.Count != 0)
            {
                #region column group

                this.ReportWriter.WriteFullBeginTag("colgroup");

                var columnWidths = from row in headerTable.Rows.Cast<DataRow>()
                                   orderby row["RowIndex"], row["ColumnIndex"]
                                   select row["ColWidth"] == DBNull.Value ? 0 : (int)row["ColWidth"];
                foreach (var columnWidth in columnWidths)
                {
                    // skip any colspan rows as child rows will have width defined
                    if (columnWidth != 0)
                    {
                        this.ReportWriter.WriteBeginTag("col");
                        this.ReportWriter.WriteAttribute("style", "width:" + columnWidth + "%;");
                        this.ReportWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                    }
                }

                this.ReportWriter.WriteEndTag("colgroup");

                #endregion column group

                #region thead

                this.ReportWriter.WriteBeginTag("thead");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                {
                    #region head row(s)

                    var rows = from row in headerTable.Rows.Cast<DataRow>()
                               orderby row["RowIndex"], row["ColumnIndex"]
                               select row;

                    var currentRowIndex = -1;
                    foreach (var row in rows)
                    {
                        if ((int)row["RowIndex"] != currentRowIndex)
                        {
                            if (currentRowIndex != -1)
                            {
                                this.ReportWriter.WriteEndTag("tr");
                                this.ReportWriter.WriteLine();
                            }

                            currentRowIndex = (int)row["RowIndex"];

                            this.ReportWriter.WriteBeginTag("tr");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                        }

                        this.WriteTableHeaderCell(row["ColumnName"] as string, (int)row["RowSpan"], (int)row["ColSpan"]);
                    }

                    this.ReportWriter.WriteEndTag("tr");
                    this.ReportWriter.WriteLine();

                    #endregion head row(s)
                }

                this.ReportWriter.WriteEndTag("thead");

                #endregion thead
            }
        }

        /// <summary>
        /// Writes the content paragraph
        /// </summary>
        /// <param name="content">The content text</param>
        protected void WriteContentParagraph(string content)
        {
            this.WriteContentParagraph(content, this.GetCssVisibilityClass());
        }

        /// <summary>
        /// Writes the content paragraph
        /// </summary>
        /// <param name="content">The content text</param>
        /// <param name="cssClass">The paragraph style class</param>
        protected void WriteContentParagraph(string content, string cssClass)
        {
            this.ReportWriter.WriteLine();
            this.ReportWriter.WriteBeginTag("p");
            this.ReportWriter.WriteAttribute("class", cssClass);
            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
            this.ReportWriter.WriteLine(content);
            this.ReportWriter.WriteEndTag("p");
        }

        /// <summary>
        /// Writes the break tag.
        /// </summary>
        protected void WriteBreakTag()
        {
            this.ReportWriter.WriteBeginTag("br");
            this.ReportWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
        }

        #region Simple Settings Sections

        /// <summary>
        /// Creates the simple settings data sets.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSimpleSettingsDataSets(int columnCount)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnCount);

            try
            {
                this.CreateSimpleSettingsDataSets(columnCount, 0);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnCount);
            }
        }

        /// <summary>
        /// Creates the simple settings data sets.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <param name="keyIndex">The zero-based index of the primary key column.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSimpleSettingsDataSets(int columnCount, int keyIndex)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'. Key Index: '{1}'.", columnCount, keyIndex);

            try
            {
                var table = new DataTable("SimpleSettings") { Locale = CultureInfo.InvariantCulture };

                for (var i = 0; i < columnCount; ++i)
                {
                    table.Columns.Add(new DataColumn("Column" + (i + 1)));
                }

                table.PrimaryKey = new[] { table.Columns[keyIndex] };

                this.PilotDataSet = new DataSet("SimpleSettings") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetSimpleSettingsPrintTable(columnCount, keyIndex);
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'. Key Index: '{1}'.", columnCount, keyIndex);
            }
        }

        /// <summary>
        /// Gets the simple settings print table.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <returns>
        /// The simple settings print table.
        /// </returns>
        protected DataTable GetSimpleSettingsPrintTable(int columnCount)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnCount);

            try
            {
                return this.GetSimpleSettingsPrintTable(columnCount, 0);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnCount);
            }
        }

        /// <summary>
        /// Gets the simple settings print table.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <param name="keyIndex">The zero-based index of the primary key column.</param>
        /// <returns>
        /// The simple settings print table.
        /// </returns>
        protected DataTable GetSimpleSettingsPrintTable(int columnCount, int keyIndex)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'. Key Index: '{1}'.", columnCount, keyIndex);

            try
            {
                var printTable = Documenter.GetPrintTable();

                for (var i = 0; i < columnCount; ++i)
                {
                    printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", i }, { "Hidden", false }, { "SortOrder", (i == keyIndex) ? 0 : -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                }

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'. Key Index: '{1}'.", columnCount, keyIndex);
            }
        }

        /// <summary>
        /// Creates the simple settings difference gram.
        /// </summary>
        protected void CreateSimpleSettingsDiffgram()
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

        #endregion Simple Settings Sections

        #region Simple Ordered Settings Sections

        /// <summary>
        /// Creates the simple ordered settings data sets.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        protected void CreateSimpleOrderedSettingsDataSets(int columnCount)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnCount);

            try
            {
                this.CreateSimpleOrderedSettingsDataSets(columnCount, 2, false);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnCount);
            }
        }

        /// <summary>
        /// Creates the simple ordered settings data sets.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <param name="alphabeticOrder">if set to <c>true</c>, the rows are sorted alphabetically.</param>
        protected void CreateSimpleOrderedSettingsDataSets(int columnCount, bool alphabeticOrder)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnCount);

            try
            {
                this.CreateSimpleOrderedSettingsDataSets(columnCount, 2, alphabeticOrder);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnCount);
            }
        }

        /// <summary>
        /// Creates the simple ordered settings data sets
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <param name="keyCount">The key count.</param>
        /// <param name="alphabeticOrder">if set to <c>true</c>, the rows are sorted alphabetically.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSimpleOrderedSettingsDataSets(int columnCount, int keyCount, bool alphabeticOrder)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'. Primary Key Count: '{1}'.", columnCount, keyCount);

            try
            {
                var table = new DataTable("SimpleOrderedSettings") { Locale = CultureInfo.InvariantCulture };

                if (alphabeticOrder)
                {
                    table.Columns.Add(new DataColumn("Column1", typeof(string)));
                }
                else
                {
                    table.Columns.Add(new DataColumn("Column1", typeof(int)));
                }

                for (var i = 1; i < columnCount; ++i)
                {
                    table.Columns.Add(new DataColumn("Column" + (i + 1)));
                }

                var primaryKey = new List<DataColumn>(keyCount);
                for (var i = 0; i < keyCount; ++i)
                {
                    primaryKey.Add(table.Columns[i]);
                }

                table.PrimaryKey = primaryKey.ToArray();

                this.PilotDataSet = new DataSet("SimpleOrderedSettings") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetSimpleOrderedSettingsPrintTable(columnCount);
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'. Primary Key Count: '{1}'.", columnCount, keyCount);
            }
        }

        /// <summary>
        /// Gets the simple ordered settings print table.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <returns>
        /// The simple ordered settings print table.
        /// </returns>
        protected DataTable GetSimpleOrderedSettingsPrintTable(int columnCount)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnCount);

            try
            {
                var printTable = Documenter.GetPrintTable();

                for (var i = 0; i < columnCount; ++i)
                {
                    printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", i }, { "Hidden", i == 0 }, { "SortOrder", (i == 0) ? 0 : -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                }

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnCount);
            }
        }

        /// <summary>
        /// Creates the simple ordered settings difference gram.
        /// </summary>
        protected void CreateSimpleOrderedSettingsDiffgram()
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

        #endregion Simple Ordered Settings Sections
    }
}
