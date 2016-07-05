//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ActiveDirectoryConnectorDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync Configuration Documenter
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
    /// The ActiveDirectoryConnectorDocumenter documents the configuration of Active Directory connector.
    /// </summary>
    internal sealed class ActiveDirectoryConnectorDocumenter : ConnectorDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ActiveDirectoryConnectorDocumenter"/> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The name.</param>
        /// <param name="productionOnly">if set to <c>true</c> [production only].</param>
        public ActiveDirectoryConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, bool productionOnly)
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
        /// Gets the active directory connector configuration report.
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

                this.ProcessActiveDirectoryConnectionInformation();
                this.ProcessActiveDirectoryPartitions();
                this.ProcessConnectorProvisioningHierarchyConfiguration();
                this.ProcessConnectorSelectedObjectTypes();
                this.ProcessConnectorSelectedAttributes();
                this.ProcessConnectorProvisioningSyncRules();
                this.ProcessConnectorStickyJoinSyncRules();
                this.ProcessConnectorNormalJoinSyncRules();
                this.ProcessConnectorSyncRules();
                this.ProcessActiveDirectoryRunProfiles();

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

        #region AD Connection Information

        /// <summary>
        /// Processes the active directory connection information.
        /// </summary>
        private void ProcessActiveDirectoryConnectionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Active Directory Connection Information.");

                // Forest Connection Information
                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order, 2 = Setting, 3 = Configuration

                this.FillActiveDirectoryConnectionInformationDataSet(true);
                this.FillActiveDirectoryConnectionInformationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintActiveDirectoryConnectionInformation();

                // Forest Connection Option
                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order, 2 = Setting, 3 = Configuration

                this.FillActiveDirectoryConnectionOptionDataSet(true);
                this.FillActiveDirectoryConnectionOptionDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintActiveDirectoryConnectionOption();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region AD Forest Information

        /// <summary>
        /// Fills the active directory connection information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillActiveDirectoryConnectionInformationDataSet(bool pilotConfig)
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

                    var forestName = (string)connector.XPathSelectElement("private-configuration/adma-configuration/forest-name");
                    var userName = (string)connector.XPathSelectElement("private-configuration/adma-configuration/forest-login-user");
                    var userDomain = (string)connector.XPathSelectElement("private-configuration/adma-configuration/forest-login-domain");

                    Documenter.AddRow(table, new object[] { 1, "Forest Name", forestName });
                    Documenter.AddRow(table, new object[] { 2, "User Name", userName });
                    Documenter.AddRow(table, new object[] { 3, "Domain", userDomain });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the active directory connection information.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintActiveDirectoryConnectionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Forest Connection Information";

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
                this.ReportWriter.WriteAttribute("class", "outer-table");
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
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion AD Forest Information

        #region AD Connection Option

        /// <summary>
        /// Fills the active directory connection option data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillActiveDirectoryConnectionOptionDataSet(bool pilotConfig)
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

                    var signAndSeal = connector.XPathSelectElement("private-configuration/adma-configuration/sign-and-seal");

                    if ((string)signAndSeal == "1")
                    {
                        Documenter.AddRow(table, new object[] { 1, "Sign and Encrypt LDAP traffic", "Yes" });
                    }
                    else
                    {
                        var sslBind = connector.XPathSelectElement("private-configuration/adma-configuration/ssl-bind");

                        if ((string)sslBind == "1")
                        {
                            Documenter.AddRow(table, new object[] { 2, "Enable SSL for the connection", "Yes" });

                            var crlCheck = sslBind.Attribute("crl-check");

                            var crlCheckEnabled = (string)crlCheck == "1" ? "Yes" : "No";
                            Documenter.AddRow(table, new object[] { 3, "Enable CRL Checking", crlCheckEnabled });
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
        /// Prints the active directory connection option.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintActiveDirectoryConnectionOption()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table");
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
                            this.ReportWriter.Write("Connection Options");
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
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion AD Connection Option

        #endregion AD Connection Information

        #region AD Partitions

        /// <summary>
        /// Processes the metaverse object types.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessActiveDirectoryPartitions()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Active Directory Partitions.");

                var sectionTitle = "Partitions Information";

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

                var xpath = "//ma-data[name ='" + this.ConnectorName + "']" + "//ma-partition-data/partition[selected = 1]";

                var pilot = this.PilotXml.XPathSelectElements(xpath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(xpath, Documenter.NamespaceManager);

                // Sort by name
                var pilotPartitions = from partition in pilot
                                      let name = (string)partition.Element("name")
                                      orderby name
                                      select name;

                foreach (var partition in pilotPartitions)
                {
                    this.ProcessActiveDirectoryPartition(partition);
                }

                // Sort by name
                var productionPartitions = from partition in production
                                           let name = (string)partition.Element("name")
                                           orderby name
                                           select name;

                productionPartitions = productionPartitions.Where(productionPartition => !pilotPartitions.Contains(productionPartition));

                foreach (var partition in productionPartitions)
                {
                    this.ProcessActiveDirectoryPartition(partition);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the active directory partition.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessActiveDirectoryPartition(string partitionName)
        {
            Logger.Instance.WriteMethodEntry("Partition: '{0}'.", partitionName);

            try
            {
                #region toc

                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc4");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, partitionName, this.ConnectorGuid, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                #endregion toc

                #region section

                this.ReportWriter.WriteFullBeginTag("h4");
                Documenter.WriteBookmarkLocation(this.ReportWriter, partitionName, "Partition: " + partitionName, this.ConnectorGuid, "TOC");
                this.ReportWriter.WriteEndTag("h4");

                #endregion section

                // Partion Settings
                this.CreateActiveDirectoryPartitionSettingsDataSets();

                this.FillActiveDirectoryPartitionSettingsDataSet(partitionName, true);
                this.FillActiveDirectoryPartitionSettingsDataSet(partitionName, false);

                this.CreateActiveDirectoryPartitionSettingsDiffGram();

                this.PrintActiveDirectoryPartitionSettings();

                // Partition Connection Option
                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order, 2 = Setting, 3 = Configuration

                this.FillActiveDirectoryConnectionOptionDataSet(partitionName, true);
                this.FillActiveDirectoryConnectionOptionDataSet(partitionName, false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintActiveDirectoryConnectionOption();

                // Container Credential Settings
                this.CreateSimpleSettingsDataSets(2); // 1 = Setting, 2 = Configuration

                this.FillActiveDirectoryContainerCredentialSettingsDataSet(partitionName, true);
                this.FillActiveDirectoryContainerCredentialSettingsDataSet(partitionName, false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintActiveDirectoryContainerCredentialSettings();
                
                // Container Include / Exclude settings
                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order, 2 = Container, 3 = Include / Exclude

                this.FillActiveDirectoryPartitionContainersDataSet(partitionName, true);
                this.FillActiveDirectoryPartitionContainersDataSet(partitionName, false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintActiveDirectoryPartitionContainers();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Partition: '{0}'.", partitionName);
            }
        }

        /// <summary>
        /// Creates the active directory partition settings data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        private void CreateActiveDirectoryPartitionSettingsDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("PartitionSettings") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("SettingDisplayOrder", typeof(int));
                var column2 = new DataColumn("Setting");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("PartitionSettingConfigurations") { Locale = CultureInfo.InvariantCulture };
                var column12 = new DataColumn("Setting");
                var column22 = new DataColumn("ConfigurationOrder", typeof(int));
                var column32 = new DataColumn("Configuration");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.PrimaryKey = new[] { column12, column22 };

                this.PilotDataSet = new DataSet("PartitionSettings") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column2 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetActiveDirectoryPartitionSettingsPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the active directory partition settings print table.
        /// </summary>
        /// <returns>The active directory partition settings print table</returns>
        private DataTable GetActiveDirectoryPartitionSettingsPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Setting Display Order
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Setting
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Setting
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Configuration Order
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Configuration
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the active directory partition settings data set.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillActiveDirectoryPartitionSettingsDataSet(string partitionName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");
                if (connector != null)
                {
                    var partition = connector.XPathSelectElement("ma-partition-data/partition[selected = 1 and name = '" + partitionName + "']");

                    if (partition != null)
                    {
                        var table = dataSet.Tables[0];
                        var table2 = dataSet.Tables[1];

                        // Preferred Domain Controllers
                        Documenter.AddRow(table, new object[] { 0, "Preferred Domain Controllers" });
                        var preferredDomainControllers = partition.XPathSelectElements("custom-data/adma-partition-data/preferred-dcs/preferred-dc");

                        for (var pdcIndex = 0; pdcIndex < preferredDomainControllers.Count(); ++pdcIndex)
                        {
                            Documenter.AddRow(table2, new object[] { "Preferred Domain Controllers", pdcIndex, (string)preferredDomainControllers.ElementAt(pdcIndex) });
                        }

                        // Last Domain Controller used
                        var lastDomainControllerUsed = (string)partition.XPathSelectElement("custom-data/adma-partition-data/last-dc");
                        Documenter.AddRow(table, new object[] { 1, "Last domain controller used" });
                        Documenter.AddRow(table2, new object[] { "Last domain controller used", 0, lastDomainControllerUsed });

                        table.AcceptChanges();
                        table2.AcceptChanges();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);
            }
        }

        /// <summary>
        /// Creates the active directory partition settings diffgram.
        /// </summary>
        private void CreateActiveDirectoryPartitionSettingsDiffGram()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = Documenter.GetDiffgram(this.PilotDataSet, this.ProductionDataSet);

                // Setup data relations
                var column2 = this.DiffgramDataSet.Tables[0].Columns["Setting"];

                var column12 = this.DiffgramDataSet.Tables[1].Columns["Setting"];

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column2 }, new[] { column12 }, false);

                this.DiffgramDataSet.Relations.Add(dataRelation12);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the active directory partition settings.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintActiveDirectoryPartitionSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table");
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
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the active directory connection option data set.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillActiveDirectoryConnectionOptionDataSet(string partitionName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var partition = connector.XPathSelectElement("ma-partition-data/partition[selected = 1 and name = '" + partitionName + "']");

                    if (partition != null)
                    {
                        var table = dataSet.Tables[0];

                        var signAndSeal = partition.XPathSelectElement("custom-data/adma-partition-data/sign-and-seal");

                        if ((string)signAndSeal == "1")
                        {
                            Documenter.AddRow(table, new object[] { 1, "Sign and Encrypt LDAP traffic", "Yes" });
                        }
                        else
                        {
                            var sslBind = partition.XPathSelectElement("custom-data/adma-partition-data/ssl-bind");

                            if ((string)sslBind == "1")
                            {
                                Documenter.AddRow(table, new object[] { 2, "Enable SSL for the connection", "Yes" });

                                var crlCheck = sslBind.Attribute("crl-check");

                                var crlCheckEnabled = (string)crlCheck == "1" ? "Yes" : "No";
                                Documenter.AddRow(table, new object[] { 3, "Enable CRL Checking", crlCheckEnabled });
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
        /// Fills the active directory container credential settings data set.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillActiveDirectoryContainerCredentialSettingsDataSet(string partitionName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var partition = connector.XPathSelectElement("ma-partition-data/partition[selected = 1 and name = '" + partitionName + "']");

                    if (partition != null)
                    {
                        var table = dataSet.Tables[0];

                        var loginUser = (string)partition.XPathSelectElement("custom-data/adma-partition-data/login-user");
                        var loginDomain = (string)partition.XPathSelectElement("custom-data/adma-partition-data/login-domain");

                        if (string.IsNullOrEmpty(loginUser))
                        {
                            Documenter.AddRow(table, new object[] { "Use default forest credentials", "Yes" });
                        }
                        else
                        {
                            Documenter.AddRow(table, new object[] { "Alternate credentials for this directory partition", loginDomain + @"\" + loginUser });
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
        /// Prints the active directory container credential settings.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintActiveDirectoryContainerCredentialSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table");
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
                            this.ReportWriter.Write("Credentials");
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
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the active directory partition containers data set.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillActiveDirectoryPartitionContainersDataSet(string partitionName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var partition = connector.XPathSelectElement("ma-partition-data/partition[selected = 1 and name = '" + partitionName + "']");

                    if (partition != null)
                    {
                        var table = dataSet.Tables[0];

                        var inclusions = partition.XPathSelectElements("filter/containers/inclusions/inclusion");

                        for (var index = 0; index < inclusions.Count(); ++index)
                        {
                            var inclusion = (string)inclusions.ElementAt(index);
                            Documenter.AddRow(table, new object[] { index, inclusion, "Include" });
                        }

                        var exclusions = partition.XPathSelectElements("filter/containers/exclusions/exclusion");

                        for (var index = inclusions.Count(); index < exclusions.Count(); ++index)
                        {
                            var exclusion = (string)exclusions.ElementAt(index);
                            Documenter.AddRow(table, new object[] { index, exclusion, "Exclude" });
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
        /// Prints the active directory partition containers.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintActiveDirectoryPartitionContainers()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                #region table

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", "outer-table");
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
                            this.ReportWriter.Write("Container");
                            this.ReportWriter.WriteEndTag("th");

                            this.ReportWriter.WriteBeginTag("th");
                            this.ReportWriter.WriteAttribute("class", "column-th");
                            this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                            this.ReportWriter.Write("Include / Exclude");
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
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion AD Partitions

        #region Run Profiles

        /// <summary>
        /// Processes the connector run profiles.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessActiveDirectoryRunProfiles()
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
                    this.ProcessActiveDirectoryRunProfile(runProfile);
                }

                // Sort by name
                var productionRunProfiles = from runProfile in production
                                            let name = (string)runProfile.Element("name")
                                            orderby name
                                            select name;

                productionRunProfiles = productionRunProfiles.Where(productionRunProfile => !pilotRunProfiles.Contains(productionRunProfile));

                foreach (var runProfile in productionRunProfiles)
                {
                    this.ProcessActiveDirectoryRunProfile(runProfile);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the Active Directory run profile.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void ProcessActiveDirectoryRunProfile(string runProfileName)
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

                this.FillActiveDirectoryRunProfileDataSet(runProfileName, true);
                this.FillActiveDirectoryRunProfileDataSet(runProfileName, false);

                this.CreateConnectorRunProfileDiffgram();

                this.PrintConnectorRunProfile();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Run Profile Name: '{0}'.", runProfileName);
            }
        }

        /// <summary>
        /// Fills the Active Directory run profile data set.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillActiveDirectoryRunProfileDataSet(string runProfileName, bool pilotConfig)
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

                        var batchSize = (string)runProfileStep.XPathSelectElement("custom-data/adma-step-data/batch-size");
                        if (!string.IsNullOrEmpty(batchSize))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Batch Size (objects)", batchSize, 7 });
                        }

                        var pageSize = (string)runProfileStep.XPathSelectElement("custom-data/adma-step-data/page-size");
                        if (!string.IsNullOrEmpty(pageSize))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Page Size (objects)", pageSize, 8 });
                        }

                        var timeout = (string)runProfileStep.XPathSelectElement("custom-data/adma-step-data/time-limit");
                        if (!string.IsNullOrEmpty(timeout))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Timeout (seconds)", timeout, 9 });
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
