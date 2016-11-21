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
        /// <param name="productionOnly">If set to <c>true</c>, indicates the connector is present in production only.</param>
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
                this.ResetDiffgram(); // reset the diffgram variables
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
                this.CreateActiveDirectoryPartitionContainersDataSets();

                this.FillActiveDirectoryPartitionContainersDataSet(partitionName, true);
                this.FillActiveDirectoryPartitionContainersDataSet(partitionName, false);

                this.CreateActiveDirectoryPartitionContainersDiffGram();

                this.PrintActiveDirectoryPartitionContainers();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Partition: '{0}'.", partitionName);
            }
        }

        #region Partition Settings

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
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
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

        #endregion Partition Settings

        #region Connection Options

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

        #endregion Connection Options

        #region Container Credential Settings

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
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Container Credential Settings

        #region Container Selections

        /// <summary>
        /// Creates the active directory partition containers data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        private void CreateActiveDirectoryPartitionContainersDataSets()
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

                var printTable = this.GetActiveDirectoryPartitionContainersPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the active directory partition containers print table.
        /// </summary>
        /// <returns>The active directory partition containers print table</returns>
        private DataTable GetActiveDirectoryPartitionContainersPrintTable()
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

                        var columnCount = table.Columns.Count;
                        foreach (var inclusion in inclusions)
                        {
                            var distinguishedName = (string)inclusion;
                            var row = this.GetContainerSelectionRow(distinguishedName, true, columnCount);
                            Documenter.AddRow(table, row);
                        }

                        var exclusions = partition.XPathSelectElements("filter/containers/exclusions/exclusion");

                        foreach (var exclusion in exclusions)
                        {
                            var distinguishedName = (string)exclusion;
                            var row = this.GetContainerSelectionRow(distinguishedName, false, columnCount);
                            Documenter.AddRow(table, row);
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
        private object[] GetContainerSelectionRow(string distinguishedName, bool inclusion, int columnCount)
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
        /// Creates the active directory partition containers diffgram.
        /// </summary>
        private void CreateActiveDirectoryPartitionContainersDiffGram()
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
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Container Selections

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
