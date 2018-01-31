//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="GenericSqlConnectorDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync Configuration Documenter Utility
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
    /// The GenericSqlConnectorDocumenter documents the configuration of a Generic SQL connector.
    /// </summary>
    internal class GenericSqlConnectorDocumenter : Extensible2ConnectorDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GenericSqlConnectorDocumenter"/> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        public GenericSqlConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
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
        /// Gets the active directory connector configuration report.
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
                this.ProcessGenericSqlSchemaInformation();

                this.ProcessExtensible2GlobalParameters();
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

        #region Schema Information

        /// <summary>
        /// Processes the Generic SQL schema information.
        /// </summary>
        protected void ProcessGenericSqlSchemaInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Schema Information");

                for (int i = 1; i <= 5; ++i)
                {
                    this.ProcessGenericSqlSchemaPage(i);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the Generic SQL schema page.
        /// </summary>
        /// <param name="pageNumber">The page number.</param>
        protected void ProcessGenericSqlSchemaPage(int pageNumber)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Schema Information");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1= Setting Display Order, 2 = Setting, 3 = Configuration

                this.FillGenericSqlSchemaPageDataSet(true, pageNumber);
                this.FillGenericSqlSchemaPageDataSet(false, pageNumber);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintGenericSqlSchemaPage(pageNumber);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the Generic SQL schema page data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        /// <param name="pageNumber">The page number.</param>
        protected void FillGenericSqlSchemaPageDataSet(bool pilotConfig, int pageNumber)
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

                    var parameterDefinitions = connector.XPathSelectElements("private-configuration/MAConfig/parameter-definitions/parameter[use = 'schema' and type != 'label' and type != 'divider' and page-number = '" + pageNumber + "']");

                    var parameterIndex = -1;
                    foreach (var parameterDefinition in parameterDefinitions)
                    {
                        ++parameterIndex;
                        var parameterName = (string)parameterDefinition.Element("name");
                        var parameter = connector.XPathSelectElement("private-configuration/MAConfig/parameter-values/parameter[@use = 'schema' and @name = '" + parameterName + "' and @page-number = '" + pageNumber + "']");
                        var encrypted = (string)parameter.Attribute("encrypted") == "1";
                        Documenter.AddRow(table, new object[] { parameterIndex, (string)parameter.Attribute("name"), encrypted ? "******" : (string)parameter });
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
        /// Prints the Generic SQL Schema information.
        /// </summary>
        /// <param name="pageNumber">The page number</param>
        protected void PrintGenericSqlSchemaPage(int pageNumber)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var sectionTitle = "Schema " + pageNumber;

                    this.WriteSectionHeader(sectionTitle, 4);

                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 30 }, { "Configuration", 70 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Schema Information
    }
}
