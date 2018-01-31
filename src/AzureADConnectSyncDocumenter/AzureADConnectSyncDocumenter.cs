//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="AzureADConnectSyncDocumenter.cs" company="Microsoft">
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
    using System.IO;
    using System.Linq;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The AzureADConnectSyncDocumenter documents the configuration of an Azure AD Connect sync deployment.
    /// </summary>
    public class AzureADConnectSyncDocumenter : Documenter
    {
        /// <summary>
        /// The current pilot / test configuration directory.
        /// This is the revised / target configuration which has introduced new changes to the baseline / production environment. 
        /// </summary>
        private string pilotConfigDirectory;

        /// <summary>
        /// The current production configuration directory.
        /// The is the baseline / reference configuration on which the changes will be reported.
        /// </summary>
        private string productionConfigDirectory;

        /// <summary>
        /// The relative path of the current pilot / test configuration directory.
        /// </summary>
        private string pilotConfigRelativePath;

        /// <summary>
        /// The relative path of the current production configuration directory.
        /// </summary>
        private string productionConfigRelativePath;

        /// <summary>
        /// The configuration report file path
        /// </summary>
        private string configReportFilePath;

        /// <summary>
        /// Initializes a new instance of the <see cref="AzureADConnectSyncDocumenter"/> class.
        /// </summary>
        /// <param name="targetSystem">The target / pilot / test system.</param>
        /// <param name="referenceSystem">The reference / baseline / production system.</param>
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "GlobalSettings", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "SynchronizationRules", Justification = "Reviewed.")]
        public AzureADConnectSyncDocumenter(string targetSystem, string referenceSystem)
        {
            Logger.Instance.WriteMethodEntry("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", targetSystem, referenceSystem);

            try
            {
                this.pilotConfigRelativePath = targetSystem;
                this.productionConfigRelativePath = referenceSystem;
                this.ReportFileName = Documenter.GetTempFilePath("Report.tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath("Report.TOC.tmp.html");

                var rootDirectory = Directory.GetCurrentDirectory().TrimEnd('\\');

                this.pilotConfigDirectory = string.Format(CultureInfo.InvariantCulture, @"{0}\Data\{1}", rootDirectory, this.pilotConfigRelativePath);
                this.productionConfigDirectory = string.Format(CultureInfo.InvariantCulture, @"{0}\Data\{1}", rootDirectory, this.productionConfigRelativePath);
                this.configReportFilePath = Documenter.ReportFolder + @"\" + (this.pilotConfigRelativePath ?? string.Empty).Replace(@"\", "_") + "_AppliedTo_" + (this.productionConfigRelativePath ?? string.Empty).Replace(@"\", "_") + "_AADConnectSync_report.html";

                this.ValidateInput();
                this.MergeSyncExports();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", targetSystem, referenceSystem);
            }
        }

        /// <summary>
        /// Generates the report.
        /// </summary>
        public void GenerateReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var report = this.GetReport();
                this.WriteReport("AAD Connect Sync Service Configuration", report.Item1, report.Item2, this.pilotConfigRelativePath, this.productionConfigRelativePath, this.configReportFilePath);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the AAD Connect sync configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            Tuple<string, string> report;

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                var sectionTitle = "AAD Connect Sync Service Configuration";

                this.WriteSectionHeader(sectionTitle, 1);

                this.ProcessGlobalSettings();
                this.ProcessMetaverseConfiguration();
                this.ProcessConnectorConfigurations();
            }
            catch (Exception e)
            {
                throw Logger.Instance.ReportError(e);
            }
            finally
            {
                report = this.GetReportTuple();

                Logger.Instance.WriteMethodExit();
            }

            return report;
        }

        /// <summary>
        /// Validates the input.
        /// </summary>
        private void ValidateInput()
        {
            Logger.Instance.WriteMethodEntry("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", this.pilotConfigRelativePath, this.productionConfigRelativePath);

            try
            {
                if (!Directory.Exists(this.pilotConfigDirectory))
                {
                    var error = string.Format(CultureInfo.CurrentUICulture, DocumenterResources.PilotConfigurationDirectoryNotFound, this.pilotConfigDirectory);
                    throw Logger.Instance.ReportError(new FileNotFoundException(error));
                }

                if (!Directory.Exists(this.productionConfigDirectory))
                {
                    var error = string.Format(CultureInfo.CurrentUICulture, DocumenterResources.ProductionConfigurationDirectoryNotFound, this.productionConfigDirectory);
                    throw Logger.Instance.ReportError(new FileNotFoundException(error));
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", this.pilotConfigRelativePath, this.productionConfigRelativePath);
            }
        }

        /// <summary>
        /// Merges the ADSync configuration export XML files into a single XML file.
        /// </summary>
        private void MergeSyncExports()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PilotXml = this.MergeConfigurationExports(true);
                this.ProductionXml = this.MergeConfigurationExports(false);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Merges the ADSync configuration exports.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <returns>
        /// An <see cref="XElement" /> object representing the combined configuration XML object.
        /// </returns>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Xml.Linq.XElement.Parse(System.String)", Justification = "Template XML is not localizable.")]
        private XElement MergeConfigurationExports(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var configDirectory = pilotConfig ? this.pilotConfigDirectory : this.productionConfigDirectory;
                var templateXml = string.Format(CultureInfo.InvariantCulture, "<Root><{0}><Connectors/><GlobalSettings/><SynchronizationRules/></{0}></Root>", pilotConfig ? "Pilot" : "Production");
                var configXml = XElement.Parse(templateXml);

                var connectors = configXml.XPathSelectElement("*//Connectors");
                foreach (var file in Directory.EnumerateFiles(configDirectory + "/Connectors", "*.xml"))
                {
                    connectors.Add(XElement.Load(file));
                }

                var globalSettings = configXml.XPathSelectElement("*//GlobalSettings");
                foreach (var file in Directory.EnumerateFiles(configDirectory + "/GlobalSettings", "*.xml"))
                {
                    globalSettings.Add(XElement.Load(file));
                }

                var synchronizationRules = configXml.XPathSelectElement("*//SynchronizationRules");
                foreach (var file in Directory.EnumerateFiles(configDirectory + "/SynchronizationRules", "*.xml"))
                {
                    synchronizationRules.Add(XElement.Load(file));
                }

                return configXml;
            }
            finally
            {
                Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        #region Global Settings

        /// <summary>
        /// Processes the global settings configuration.
        /// </summary>
        private void ProcessGlobalSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Global Settings.");

                this.CreateSimpleSettingsDataSets(2);  // 1 = Name, 2 = Value

                this.FillGlobalSettingsDataSet(true);
                this.FillGlobalSettingsDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintGlobalSettings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the global settings data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillGlobalSettingsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var parameters = config.XPathSelectElements(Documenter.GetMetaverseXmlRootXPath(pilotConfig) + "/mv-data//parameter-values/parameter");

                // Sort by name
                parameters = from parameter in parameters
                             let name = (string)parameter.Attribute("name")
                             orderby name
                             select parameter;

                for (var parameterIndex = 0; parameterIndex < parameters.Count(); ++parameterIndex)
                {
                    var parameter = parameters.ElementAt(parameterIndex);
                    Documenter.AddRow(table, new object[] { (string)parameter.Attribute("name"), (string)parameter });
                }

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the global settings.
        /// </summary>
        private void PrintGlobalSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Global Settings";

                this.WriteSectionHeader(sectionTitle, 2);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 50 }, { "Value", 50 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Global Settings

        /// <summary>
        /// Processes the metaverse configuration.
        /// </summary>
        private void ProcessMetaverseConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var metaverseDocumenter = new MetaverseDocumenter(this.PilotXml, this.ProductionXml);
                var report = metaverseDocumenter.GetReport();

                this.ReportWriter.Write(report.Item1);
                this.ReportToCWriter.Write(report.Item2);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the connector configurations.
        /// </summary>
        private void ProcessConnectorConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var pilot = this.PilotXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(true) + "/ma-data", Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(Documenter.GetConnectorXmlRootXPath(false) + "/ma-data", Documenter.NamespaceManager);

                // Sort by name
                pilot = from connector in pilot
                        let name = (string)connector.Element("name")
                        orderby name
                        select connector;

                // Sort by name
                production = from connector in production
                             let name = (string)connector.Element("name")
                             orderby name
                             select connector;

                var pilotConnectors = from pilotConnector in pilot
                                      select (string)pilotConnector.Element("name");

                foreach (var connector in pilot)
                {
                    var configEnvironment = production.Any(productionConnector => (string)productionConnector.Element("name") == (string)connector.Element("name")) ? ConfigEnvironment.PilotAndProduction : ConfigEnvironment.PilotOnly;
                    if (configEnvironment == ConfigEnvironment.PilotOnly)
                    {
                        Logger.Instance.WriteWarning("The connector '{0}' only exists in the first set of configuration files. If this is not expected, please edit the config files to match names of this connector in the two enviorments as instructed in the ReadMe wiki.", (string)connector.Element("name"));
                    }

                    this.ProcessConnectorConfiguration(connector, configEnvironment);
                }

                production = production.Where(productionConnector => !pilotConnectors.Contains((string)productionConnector.Element("name")));

                foreach (var connector in production)
                {
                    Logger.Instance.WriteWarning("The connector '{0}' only exists in the second set of configuration files and will be documented as a connector that is deleted from the configuration. If this is not intended, please edit the config files to match names of this connector in the two enviorments as instructed in the ReadMe wiki.", (string)connector.Element("name"));

                    this.ProcessConnectorConfiguration(connector, ConfigEnvironment.ProductionOnly);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the connector configuration of specified connector.
        /// </summary>
        /// <param name="connector"> The connector config node</param>
        /// <param name="configEnvironment">The config environment.</param>
        private void ProcessConnectorConfiguration(XElement connector, ConfigEnvironment configEnvironment)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var connectorName = (string)connector.Element("name");
                var connectorCategory = (string)connector.Element("category");
                ConnectorDocumenter connectorDocumenter;

                switch (connectorCategory.ToUpperInvariant())
                {
                    case "AD":
                        connectorDocumenter = new ActiveDirectoryConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                    case "EXTENSIBLE2":
                        {
                            var connectorSubType = (string)connector.Element("subtype");
                            switch (connectorSubType.ToUpperInvariant())
                            {
                                case "WINDOWS AZURE ACTIVE DIRECTORY (MICROSOFT)":
                                    connectorDocumenter = new AzureActiveDirectoryConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "POWERSHELL (MICROSOFT)":
                                    connectorDocumenter = new PowerShellConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "GENERIC SQL (MICROSOFT)":
                                    connectorDocumenter = new GenericSqlConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "GENERIC LDAP (MICROSOFT)":
                                    connectorDocumenter = new GenericLdapConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "MICROSOFT WEB SERVICE ECMA2 CONNECTOR (MICROSOFT)":
                                    connectorDocumenter = new WebServicesConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                default:
                                    if (!string.IsNullOrEmpty(connectorSubType))
                                    {
                                        Logger.Instance.WriteWarning("ECMA2 Connector of subtype '{0}' is currently not supported. The connector '{1}' with be treated as a generic ECMA2 connector.", connectorSubType, connectorName);
                                    }

                                    connectorDocumenter = new Extensible2ConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                            }
                        }

                        break;
                    default:
                        Logger.Instance.WriteWarning("Connector of type '{0}' is currently not supported. The connector '{1}' with be treated as a generic ECMA2 connector. Documentation may not be complete.", connectorCategory, connectorName);
                        connectorDocumenter = new Extensible2ConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                }

                var report = connectorDocumenter.GetReport();
                this.ReportWriter.Write(report.Item1);
                this.ReportToCWriter.Write(report.Item2);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }
    }
}
