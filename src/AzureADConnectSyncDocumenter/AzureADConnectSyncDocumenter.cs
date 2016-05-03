//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="AzureADConnectSyncDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync Configuration Documenter Utility
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace AzureADConnectSyncDocumenter
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Web.UI;
    using System.Xml.XPath;

    /// <summary>
    /// The AzureADConnectSyncDocumenter documents the configuration of an Azure AD Connect sync deployment.
    /// </summary>
    public class AzureADConnectSyncDocumenter : Documenter
    {
        /// <summary>
        /// The current pilot / test configuration directory.
        /// This is the revised / target configuration which has introduces new changes to the baseline / production environment. 
        /// </summary>
        private string pilotConfigDirectory;

        /// <summary>
        /// The current production configuration directory.
        /// The is the baseline / reference configuration on which the changes will be reported.
        /// </summary>
        private string productionConfigDirectory;

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
                this.ReportFileName = Documenter.GetTempFilePath("Report.tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath("Report.TOC.tmp.html");

                this.ValidateInput(targetSystem, referenceSystem);
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

                using (var outputFile = new StreamWriter(this.configReportFilePath))
                {
                    outputFile.WriteLine(report.Item1.Replace("##TOC##", report.Item2));
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the AADSync configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            string report;
            string toc;

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                this.ReportWriter.WriteFullBeginTag("html");

                Documenter.WriteReportHeader(this.ReportWriter);

                this.ReportWriter.WriteFullBeginTag("body");

                this.ReportWriter.WriteFullBeginTag("h1");
                this.ReportWriter.Write("AADSync Service Configuration");
                this.ReportWriter.WriteEndTag("h1");
                this.ReportWriter.WriteLine();

                Documenter.WriteDocumenterInfo(this.ReportWriter);

                this.ReportWriter.WriteFullBeginTag("h1");
                this.ReportWriter.Write("Table of Contents");
                this.ReportWriter.WriteEndTag("h1");
                this.ReportWriter.WriteLine();

                this.ReportWriter.WriteLine("##TOC##");

                var sectionTitle = "AADSync Service Configuration";
                this.ReportToCWriter.WriteBeginTag("span");
                this.ReportToCWriter.WriteAttribute("class", "toc1");
                this.ReportToCWriter.Write(HtmlTextWriter.TagRightChar);
                Documenter.WriteJumpToBookmarkLocation(this.ReportToCWriter, sectionTitle, null, "TOC");
                this.ReportToCWriter.WriteEndTag("span");
                this.ReportToCWriter.WriteBeginTag("br");
                this.ReportToCWriter.Write(HtmlTextWriter.SelfClosingTagEnd);
                this.ReportToCWriter.WriteLine();

                this.ReportWriter.WriteFullBeginTag("h1");
                Documenter.WriteBookmarkLocation(this.ReportWriter, sectionTitle, null, "TOC");
                this.ReportWriter.WriteEndTag("h1");
                this.ReportWriter.WriteLine();

                this.ProcessMetaverseConfiguration();
                this.ProcessConnectorConfigurations();
            }
            catch (Exception e)
            {
                throw Logger.Instance.ReportError(e);
            }
            finally
            {
                this.ReportWriter.WriteEndTag("body");

                this.ReportWriter.WriteEndTag("html");

                this.ReportWriter.Close();
                this.ReportToCWriter.Close();

                using (var reportReader = new StreamReader(this.ReportFileName))
                {
                    report = reportReader.ReadToEnd();
                    using (var tocReader = new StreamReader(this.ReportToCFileName))
                    {
                        toc = tocReader.ReadToEnd();
                    }
                }

                Logger.Instance.WriteMethodExit();
            }

            return new Tuple<string, string>(report, toc);
        }

        /// <summary>
        /// Validates the input.
        /// </summary>
        /// <param name="targetSystem">The target system.</param>
        /// <param name="referenceSystem">The reference system.</param>
        private void ValidateInput(string targetSystem, string referenceSystem)
        {
            Logger.Instance.WriteMethodEntry("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", targetSystem, referenceSystem);

            try
            {
                var rootDirectory = Directory.GetCurrentDirectory().TrimEnd('\\');

                this.pilotConfigDirectory = string.Format(CultureInfo.InvariantCulture, @"{0}\Data\{1}", rootDirectory, targetSystem);
                this.productionConfigDirectory = string.Format(CultureInfo.InvariantCulture, @"{0}\Data\{1}", rootDirectory, referenceSystem);

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

                this.configReportFilePath = Documenter.ReportFolder + @"\" + (targetSystem ?? string.Empty).Replace(@"\", "_") + "_To_" + (referenceSystem ?? string.Empty).Replace(@"\", "_") + "_AADSync_report.html";
            }
            finally
            {
                Logger.Instance.WriteMethodExit("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", targetSystem, referenceSystem);
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
                this.PilotXml = Documenter.MergeConfigurationExports(this.pilotConfigDirectory, true);
                this.ProductionXml = Documenter.MergeConfigurationExports(this.productionConfigDirectory, false);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the metaverse configuration.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
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
                const string XPath = "//ma-data";

                var pilot = this.PilotXml.XPathSelectElements(XPath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(XPath, Documenter.NamespaceManager);

                // Sort by name
                pilot = from connector in pilot
                        let name = (string)connector.Element("name")
                        orderby name
                        select connector;

                foreach (var connector in pilot)
                {
                    var connectorName = (string)connector.Element("name");
                    var connectorCategory = (string)connector.Element("category");

                    switch (connectorCategory.ToUpperInvariant())
                    {
                        case "AD":
                            {
                                var connectorDocumenter = new ActiveDirectoryConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, false);
                                var report = connectorDocumenter.GetReport();
                                this.ReportWriter.Write(report.Item1);
                                this.ReportToCWriter.Write(report.Item2);
                            }

                            break;

                        default:
                            {
                                var connectorDocumenter = new Extensible2ConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, false);
                                var report = connectorDocumenter.GetReport();
                                this.ReportWriter.Write(report.Item1);
                                this.ReportToCWriter.Write(report.Item2);
                            }

                            break;
                    }
                }

                // Sort by name
                production = from connector in production
                             let name = (string)connector.Element("name")
                             orderby name
                             select connector;

                var pilotConnectors = from pilotConnector in pilot
                                      select (string)pilotConnector.Element("name");

                production = production.Where(productionConnector => !pilotConnectors.Contains((string)productionConnector.Element("name")));

                foreach (var connector in production)
                {
                    var connectorName = (string)connector.Element("name");
                    var connectorCategory = (string)connector.Element("category");

                    switch (connectorCategory.ToUpperInvariant())
                    {
                        case "AD":
                            {
                                var connectorDocumenter = new ActiveDirectoryConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, true);
                                var report = connectorDocumenter.GetReport();
                                this.ReportWriter.Write(report.Item1);
                                this.ReportToCWriter.Write(report.Item2);
                            }

                            break;

                        default:
                            {
                                var connectorDocumenter = new Extensible2ConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, true);
                                var report = connectorDocumenter.GetReport();
                                this.ReportWriter.Write(report.Item1);
                                this.ReportToCWriter.Write(report.Item2);
                            }

                            break;
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }
    }
}
