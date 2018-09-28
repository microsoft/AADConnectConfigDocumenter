//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// Azure AD Connect Sync Configuration Documenter Main Program
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace AzureADConnectConfigDocumenter
{
    using System;
    using System.Globalization;
    using System.Reflection;

    /// <summary>
    /// Azure AD Connect Sync Configuration Documenter Entry Point
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Azure AD Connect Sync Configuration Documenter Entry Point.
        /// </summary>
        /// <param name="args">The command-line arguments.</param>
        public static void Main(string[] args)
        {
            try
            {
                if (args == null || args.Length < 2)
                {
                    var errorMsg = string.Format(CultureInfo.CurrentUICulture, "Missing commnad-line arguments. Usage: {0} {1} {2}.", new object[] { Assembly.GetExecutingAssembly().GetName().Name, "{Pilot / Target Config Folder}", "{Production / Reference / Baseline Config Folder}" });
                    Console.Error.WriteLine(errorMsg);

                    errorMsg = string.Format(CultureInfo.CurrentUICulture, "Example: \t{0} {1} {2} {3}", new object[] { Environment.NewLine, Assembly.GetExecutingAssembly().GetName().Name, "\"Contoso\\Pilot\"", "\"Contoso\\Production\"" });
                    Console.Error.WriteLine(errorMsg);

                    Console.ReadKey();
                    return;
                }

                var documenter = new AzureADConnectSyncDocumenter(args[0], args[1]);
                documenter.GenerateReport();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }
    }
}
