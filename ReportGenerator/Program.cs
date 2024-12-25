// stimulsoft
using Azure;
using Stimulsoft.Base;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Export;

// system
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Drawing.Text;
using System.Text.RegularExpressions;


namespace ReportGenerator
{
    internal class Program
    {
        public const int DEFAULT_IMAGE_RESOLUTION = 300;
        /// <summary>
        /// Generate a report for the provided stimulsoft mrt file and relevant data
        /// </summary>
        /// <param name="args">
        /// args[0] = full path to the mrt file
        /// args[1] = string representation of the json data
        /// args[2] = type of the generated report
        /// args[3] = full path to the output file
        /// </param>
        /// <returns>an integer showing if the operation was successful(1) or not (other) </returns>
        static int Main(string[] args)
        {
            ReportEngine reportEngine = new();
            try
            {
                if (args.Length < 5)
                {
                    return 1;
                }

                // obtain input args
                string reportPath = args[0];
                string jsonData = args[1];
                int reportResolution = Convert.ToInt32(args[2]);
                string outputType = args[3];
                string outputPath = args[4];
                string dataSourceName = "Data";
                string licenseKey = "6vJhGtLLLz2GNviWmUTrhSqnOItdDwjBylQzQcAOiHkO46nMQvol4ASeg91in+mGJLnn2KMIpg3eSXQSgaFOm15+0l" +
                    "hekKip+wRGMwXsKpHAkTvorOFqnpF9rchcYoxHXtjNDLiDHZGTIWq6D/2q4k/eiJm9fV6FdaJIUbWGS3whFWRLPHWC" +
                    "BsWnalqTdZlP9knjaWclfjmUKf2Ksc5btMD6pmR7ZHQfHXfdgYK7tLR1rqtxYxBzOPq3LIBvd3spkQhKb07LTZQoyQ" +
                    "3vmRSMALmJSS6ovIS59XPS+oSm8wgvuRFqE1im111GROa7Ww3tNJTA45lkbXX+SocdwXvEZyaaq61Uc1dBg+4uFRxv" +
                    "yRWvX5WDmJz1X0VLIbHpcIjdEDJUvVAN7Z+FW5xKsV5ySPs8aegsY9ndn4DmoZ1kWvzUaz+E1mxMbOd3tyaNnmVhPZ" +
                    "eIBILmKJGN0BwnnI5fu6JHMM/9QR2tMO1Z4pIwae4P92gKBrt0MqhvnU1Q6kIaPPuG2XBIvAWykVeH2a9EP6064e11" +
                    "PFCBX4gEpJ3XFD0peE5+ddZh+h495qUc1H2B";

                object clearedData = reportEngine.CleanAndValidateData(jsonData, []);
                reportEngine.SetLicense([licenseKey]);
                reportEngine.LoadTemplate(reportPath, []);
                string[] renderArgs = [outputType, reportPath, dataSourceName, jsonData];
                reportEngine.RegisterData(dataSourceName, clearedData, []);
                reportEngine.Render(renderArgs);
                reportEngine.Export(outputType, outputPath, reportResolution, []);

                return 0;
            }
            catch
            {
                Console.WriteLine(reportEngine.GetLastError());
                return 1;
            }
        }
    }
}
