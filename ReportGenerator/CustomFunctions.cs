// system
using System;
using System.Text.RegularExpressions;

// stimulsoft
using Stimulsoft.Report.Dictionary;


namespace ReportGenerator
{
    public static class CustomFunctions
    {
        /// <summary>
        /// Registers custom functions in the Stimulsoft functions list
        /// </summary>
        public static void RegisterFunctions()
        {
            StiFunctions.AddFunction(
                "Custom",                                           // Category name
                "CalculateMinorQuantityBySquareMeter",              // Function name
                "Calculate quantity per square meter",              // Description
                typeof(CustomFunctions),                            // Type containing the function
                typeof(double),                                     // Return type
                "Returns quantity divided by area (m²)",            // Return description
                new[] { typeof(string), typeof(double) },           // Parameter types
                new[] { "info", "quantity" },                       // Parameter names
                new[] { "Size like 60*30 (cm)", "Quantity value" }  // Parameter descriptions
            );
        }

        /// <summary>
        /// Calculates the count of product based on the square meter data and the major quantity
        /// of the product.
        /// </summary>
        /// <param name="info">a string containing 'number1 * number2' pattern</param>
        /// <param name="quantity">major quantity of the product</param>
        /// <returns>the minor quantity of the product</returns>
        public static double CalculateMinorQuantityBySquareMeter(string info, double quantity)
        {
            if (string.IsNullOrWhiteSpace(info)) return 0;

            string pattern = @"\d+\s*\*\s*\d+";
            info = Regex.Match(info, pattern).Value;

            string[] splitList = info.Trim().Split('*');
            if (splitList.Length != 2) return 0;

            double height;
            double width;
            try
            {
                width = Convert.ToDouble(splitList[0].Trim());
                height = Convert.ToDouble(splitList[1].Trim());
            }
            catch
            {
                return 0;
            }

            double widthInMeters = width / 100.0;
            double heightInMeters = height / 100.0;
            double squareMeter = widthInMeters * heightInMeters;
            if (squareMeter < 0.0) return 0;

            double res = quantity / squareMeter;
            return Math.Round(res, 2);
        }
    }
}
