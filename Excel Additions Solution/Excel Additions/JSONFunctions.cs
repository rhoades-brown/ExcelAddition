using ExcelDna.Integration;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;

namespace Excel_Additions
{
    public static class JSONFunctions
    {
        #region Helper Functions

        private static void RecurseObject(
            JObject sourceObject,
            JToken propertyObject,
            string[] elements
        )
        {
            if (sourceObject.SelectToken(elements[0]) != null)
            {
                if (elements.Length == 1)
                {
                    sourceObject.Add(elements[0], propertyObject);
                }
                else
                {
                    RecurseObject(
                        sourceObject.SelectToken(elements[0]) as JObject,
                        propertyObject,
                        elements.Skip(1).ToArray()
                    );
                    sourceObject.Property(elements[0]);
                }
            }
            else
            {
                if (elements.Length == 1)
                {
                    sourceObject.Add(elements[0], propertyObject);
                }
                else
                {
                    JObject newObject = new JObject();
                    sourceObject.Add(elements[0], newObject);
                    RecurseObject(
                            newObject,
                            propertyObject,
                            elements.Skip(1).ToArray()
                        );
                }
            }
        }

        private static void ProcessJSONItem(JArray resultsArray, object item)
        {
            switch (item.GetType().ToString())
            {
                case "System.Object[,]":
                    ProcessJSONArray(resultsArray, (Object[,])item);
                    break;

                case "System.Object[]":
                case "System.Array":
                    ProcessJSONItem(resultsArray, item);
                    break;

                default:
                    if (typeof(string) == item.GetType())
                    {
                        try
                        {
                            resultsArray.Add(JToken.Parse(item as string));
                        }
                        catch
                        {
                            resultsArray.Add(item);
                        }
                    }
                    else
                    {
                        resultsArray.Add(item);
                    }
                    break;
            }
        }

        private static void ProcessJSONArray(JArray resultsArray, object[,] arrayObject)
        {
            for (int i = 0; i < arrayObject.GetLength(0); i++)
            {
                for (int j = 0; j < arrayObject.GetLength(1); j++)
                {
                    ProcessJSONItem(resultsArray, arrayObject[i, j]);
                }
            }
        }

        #endregion Helper Functions

        [ExcelFunction(
             Description = "Adds one ore more object to a JSON array",
             Category = "JSON",
             ExplicitRegistration = true
         )]
        public static object APPENDJSONARRAY(
             [ExcelArgument(
                   Name = "JSONArray",
                   Description = "The original JSON array, [] for an empty array."
             )]string jsonArray,

             [ExcelArgument(
                 Name = "values",
                 Description = "are values or selections to be joined"
             )]params object[] values
        )
        {
            JArray source;
            try
            {
                source = string.IsNullOrEmpty(jsonArray) ? new JArray() : JArray.Parse(jsonArray);
            }
            catch
            {
                return ExcelError.ExcelErrorRef;
            }

            try
            {
                foreach (var value in values)
                {
                    ProcessJSONItem(source, value);
                }
            }
            catch
            {
                return ExcelError.ExcelErrorNA;
            }

            return source.ToString(0);
        }

        [ExcelFunction(
            Description = "Joins a list of strings together, similar to concatenate, but with a seperator",
            Category = "JSON",
            ExplicitRegistration = true
        )]
        public static object ADDJSONPROPERTY(
            [ExcelArgument(
                   Name = "JSON Object",
                   Description = "The original JSON object, {} for an empty array."
             )]string source,

            [ExcelArgument(
                   Name = "PropertyPath",
                   Description = "dot (.) seperated property path."
             )]string property,

            [ExcelArgument(
                   Name = "Value",
                   Description = "value to add to property"
             )]object value,

            [ExcelArgument(
                Name = "[include_empty]",
                Description = "include empty items"
            )]bool IncludeEmpty = true
            )
        {
            JToken propertyObject;
            JObject sourceObject;
            string[] elements = property.Split('.');

            if (typeof(string) == value.GetType())
            {
                if (!IncludeEmpty & string.IsNullOrEmpty(value.ToString()))
                    return source;

                try
                {
                    propertyObject = JToken.Parse(value.ToString());
                }
                catch
                {
                    propertyObject = JToken.FromObject(value.ToString());
                }
            }
            else
            {
                if (!IncludeEmpty & value.GetType() == typeof(ExcelMissing))
                    return source;

                try
                {
                    propertyObject = JObject.Parse(value as string);
                }
                catch
                {
                    propertyObject = JToken.FromObject(value);
                }
            }

            try
            {
                sourceObject = JObject.Parse(source);
            }
            catch
            {
                return ExcelError.ExcelErrorRef;
            }

            RecurseObject(
                sourceObject as JObject,
                propertyObject,
                elements);
            return sourceObject.ToString(0);
        }
    }
}