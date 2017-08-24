using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.Json;
using CsvHelper;

namespace csvExporter
{
    internal class Program
    {
        public static readonly List<string> CsvHeaderRow = new List<string>
        {
            "Item Type",
            "Label",
            "Response",
            "Comment",
            "Media Hypertext Reference",
            "Location Coordinates",
            "Item Score",
            "Item Max Score",
            "Item Score Percentage",
            "Mandatory",
            "Failed Response",
            "Inactive",
            "Item ID",
            "Response ID",
            "Parent ID",
            "Audit Owner",
            "Audit Author",
            "Audit Name",
            "Audit Score",
            "Audit Max Score",
            "Audit Score Percentage",
            "Audit Duration (seconds)",
            "Date Started",
            "Time Started",
            "Date Completed",
            "Time Completed",
            "Audit ID",
            "Template ID",
            "Template Name",
            "Template Author"
        };

        // audit item empty response 
        private const string EmptyResponse = "";

        // audit item property constants
        private const string Label = "label";

        private const string Comments = "comments";
        private const string Type = "type";
        private const string Failed = "failed";
        private const string Score = "score";
        private const string MaxScore = "max_score";
        private const string ScorePercentage = "score_percentage";
        private const string CombinedScore = "combined_score";
        private const string CombinedMaxScore = "combined_max_score";
        private const string CombinedScorePercentage = "combined_score_percentage";
        private const string ParentId = "parent_id";
        private const string Response = "response";
        private const string Inactive = "inactive";
        private const string Id = "item_id";
        private const string Href = "href";
        private const string Signature = "signature";
        private const string Information = "information";
        private const string Media = "media";
        private const string Responses = "responses";

        // maps smartfield conditional statement IDs to the corresponding text
        private static readonly Dictionary<object, object> SmartfieldConditionalIdToStatementMap =
            new Dictionary<object, object>
            {
                // conditional statements for question field
                {"3f206180-e4f6-11e1-aff1-0800200c9a66", "if response selected"},
                {"3f206181-e4f6-11e1-aff1-0800200c9a66", "if response not selected"},
                {"3f206182-e4f6-11e1-aff1-0800200c9a66", "if response is"},
                {"3f206183-e4f6-11e1-aff1-0800200c9a66", "if response is not"},
                {"3f206184-e4f6-11e1-aff1-0800200c9a66", "if response is one of"},
                {"3f206185-e4f6-11e1-aff1-0800200c9a66", "if response is not one of"},
                // conditional statements for list field
                {"35f6c130-e500-11e1-aff1-0800200c9a66", "if response selected"},
                {"35f6c131-e500-11e1-aff1-0800200c9a66", "if response not selected"},
                {"35f6c132-e500-11e1-aff1-0800200c9a66", "if response is"},
                {"35f6c133-e500-11e1-aff1-0800200c9a66", "if response is not"},
                {"35f6c134-e500-11e1-aff1-0800200c9a66", "if response is one of"},
                {"35f6c135-e500-11e1-aff1-0800200c9a66", "if response is not one of"},
                // conditional statements for slider field
                {"cda7c330-e500-11e1-aff1-0800200c9a66", "if slider value is less than"},
                {"cda7c331-e500-11e1-aff1-0800200c9a66", "if slider value is less than or equal to"},
                {"cda7c332-e500-11e1-aff1-0800200c9a66", "if slider value is equal to"},
                {"cda7c333-e500-11e1-aff1-0800200c9a66", "if slider value is not equal to"},
                {"cda7c334-e500-11e1-aff1-0800200c9a66", "if slider value is greater than or equal to"},
                {"cda7c335-e500-11e1-aff1-0800200c9a66", "if the slider value is greater than"},
                {"cda7c336-e500-11e1-aff1-0800200c9a66", "if the slider value is between"},
                {"cda7c337-e500-11e1-aff1-0800200c9a66", "if the slider value is not between"},
                // conditional statements for checkbox field
                {"4e671f40-e4ff-11e1-aff1-0800200c9a66", "if the checkbox is checked"},
                {"4e671f41-e4ff-11e1-aff1-0800200c9a66", "if the checkbox is not checked"},
                // conditional statements for switch field
                {"3d346f00-e501-11e1-aff1-0800200c9a66", "if the switch is on"},
                {"3d346f01-e501-11e1-aff1-0800200c9a66", "if the switch is off"},
                // conditional statements for text field
                {"7c441470-e501-11e1-aff1-0800200c9a66", "if text is"},
                {"7c441471-e501-11e1-aff1-0800200c9a66", "if text is not"},
                // conditional statements for textsingle field
                {"6ff300f0-e501-11e1-aff1-0800200c9a66", "if text is"},
                {"6ff300f1-e501-11e1-aff1-0800200c9a66", "if text is not"},
                // conditional statements for signature field
                {"831f8ff0-e500-11e1-aff1-0800200c9a66", "if signature exists"},
                {"831f8ff1-e500-11e1-aff1-0800200c9a66", "if the signature does not exist"},
                {"831f8ff2-e500-11e1-aff1-0800200c9a66", "if the signature name is"},
                {"831f8ff3-e500-11e1-aff1-0800200c9a66", "if the signature name is not"},
                // conditional statements for barcode field
                {"8259d900-12e3-11e4-9191-0800200c9a66", "if the scanned barcode is"},
                {"8259d901-12e3-11e4-9191-0800200c9a66", "if the scanned barcode is not"}
            };

        // maps default answer IDs to corresponding Text
        private static readonly Dictionary<object, object> StandardResponseIdMap = new Dictionary<object, object>
        {
            {"8bcfbf00-e11b-11e1-9b23-0800200c9a66", "Yes"},
            {"8bcfbf01-e11b-11e1-9b23-0800200c9a66", "No"},
            {"8bcfbf02-e11b-11e1-9b23-0800200c9a66", "N/A"},
            {"b5c92350-e11b-11e1-9b23-0800200c9a66", "Safe"},
            {"b5c92351-e11b-11e1-9b23-0800200c9a66", "At Risk"},
            {"b5c92352-e11b-11e1-9b23-0800200c9a66", "N/A"}
        };

        public static object get_json_property(object obj, params object[] args)
        {
            foreach (var arg in args)
            {
                if (obj is List<object> && arg is int)
                {
                    var list = (List<object>) obj;
                    if (list.Count == 0)
                    {
                        return EmptyResponse;
                    }
                    obj = list[(int) arg];
                }
                else if (obj is Dictionary<string, object> &&
                         ((Dictionary<string, object>) obj).ContainsKey((string) arg))
                {
                    obj = ((Dictionary<string, object>) obj)[((string) arg)];
                }
                else
                {
                    return EmptyResponse;
                }
            }

            return obj ?? EmptyResponse;
        }

        class CsvExporter
        {
            private Dictionary<string, object> _auditJson;
            private List<object> _auditTable;

            private bool _exportInactiveItems;


            public CsvExporter(object auditJson, bool exportInactiveItems = true)
            {
                _auditJson = (Dictionary<string, object>) auditJson;
                _exportInactiveItems = exportInactiveItems;
                _auditTable = convert_audit_to_table();
            }

            public object audit_id()
            {
                return _auditJson["audit_id"];
            }

            public List<object> audit_items()
            {
                var objects = new List<object>();
                objects.AddRange((IEnumerable<object>) _auditJson["header_items"]);
                objects.AddRange((IEnumerable<object>) _auditJson["items"]);
                return objects;
            }

            public Dictionary<string, object> audit_custom_response_id_to_label_map()
            {
                var customResponseSets =
                    (Dictionary<string, object>) get_json_property(_auditJson,
                        new List<object> {"template_data", "response_sets"});
                var auditCustomResponseIdToLabelMap = new Dictionary<string, object>();
                foreach (var customResponseSet in customResponseSets)
                {
                    var dictionary = (Dictionary<string, object>) customResponseSet.Value;
                    var list = (List<object>) dictionary[Responses];
                    foreach (var response in list)
                    {
                        var dict = (Dictionary<string, object>) response;
                        auditCustomResponseIdToLabelMap[dict["id"].ToString()] = dict[Label];
                    }
                }
                return auditCustomResponseIdToLabelMap;
            }

            public List<object> common_audit_data()
            {
                var auditDataProperty = (Dictionary<string, object>) _auditJson["audit_data"];
                var templateDataProperty = (Dictionary<string, object>) _auditJson["template_data"];
                var auditDateCompleted = auditDataProperty["date_completed"];

                var auditDataAsList = new List<object>();
                auditDataAsList.Add(((Dictionary<string, object>) auditDataProperty["authorship"])["owner"]);
                auditDataAsList.Add(((Dictionary<string, object>) auditDataProperty["authorship"])["author"]);
                auditDataAsList.Add(auditDataProperty["name"]);
                auditDataAsList.Add(auditDataProperty[Score]);
                auditDataAsList.Add(auditDataProperty["total_score"]);
                auditDataAsList.Add(auditDataProperty[ScorePercentage]);
                auditDataAsList.Add(auditDataProperty["duration"]);
                auditDataAsList.Add(format_date(auditDataProperty["date_started"]));
                auditDataAsList.Add(format_time(auditDataProperty["date_started"]));
                auditDataAsList.Add(format_date(auditDateCompleted));
                auditDataAsList.Add(format_time(auditDateCompleted));
                auditDataAsList.Add(audit_id());
                auditDataAsList.Add(_auditJson["template_id"]);
                auditDataAsList.Add(((Dictionary<string, object>) templateDataProperty["metadata"])["name"]);
                auditDataAsList.Add(((Dictionary<string, object>) templateDataProperty["authorship"])["author"]);

                return auditDataAsList;
            }

            public static string format_date(object date)
            {
                if (date == null)
                {
                    return "";
                }

                var parsedDate = DateTime
                    .ParseExact((string) date, "yyyy-M-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture)
                    .ToUniversalTime();
                return parsedDate.ToString("dd MMMM yyyy", DateTimeFormatInfo.InvariantInfo);
            }


            public static string format_time(object date)
            {
                if (date == null)
                {
                    return "";
                }
                var parsedDate = DateTime
                    .ParseExact((string) date, "yyyy-M-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture)
                    .ToUniversalTime();
                return parsedDate.ToString("hh:mmtt", DateTimeFormatInfo.InvariantInfo);
            }

            public List<object> convert_audit_to_table()
            {
                _auditTable = new List<object>();
                foreach (Dictionary<string, object> item in audit_items())
                {
                    var rowArray = item_properties_as_list(item).Concat(common_audit_data()).ToList();
                    if (ToBool(get_json_property(item, Inactive)) && !ToBool(_exportInactiveItems))
                    {
                        continue;
                    }
                    _auditTable.Add(rowArray);
                }

                return _auditTable;
            }

            public void append_converted_audit_to_bulk_export_file(string outputCsvPath)
            {
                if (!File.Exists(outputCsvPath) && !_auditTable[0].Equals(CsvHeaderRow))
                {
                    _auditTable.Insert(0, CsvHeaderRow);
                }
                write_file(outputCsvPath, "ab");
            }

            public void save_converted_audit_to_file(string outputCsvPath, bool allowOverwrite)
            {
                var fileExists = !File.Exists(outputCsvPath);
                if (fileExists && !allowOverwrite)
                {
                    Console.WriteLine("File already exists at " + outputCsvPath);
                    Console.WriteLine(
                        "Please set allow_overwrite to True in config.yaml file. See ReadMe.md for further instruction");
                    Environment.Exit(1);
                }
                else if (fileExists)
                {
                    Console.WriteLine("Overwriting file at " + outputCsvPath);
                }
                
                if (!_auditTable[0].Equals(CsvHeaderRow))
                {
                    _auditTable.Insert(0, CsvHeaderRow);
                }
                write_file(outputCsvPath, "wb");
            }

            public void write_file(string outputCsvPath, string mode)
            {
                try
                {
                    var writer = mode.Contains("a")
                        ? File.AppendText(outputCsvPath)
                        : File.CreateText(outputCsvPath);
                    var wr = new CsvWriter(writer);
                    foreach (var row in _auditTable)
                    {
                        foreach (var value in (IEnumerable) row)
                        {
                            wr.WriteField((value ?? "").ToString());
                        }
                        wr.NextRecord();
                    }
                    wr.Dispose();
                    writer.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error saving audit_table to " + outputCsvPath);
                    Console.WriteLine(e);
                    Console.WriteLine(Environment.StackTrace);
                }
            }

            public object get_item_response(object item)
            {
                object response = EmptyResponse;
                var itemType = get_json_property(item, Type);
                if ("question".Equals(itemType))
                {
                    response = get_json_property(item, Responses, "selected", 0, Label);
                }
                else if ("list".Equals(itemType))
                {
                    foreach (var singleResponse in (IEnumerable) get_json_property(item, Responses, "selected"))
                    {
                        if (ToBool(singleResponse))
                        {
                            response += get_json_property(singleResponse, Label) + "\n";
                        }
                        var str = ((string) response);
                        response = str.Substring(0, str.Length - 1);
                    }
                }
                else if ("address".Equals(itemType))
                {
                    response = get_json_property(item, Responses, "location_text");
                }
                else if ("checkbox".Equals(itemType))
                {
                    response = ToBool(get_json_property(item, Responses, "value"));
                }
                else if ("switch".Equals(itemType))
                {
                    response = get_json_property(item, Responses, "value");
                }
                else if ("slider".Equals(itemType))
                {
                    response = get_json_property(item, Responses, "value");
                }
                else if ("drawing".Equals(itemType))
                {
                    response = get_json_property(item, Responses, "image", "media_id");
                }
                else if ("MEDIA".Equals(itemType))
                {
                    foreach (var image in (IEnumerable) get_json_property(item, Media))
                    {
                        response += "\n" + get_json_property(image, "media_id");
                    }
                    var str = ((string) response);
                    if (str.Length > 0)
                    {
                        response = str.Substring(1);
                    }
                }
                else if ("SIGNATURE".Equals(itemType))
                {
                    response = get_json_property(item, Responses, "name");
                }
                else if ("smartfield".Equals(itemType))
                {
                    response = get_json_property(item, "evaluation");
                }
                else if ("datetime".Equals(itemType))
                {
                    response = format_date(get_json_property(item, Responses, "datetime"));
                    response = response + " at " + format_time(get_json_property(item, Responses, "datetime"));
                }
                else if ("text".Equals(itemType) || "textsingle".Equals(itemType))
                {
                    response = get_json_property(item, Responses, "text");
                }
                else if (Information.Equals(itemType) && "link".Equals(get_json_property(item, "options", Type)))
                {
                    response = get_json_property(item, "options", "link");
                }
                else if (new HashSet<string>
                {
                    "dynamicfield",
                    "element",
                    "primeelement",
                    "asset",
                    "scanner",
                    "category",
                    "section",
                    Information
                }.Contains(itemType))
                {
                }
                else
                {
                    object id;
                    ((Dictionary<string, object>) item).TryGetValue(Id, out id);
                    Console.WriteLine("Unhandled item type: " + itemType + " from " + audit_id() + ", " + id);
                }
                return response;
            }

            public object get_item_response_id(object item)
            {
                object responseId = EmptyResponse;
                var itemType = get_json_property(item, Type);
                if ("question".Equals(itemType))
                {
                    responseId = get_json_property(item, Responses, "selected", 0, "id");
                }
                else if ("list".Equals(itemType))
                {
                    foreach (var singleResponse in (IEnumerable) get_json_property(item, Responses, "selected"))
                    {
                        if (ToBool(singleResponse))
                        {
                            responseId += get_json_property(singleResponse, "id") + "\n";
                        }
                        var str = (string) responseId;
                        responseId = str.Substring(0, str.Length - 1);
                    }
                }

                return responseId;
            }

            public object get_item_score(object item)
            {
                var score = get_json_property(item, "scoring", Score);
                if (score is int)
                {
                    return score;
                }
                var combinedScore = get_json_property(item, "scoring", CombinedScore);
                if (combinedScore is int)
                {
                    return combinedScore;
                }
                return EmptyResponse;
            }

            public object get_item_max_score(object item)
            {
                var maxScrore = get_json_property(item, "scoring", MaxScore);
                if (maxScrore is int)
                {
                    return maxScrore;
                }
                var combinedMaxScore = get_json_property(item, "scoring", CombinedMaxScore);
                if (combinedMaxScore is int)
                {
                    return combinedMaxScore;
                }
                return EmptyResponse;
            }

            public object get_item_score_percentage(object item)
            {
                var p = get_json_property(item, "scoring", ScorePercentage);
                if (p is int)
                {
                    return p;
                }
                var cp = get_json_property(item, "scoring", CombinedScorePercentage);
                if (cp is int)
                {
                    return cp;
                }
                return EmptyResponse;
            }

            public object get_item_label(object item)
            {
                object label = EmptyResponse;
                var itemType = get_json_property(item, Type);
                if ("smartfield".Equals(itemType))
                {
                    var customResponseIdToLabelMap = audit_custom_response_id_to_label_map();
                    var conditionalId = (string) get_json_property(item, "options", "condition");

                    if (ToBool(conditionalId))
                    {
                        label = DeepClone(
                            SmartfieldConditionalIdToStatementMap.TryGetValue(conditionalId, out label));
                    }
                    foreach (var value in (IEnumerable) get_json_property(item, "options", "values"))
                    {
                        label += "|";
                        if (StandardResponseIdMap.ContainsKey(value))
                        {
                            label += StandardResponseIdMap[value].ToString();
                        }
                        else if (customResponseIdToLabelMap.ContainsKey(value.ToString()))
                        {
                            label += customResponseIdToLabelMap[value.ToString()].ToString();
                        }
                        else
                        {
                            label += value.ToString();
                        }
                        label += "|";
                    }

                    return label;
                }

                return get_json_property(item, Label);
            }

            public object get_item_type(object item)
            {
                var itemType = get_json_property(item, Type);
                if (Information.Equals(itemType))
                {
                    itemType += " - " + get_json_property(item, "options", Type);
                }
                return itemType;
            }

            public object get_item_media(object item)
            {
                var itemType = get_json_property(item, Type);
                if (Information.Equals(itemType) && Media.Equals(get_json_property(item, "options", Type)))
                {
                    return get_json_property(item, "options", Media, Href);
                }
                if ("drawing".Equals(itemType) || Signature.Equals(itemType))
                {
                    return get_json_property(item, Responses, "image", Href);
                }

                var media = get_json_property(item, Media);
                if (media is string)
                {
                    return media;
                }

                return string.Join("\n", ((IEnumerable<object>) media)
                    .Select(x => ((Dictionary<string, object>) x)[Href]).ToList());
            }

            public object get_item_location_coordinates(object item)
            {
                var itemType = get_json_property(item, Type);
                if ("address".Equals(itemType))
                {
                    var locationCoordinates =
                        get_json_property(item, "responses", "location", "geometry", "coordinates");
                    if (locationCoordinates is List<object>)
                    {
                        return String.Join(", ", ((IEnumerable) locationCoordinates));
                    }
                }
                return EmptyResponse;
            }

            public static T DeepClone<T>(T obj)
            {
                using (var ms = new MemoryStream())
                {
                    var formatter = new BinaryFormatter();
                    formatter.Serialize(ms, obj);
                    ms.Position = 0;

                    return (T) formatter.Deserialize(ms);
                }
            }

            public bool ToBool(object value)
            {
                if (value == null)
                {
                    return false;
                }
                if (value is bool)
                {
                    return (bool) value;
                }
                if (value is IEnumerable<object>)
                {
                    return ((IEnumerable<object>) value).FirstOrDefault() != null;
                }
                if (value is int)
                {
                    return ((int) value) != 0;
                }
                if (value is string)
                {
                    return ((string) value).Length > 0;
                }

                return true;
            }

            private List<object> item_properties_as_list(Dictionary<string, object> item)
            {
                object a = EmptyResponse;
                if (!item.ContainsKey(Type) || "text" != item[Type] as string && "textsingle" != item[Type] as string)
                {
                    a = get_json_property(item, Responses, "text");
                }

                return new List<object>
                {
                    get_item_type(item),
                    get_item_label(item),
                    get_item_response(item),
                    a,
                    get_item_media(item),
                    get_item_location_coordinates(item),
                    get_item_score(item),
                    get_item_max_score(item),
                    get_item_score_percentage(item),
                    get_json_property(item, "options", "is_mandatory"),
                    get_json_property(item, Responses, Failed),
                    get_json_property(item, Inactive),
                    get_json_property(item, Id),
                    get_item_response_id(item),
                    get_json_property(item, ParentId)
                };
            }
        }

        public static int Main(string[] args)
        {
            try
            {
                DoMain(args);
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine(Environment.StackTrace);
                return 1;
            }
        }

        private static void DoMain(string[] args)
        {
            if (CsvExporter.format_date("2017-03-03T03:45:58.090Z") != "03 March 2017")
            {
                throw new Exception("test 1 failed");
            }

            if (CsvExporter.format_time("2017-03-03T03:45:58.090Z") != "03:45AM")
            {
                throw new Exception("test 2 failed");
            }

            for (var i = 0; i < args.Length; i++)
            {
                var arg = args[i];
                var readAllText = File.ReadAllText(arg);
                var auditJson = new JsonParser().Parse(readAllText);
                var csvExporter = new CsvExporter(auditJson);

                csvExporter.save_converted_audit_to_file(Path.GetFileName(arg) + ".csv", allowOverwrite: true);
            }
        }
    }
}