using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;

namespace TestLinkUtil
{
    class Program
    {
        private static String workSpace = "";
        private static String action = "";

        private static WebClient client = new WebClient();
        private static String url = "";
        private static String devKey = "";
        private static String testProjectName = "";
        private static String testPlanName = "";

        private static String fileOut = "";
        private static String customFieldName = "";
        private static String[] customFieldNames = { };

        static void Main(string[] args)
        {
            String seekPath = "";
            //url = "http://127.0.0.1:81/testlink/lib/api/xmlrpc/v1/xmlrpc.php"; 
            //devKey = "10cbc3572531d71ee30ed19e23e89850";

            //String testPlanName = "AZ TestPlan";
            //String testProjectName = "AutoZ_Teste";
            //String fileOut = "TestCases";
            //String testProjectName = args[0];
            //String testPlanName = args[1];
            //String fileOut = args[2];

            //WORKSPACE set by Jenkins
            workSpace = Environment.GetEnvironmentVariable("WORKSPACE");

            int i = 0;
            while (i < args.Length) {
                switch (args[i])
                {
                    case "-o": //out file path/name to be used in actions that have a file output
                        fileOut = args[i + 1];
                        i += 2;
                        break;
                    case "-u": //specify TestLink URL
                               //default readed from hudson.plugins.testlink.TestLinkBuilder.xml in the agent folder root 
                        url = args[i + 1];
                        i += 2;
                        break;
                    case "-k": //specify TestLink DevKey
                               //default readed from hudson.plugins.testlink.TestLinkBuilder.xml in the agent folder root 
                        devKey = args[i + 1];
                        i += 2;
                        break;
                    case "-pj": //specify TestLink Project Name
                                //default set by TestLink Jenkins plugin on TESTLINK_TESTPROJECT_NAME
                        testProjectName = args[i + 1];
                        i += 2;
                        break;
                    case "-pl": //specify TestLink Test Plan Name 
                                //default set by TestLink Jenkins plugin on TESTLINK_TESTPLAN_NAME
                        testPlanName = args[i + 1];
                        i += 2;
                        break;
                    case "-a": //specify the actions to be taken
                        //GetTestCasesMSTest: Get TestCase names from TestPlan and set to output file in the format: "TESTCASES= /test:<TCName>"
                        //      Use the names from the CustomField passed by parameter or the TCName if CF parameter = ""
                        //TapResults: read all .tap files from seekPath (-s) and report the results to testlink, also uploading attachments
                        //SaveCustomFields: save all the specified custom fields (-cf) values from all the TC in the TestPlan to the output file. One line for each TC, values separated by ","
                        action = args[i + 1];
                        i += 2;
                        break;
                    case "-s": //specify the path where to seek results or other files
                        seekPath = args[i + 1];
                        i += 2;
                        break;
                    case "-cf": //specify TestLink Custom Field Names separated by ","
                                //  the correspondence with each TC is always made using the first CFName specified
                        string cf = args[i + 1];
                        customFieldNames = cf.Split(',');
                        customFieldName = customFieldNames[0];
                        i += 2;
                        break;
                    default:
                        i++;
                        break;
                }
            }

            if (url.Equals("") || devKey.Equals(""))
            {
                if (!getTestLinkInstall(workSpace + "\\..\\..\\hudson.plugins.testlink.TestLinkBuilder.xml"))
                {
                    if (!getTestLinkInstall(workSpace + "\\..\\..\\..\\hudson.plugins.testlink.TestLinkBuilder.xml"))
                    {
                        return;
                    }
                }
            }

            if (testProjectName.Equals(""))
            {
                testProjectName = Environment.GetEnvironmentVariable("TESTLINK_TESTPROJECT_NAME");
            }

            if (testPlanName.Equals(""))
            {
                testPlanName = Environment.GetEnvironmentVariable("TESTLINK_TESTPLAN_NAME");
            }

            switch (action)
            {
                case "GetTestCasesMSTest":
                    getTestCasesMSTest(customFieldName);
                    break;
                case "TapResults":
                    setTapResults(tapResults(seekPath), customFieldName);
                    break;
                case "SaveCustomFields":
                    saveCustomFields(customFieldNames);
                    break;
                default:
                    break;
            }


        }

        private static void saveCustomFields(String[] customFields)
        {
            String testPlanId = getTestPlanByName(testProjectName, testPlanName);

            List<Dictionary<string, string>> testCases = getTestCasesForTestPlan(testPlanId);

            foreach (string cf in customFields)
            {
                testCases = addCustomField(testCases, cf);
            }

            FileStream fs = File.Open(fileOut, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            
            foreach (Dictionary<string, string> tc in testCases)
            {
                string line = tc["name"];
                foreach (string cf in customFields)
                {
                    line = line + "," + tc[cf];
                }
                sw.WriteLine(line);
            }
            
            sw.Close();
        }

        private static void setTapResults(List<Dictionary<string, object>> testCasesResults, string cfName)
        {
            Dictionary<string, string> build = getLatestBuildForTestPlan(getTestPlanByName(testProjectName, testPlanName));

            String testPlanId = getTestPlanByName(testProjectName, testPlanName);

            List<Dictionary<string, string>> testCases = getTestCasesForTestPlan(testPlanId);

            testCases = addCustomField(testCases, cfName);

            foreach (Dictionary<string, string> tc in testCases)
            {
                Dictionary<string, object> tcR = testCasesResults.Find(t => t["name"].ToString() == tc[cfName]);

                if (tcR!=null)
                {
                    String execId = reportTCResult(tc["id"], tcR, testPlanId, build);
                    uploadExecutionAttachment(execId, "text/plain", tc[cfName]+".tap", "TAP File", tcR["filePath"].ToString());
                    List<Dictionary<string, string>> files = (List < Dictionary < string, string>>) tcR["attachments"];
                    foreach (Dictionary<string, string> file in files)
                    {
                        uploadExecutionAttachment(execId, file["Type"], file["Name"], file["Description"], file["Location"]);
                    }
                }
            }
        }

        private static void uploadExecutionAttachment(string execId, string type, string name, string description, string path)
        {
            byte[] fileContent = File.ReadAllBytes(path);
            String base64 = Convert.ToBase64String(fileContent);

            client.Headers.Add("Content-Type", "text/xml");

            String content = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?><methodCall xmlns:ex=\"http://ws.apache.org/xmlrpc/namespaces/extensions\"><methodName>tl.uploadExecutionAttachment</methodName><params><param><value><struct><member><name>devKey</name><value>" +
                devKey +
                "</value></member><member><name>executionid</name><value>" +
                execId +
                "</value></member><member><name>filetype</name><value>" +
                type +
                "</value></member><member><name>filename</name><value>" +
                name +
                "</value></member><member><name>description</name><value>" +
                description +
                "</value></member><member><name>title</name><value>" +
                name +
                "</value></member><member><name>content</name><value>" +
                base64 +
                "</value></member><member><name>platformid</name><value><ex:nil/></value></member><member><name>customfields</name><value><ex:nil/></value></member><member><name>platformname</name><value><ex:nil/></value></member>" +
                "<member><name>fktable</name><value>executions</value></member></struct></value></param></params></methodCall>";
            byte[] responseString = client.UploadData(url, Encoding.UTF8.GetBytes(content));
        }

        private static string reportTCResult(string id, Dictionary<string, object> tcR, string testPlanId, Dictionary<string, string> build)
        {
            String content = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?><methodCall xmlns:ex=\"http://ws.apache.org/xmlrpc/namespaces/extensions\"><methodName>tl.reportTCResult</methodName><params><param><value><struct><member><name>devKey</name><value>" +
                devKey +
                "</value></member><member><name>testplanid</name><value>" +
                testPlanId +
                "</value></member><member><name>testcaseid</name><value>" +
                id +
                "</value></member><member><name>buildid</name><value>" +
                build["id"] +
                "</value></member><member><name>buildname</name><value>" +
                build["name"] +
                "</value></member><member><name>notes</name><value>" +
                //tcR["notes"] +
                "</value></member><member><name>status</name><value>" +
                tcR["outcome"] +
                "</value></member><member><name>platformid</name><value><ex:nil/></value></member><member><name>customfields</name><value><ex:nil/></value></member><member><name>platformname</name><value><ex:nil/></value></member>" +
                "<member><name>bugid</name><value><ex:nil/></value></member><member><name>guess</name><value><ex:nil/></value></member><member><name>testcaseexternalid</name><value><ex:nil/></value></member><member><name>overwrite</name><value><ex:nil/></value></member></struct></value></param></params></methodCall>";
            String responseString = client.UploadString(url, content);

            Regex rx = new Regex("<name>id<\\/name><value><int>([^<]+)<\\/int>");
            Match m = rx.Match(responseString);

            return m.Groups[1].Value;
        }

        private static List<Dictionary<string, string>> addCustomField(List<Dictionary<string, string>> testCases, string cfName)
        {
            String projectId = getTestProjectByName(testProjectName);

            foreach (Dictionary<string, string> tc in testCases)
            {
                String cfValue = getTestCaseCustomFieldDesignValue(tc["id"], cfName, projectId);
                tc.Add(cfName, cfValue);
            }

            return testCases;
        }

        private static string getTestCaseCustomFieldDesignValue(string tcId, string cfName, string projectId)
        {
            String content = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?><methodCall xmlns:ex=\"http://ws.apache.org/xmlrpc/namespaces/extensions\"><methodName>tl.getTestCaseCustomFieldDesignValue</methodName><params><param><value><struct><member><name>devKey</name><value>" +
                devKey +
                "</value></member><member><name>testprojectid</name><value>" +
                projectId +
                "</value></member><member><name>testcaseid</name><value>" +
                tcId +
                "</value></member><member><name>customfieldname</name><value>" +
                cfName +
                "</value></member><member><name>details</name><value>full</value></member><member><name>testcaseexternalid</name><value><ex:nil/></value></member><member><name>version</name><value><i4>1</i4></value></member></struct></value></param></params></methodCall>";
            String responseString = client.UploadString(url, content);

            Regex rx = new Regex("<name>value<\\/name><value><string>([^<]+)<\\/string>");
            Match m = rx.Match(responseString);

            return m.Groups[1].Value;
        }

        private static string getTestProjectByName(string testProjectName)
        {
            String content = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?><methodCall xmlns:ex=\"http://ws.apache.org/xmlrpc/namespaces/extensions\"><methodName>tl.getTestProjectByName</methodName><params><param><value><struct><member><name>devKey</name><value>" +
                devKey +
                "</value></member><member><name>testprojectname</name><value>" +
                testProjectName +
                "</value></member></struct></value></param></params></methodCall>";
            String responseString = client.UploadString(url, content);

            Regex rx = new Regex("<name>id<\\/name><value><string>(\\d+)<\\/string>");
            Match m = rx.Match(responseString);

            return m.Groups[1].Value;
        }

        private static List<Dictionary<string, object>> tapResults(String path)
        {
            if (workSpace != null && !Directory.Exists(path))
            {
                path = workSpace + "\\" + path;
                if (!Directory.Exists(path))
                {
                    return null;
                }
            }

            IEnumerable<String> files = Directory.EnumerateFiles(path, "*.tap");

            List<Dictionary<string, object>> testCases = new List<Dictionary<string, object>>();

            foreach (String file in files)
            {
                Dictionary<string, object> tc = new Dictionary<string, object>();

                tc.Add("filePath", file);
                tc.Add("name", Path.GetFileNameWithoutExtension(file));

                FileStream fs = File.Open(file, FileMode.Open);
                StreamReader sr = new StreamReader(fs);

                StringBuilder notes = new StringBuilder();

                String outcome = "b";
                List<Dictionary<string, string>> attachments = new List<Dictionary<string, string>>();

                while (!sr.EndOfStream)
                {
                    String line = sr.ReadLine();

                    if (line.ToUpper().StartsWith("OK") && outcome.Equals("b"))
                    {
                        outcome = "p";
                    }
                    else if (line.ToUpper().StartsWith("NOT OK"))
                    {
                        outcome = "f";
                    }
                    else if (line.StartsWith("#"))
                    {
                        notes.AppendLine(line.Substring(1));
                    }
                    else if (line.TrimStart(' ').Equals("---")) {
                        bool extFiles = false;
                        Dictionary<string, string> attachment = new Dictionary<string, string>();
                        while (!line.Trim(' ').Equals("..."))
                        {
                            line = sr.ReadLine();
                            if (line.Trim(' ').Equals("Files:"))
                            {
                                extFiles = true;
                            }
                            else if (line.TrimEnd(' ').EndsWith(":") && extFiles)
                            {
                                if (attachment.Count > 0)
                                {
                                    attachments.Add(attachment);
                                    attachment = new Dictionary<string, string>();
                                }
                            }
                            else if (line.TrimStart(' ').ToUpper().StartsWith("FILE-LOCATION"))
                            {
                                attachment.Add("Location", line.Substring(line.IndexOf(":") + 1).Trim(' '));
                            }
                            else if (line.TrimStart(' ').ToUpper().StartsWith("FILE-NAME"))
                            {
                                attachment.Add("Name", line.Substring(line.IndexOf(":") + 1).Trim(' '));
                            }
                            else if (line.TrimStart(' ').ToUpper().StartsWith("FILE-SIZE"))
                            {
                                attachment.Add("Size", line.Substring(line.IndexOf(":") + 1).Trim(' '));
                            }
                            else if (line.TrimStart(' ').ToUpper().StartsWith("FILE-TYPE"))
                            {
                                attachment.Add("Type", line.Substring(line.IndexOf(":") + 1).Trim(' '));
                            }
                            else if (line.TrimStart(' ').ToUpper().StartsWith("FILE-DESCRIPTION"))
                            {
                                attachment.Add("Description", line.Substring(line.IndexOf(":") + 1).Trim(' '));
                            }
                        }
                        if (attachment.Count > 0)
                        {
                            attachments.Add(attachment);
                        }
                    }
                }

                tc.Add("notes", notes);
                tc.Add("outcome", outcome);
                tc.Add("attachments", attachments);

                testCases.Add(tc);
                sr.Close();
            }

            return testCases;
        }

        private static Dictionary<string, string> getLatestBuildForTestPlan(String testPlanId)
        {
            String content = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?><methodCall xmlns:ex=\"http://ws.apache.org/xmlrpc/namespaces/extensions\"><methodName>tl.getLatestBuildForTestPlan</methodName><params><param><value><struct><member><name>devKey</name><value>" +
                devKey +
                "</value></member><member><name>testplanid</name><value><i4>" +
                testPlanId +
                "</i4></value></member></struct></value></param></params></methodCall>";
            String responseString = client.UploadString(url, content);

            Dictionary<string, string> results = new Dictionary<string, string>();

            Regex rx = new Regex("<name>id<\\/name><value><string>([^<]+)<\\/string>");
            results.Add("id", rx.Match(responseString).Groups[1].Value);
            rx = new Regex("<name>name<\\/name><value><string>([^<]+)<\\/string>");
            results.Add("name", rx.Match(responseString).Groups[1].Value);

            return results;
        }

        private static void getTestCasesMSTest(String cfName)
        {
            String testPlanId = getTestPlanByName(testProjectName, testPlanName);

            List<Dictionary<string, string>> testCases = getTestCasesForTestPlan(testPlanId);

            String referenceField = "name";

            if (!cfName.Equals(""))
            {
                testCases = addCustomField(testCases, cfName);
                referenceField = cfName;
            }

            String tcString = "TESTCASES=";
            foreach (Dictionary<string, string> tc in testCases)
            {
                    tcString += " /test:" + tc[referenceField];
            }

            FileStream fs = File.Open(fileOut, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(tcString);
            sw.Close();
        }

        private static bool getTestLinkInstall(string file)
        {
            if (!File.Exists(file))
            {
                return false;
            }

            FileStream fs = File.Open(file, FileMode.Open);
            StreamReader sr = new StreamReader(fs);
            String fileContent = sr.ReadToEnd();
            sr.Close();

            try
            {
                Regex rx = new Regex("<url>([^<]+)<\\/url>");
                url = rx.Match(fileContent).Groups[1].Value;
                rx = new Regex("<devKey>([^<]+)<\\/devKey>");
                devKey = rx.Match(fileContent).Groups[1].Value;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static List<Dictionary<string, string>> getTestCasesForTestPlan(string testPlanId)
        {
            String content = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?><methodCall xmlns:ex=\"http://ws.apache.org/xmlrpc/namespaces/extensions\"><methodName>tl.getTestCasesForTestPlan</methodName><params><param><value><struct><member><name>testcaseid</name><value><ex:nil/></value></member><member><name>devKey</name><value>" +
                devKey +
                "</value></member><member><name>keywordid</name><value><ex:nil/></value></member><member><name>keywords</name><value><ex:nil/></value></member><member><name>getstepsinfo</name><value><boolean>1</boolean></value></member><member><name>executiontype</name><value>2</value></member><member><name>testplanid</name><value><i4>" +
                testPlanId +
                "</i4></value></member><member><name>buildid</name><value><ex:nil/></value></member><member><name>executed</name><value><ex:nil/></value></member><member><name>details</name><value>full</value></member><member><name>executestatus</name><value><ex:nil/></value></member><member><name>assignedto</name><value><ex:nil/></value></member></struct></value></param></params></methodCall>";
            String responseString = client.UploadString(url, content);

            Regex rx = new Regex("tcase_name<\\/name><value><string>([^<]+)<\\/string>[\\s\\S]+?tcase_id<\\/name><value><string>([^<]+)<\\/string>");
            MatchCollection mc = rx.Matches(responseString);

            List<Dictionary<string, string>> tc = new List<Dictionary<string, string>>();

            foreach (Match m in mc)
            {
                Dictionary<string, string> t = new Dictionary<string, string>();
                t.Add("name", m.Groups[1].Value);
                t.Add("id", m.Groups[2].Value);
                tc.Add(t);
            }

            return tc;
        }

        private static string getTestPlanByName(string testProjectName, string testPlanName)
        {
            String content = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?><methodCall xmlns:ex=\"http://ws.apache.org/xmlrpc/namespaces/extensions\"><methodName>tl.getTestPlanByName</methodName><params><param><value><struct><member><name>devKey</name><value>" +
                devKey + 
                "</value></member><member><name>testplanname</name><value>" +
                testPlanName + 
                "</value></member><member><name>testprojectname</name><value>" +
                testProjectName +
                "</value></member></struct></value></param></params></methodCall>";
            String responseString = client.UploadString(url, content);

            Regex rx = new Regex("<name>id<\\/name><value><string>(\\d+)<\\/string>");
            Match m = rx.Match(responseString);

            return m.Groups[1].Value;
        }
    }
}
