using System;
using System.Collections.Generic;
using System.Net;
using System.Xml;
using System.Linq;
using System.IO;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;

namespace UCCX_API_Testing
{
    struct Skill
    {
        public string SkillName;
        public string refURL;
        public string Id;
    };
    struct SkillsToAdd
    {
        public string Name;
        public string CompetencyLevel;
        public string refURL;
    };
    struct UserToEdit
    {
        public string Firstname;
        public string Lastname;
        public string UserID;
        public string Extension;
        public string refURL;
        public string newQueue;
    };
    class Program
    {
        public static int usersProcessed = 0;
        static void Main(string[] args)
        {
            //Required for dotnet run --project <PATH> command to be used to execute the process via batch file
            string workingDirectory = Environment.CurrentDirectory;
            Console.WriteLine("\n\nWORKING DIRECTORY: " + workingDirectory);
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            //Console.WriteLine(projectDirectory);
            string jsonPath = workingDirectory + "\\appsettings.json";

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(jsonPath, optional: true, reloadOnChange: true);
            
            IConfigurationRoot configuration = builder.Build();
            
            //Pull Config items
            string username = configuration.GetSection("UCCXCredentials")["Username"];
            string password = configuration.GetSection("UCCXCredentials")["Password"];
            string env = SetEnv();
            string rootURL = SetRootURL(env, configuration);
            string filePath = SetFilePath(env, configuration);

            string logPath = configuration.GetSection("Logging")["Path"] + env;
            string logName = "WFM_UpdateQueue_" + System.DateTime.Now.ToString("MMddyyyy_hh-mm-ss") + ".txt";
            logName = logPath + "\\" + logName;
            using (StreamWriter w = File.AppendText(logName))
            {
                Console.WriteLine("\n>Log File Created: " + logName + "\n");
                Log("Beginning Run... [ENVIRONMENT: " + env + "] [URL: " + rootURL + "] [USERNAME: " + username + "] [EXCEL FILE: " + filePath + "]", w);
                //Create Config Object

                //SkillsToAdd Test Information, this will come from an Excel File on NAS Drive in the future
                List<SkillsToAdd> nsl = new List<SkillsToAdd>();

                //---------------------------------------------------------------------------------------------------------
                //TESTING INFORMATION -------------------------------------------------------------------------------------
                //---------------------------------------------------------------------------------------------------------


                ////Users to be updated List, this will come from an Excel File on NAS Drive in the future
                //List<String> users = new List<String>();
                //
                ////User 1
                //users.Add("Shannon Baskette");
                //
                ////User 2
                //users.Add("Sarah Young");
                //
                ////User 3
                //users.Add("Andrea Cunningham");
                //
                ////User 4
                //users.Add("Sonny Kidd");
                //
                //
                //SkillsToAdd ns = new SkillsToAdd();
                //
                ////Skill 1
                //ns.Name = "CS_Tier1";
                //ns.CompetencyLevel = "5";
                //nsl.Add(ns);
                //
                ////Skill 2
                //ns.Name = "CS_Tier2";
                //ns.CompetencyLevel = "5";
                //nsl.Add(ns);
                //
                ////Skill 3
                //ns.Name = "CS_Tier3";
                //ns.CompetencyLevel = "5";
                //nsl.Add(ns);
                //
                ////Skill 4
                //ns.Name = "CS_Priority";
                //ns.CompetencyLevel = "8";
                //nsl.Add(ns);
                //
                ////Skill 5
                //ns.Name = "CS_Renewals";
                //ns.CompetencyLevel = "8";
                //nsl.Add(ns);
                ////---------------------------------------------------------------------------------------------------------
                ////---------------------------------------------------------------------------------------------------------
                ////---------------------------------------------------------------------------------------------------------



                //Initialize Excel Reader Object with file path
                Reader xlReader = new Reader(filePath);

                //Build List of agents that need their queue changed/what queue to move them to
                List<AgentData> agentData = new List<AgentData>();
                Console.WriteLine(">Reading Agent Data from Excel File...\n");
                agentData = xlReader.ReadAgentData(1);

                //Build list of skills required for each queue
                List<SkillData> skillData = new List<SkillData>();
                Console.WriteLine(">Reading Skill Data from Excel File...\n");
                skillData = xlReader.ReadSkillData(2);

                //Retrieves all Skill Information from the ../adminapi/skill endpoint to prevent multiple redundant calls
                List<Skill> skillsList = new List<Skill>();
                Console.WriteLine(">Retrieving All Skill via API Endpoint <URL>/adminapi/skill...\n");
                skillsList = GetAllSkills(username, password, rootURL + "/skill");
                
                //Cross References the SkillsToAdd List with the Skills List pulled from API, adds necessary information to update user skillMaps
                //DEBUG -- Print Skills to Add
                //foreach (SkillsToAdd skn in nsl)
                //{
                //    Console.WriteLine("{0} -- {1}\n{2}\n\n", skn.Name, skn.CompetencyLevel, skn.refURL);
                //}
                //----------------------------

                var client = new WebClient { Credentials = new NetworkCredential(username, password) };
                var response = client.DownloadString(rootURL + "/resource");


                XmlDocument xml = new XmlDocument();
                xml.LoadXml(response);
                Console.WriteLine(">Users Processed: " + usersProcessed.ToString() + "/" + agentData.Count.ToString());
                foreach (XmlNode xn in xml.DocumentElement)
                {
                    if(agentData.Any(agent => agent.agentName == xn["firstName"].InnerText + " " + xn["lastName"].InnerText))
                    {
                        usersProcessed++;
                        //Define User Variable to pass as parameter when updating user skillMap in UpdateUserSkillMap()
                        UserToEdit user = BuildUser(xn["firstName"].InnerText, xn["lastName"].InnerText, xn["userID"].InnerText, xn["self"].InnerText, xn["extension"].InnerText, agentData.Where(agent => agent.agentName == xn["firstName"].InnerText + " " + xn["lastName"].InnerText).Select(agent => agent.Queue).First());
                        
                        nsl = UpdateSkillsToAdd(skillsList, skillData, user.newQueue, w);
                        
                        //Console.WriteLine("\nSKILLS BEING ADDED: \n");
                        //foreach(SkillsToAdd skta in nsl)
                        //{
                        //    Console.WriteLine("SKILL: {0} -- {1}\n{2}\n\n", skta.Name, skta.CompetencyLevel, skta.refURL);
                        //}

                        //Returns XML String info of skillMap node
                        string skillMap = ReturnSkillMap(xn);

                        //Creates new skillMap XML String to insert into user skillMap
                        skillMap = AppendMultipleSkills(skillMap, nsl);
                        //Console.WriteLine("SKILL MAP FOR: " + user.newQueue + "\n\n" + skillMap);
                        
                        //Log which user is being edited
                        BeginLog(user, w);
                        Log("Attempting to update the XMLNode-->skillMap for " + user.Firstname + " " + user.Lastname, w);
                        try
                        {
                            //Attempt to update User skillMap
                            UpdateUserSkillMap(user, skillMap, xn.OuterXml, rootURL, username, password);
                            Log(">" + user.Firstname + " " + user.Lastname + "'s skills have been successfully updated for the following Queue: " + user.newQueue, w);
                        }
                        catch (Exception e)
                        {
                            //Log reason for error if occurred
                            Log(">An Error occurred updating the XMLNode-->skillMap. Please refer to the caught exception: " + e.Message.ToString(), w);
                        }
                        EndLog(w);
                        Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop - 1);
                        Console.WriteLine(">Users Processed: " + usersProcessed.ToString() + "/" + agentData.Count.ToString());
                    }
                }

            //using (StreamReader r = File.OpenText(logName))
            //{
            //    DumpLog(r);
            //}
            }
        }
        public static string SetEnv()
        {
            string env = "DEV";
            switch (Environment.MachineName.ToUpper())
            {
                case "VAL-H7T4SQ2":
                    env = "DEV";
                    break;
                case "VAL-61LJXT2":
                    env = "DEV";
                    break;
                case "VAVPC-ROBO-02":
                    env = "STAGE";
                    break;
                case "VAVPC-ROBO-05":
                    env = "STAGE";
                    break;
                case "VAVPC-ROBO-07":
                    env = "STAGE";
                    break;
                case "VAVPC-ROBO-01":
                    env = "PROD";
                    break;
                case "VAVPC-ROBO-03":
                    env = "PROD";
                    break;
                case "VAVPC-ROBO-04":
                    env = "PROD";
                    break;
                case "VAVPC-ROBO-06":
                    env = "PROD";
                    break;
            }
            Console.WriteLine("Current Environment: {0}", env.ToString());
            return env;
        }
        public static string SetRootURL(string env, IConfigurationRoot configuration)
        {
            string rootURL = "";
            if (env == "PROD")
            {
                rootURL = configuration.GetSection("UCCX_URL")["PROD"];
                Console.WriteLine("Current Root URL: " + rootURL);
            }
            else
            {
                rootURL = configuration.GetSection("UCCX_URL")["DEV"];
                Console.WriteLine("Current Root URL: " + rootURL);
            }
            return rootURL;
        }
        public static string SetFilePath(string env, IConfigurationRoot configuration)
        {
            string filePath = "";
            if(env == "PROD")
            {
                filePath = configuration.GetSection("ExcelFile")["PROD"];
            }
            else
            {
                filePath = configuration.GetSection("ExcelFile")["DEV"];
            }
            Console.WriteLine("Current Excel File: " + filePath);
            return filePath;
        }
        public static List<Skill> GetAllSkills(string username, string password, string endpoint)
        {
            List<Skill> skills = new List<Skill>();
            var client = new WebClient { Credentials = new NetworkCredential(username, password) };
            var response = client.DownloadString(endpoint);

            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(response);
            XmlNodeList xnl = xdoc.GetElementsByTagName("skill");
            foreach(XmlNode xn in xnl)
            {
                Skill newSkill = new Skill();
                newSkill.SkillName = xn["skillName"].InnerText;
                newSkill.Id = xn["skillId"].InnerText;
                newSkill.refURL = xn["self"].InnerText;
                skills.Add(newSkill);
            }
            return skills;
        }
        public static List<SkillsToAdd> UpdateSkillsToAdd(List<Skill> skta, List<SkillData> skData, string queue, TextWriter w)
        {
            List<SkillsToAdd> nslU = new List<SkillsToAdd>();
            Dictionary<string, int> queueDict = new Dictionary<string, int>();
            queueDict = skData.Where(p => p.Name == queue).First().SkillsAdded;
            foreach (KeyValuePair<string, int> kvp in queueDict)
            {
                Skill skillInfo = new Skill();
                if(skta.Any(p => p.SkillName == kvp.Key))
                {
                    skillInfo = skta.Where(p => p.SkillName == kvp.Key).First();
                    SkillsToAdd newSkill = new SkillsToAdd();
                    newSkill.Name = kvp.Key;
                    newSkill.CompetencyLevel = kvp.Value.ToString();
                    newSkill.refURL = skillInfo.refURL;
                    nslU.Add(newSkill);
                }
                else
                {
                    Log(kvp.Key + " not found.", w);
                }
            }
            return nslU;
        }
        public static string ReturnSkillMap(XmlNode xNode)
        {
            XmlDocument xmlInner = new XmlDocument();
            xmlInner.LoadXml(xNode.OuterXml);
            XmlNodeList xnl = xmlInner.GetElementsByTagName("skillMap");
            if(xnl.Count > 0)
            {
                return xnl[0].OuterXml;
            }
            else
            {
                return "<skillMap />";
            }
        }
        public static string AppendSkill(string xmlMain, int competencyLevel, string skillRefURL, string skillName)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(xmlMain);

            string fragXML = "<skillCompetency><competenceLevel>COMPETENCY_LEVEL</competenceLevel><skillNameUriPair name=\"SKILL_NAME\"><refURL>REF_URL</refURL></skillNameUriPair></skillCompetency>";
            fragXML = fragXML.Replace("COMPETENCY_LEVEL", competencyLevel.ToString()).Replace("SKILL_NAME", skillName).Replace("REF_URL", skillRefURL);

            XmlDocumentFragment xfrag = xdoc.CreateDocumentFragment();
            xfrag.InnerXml = fragXML;

            xdoc.DocumentElement.AppendChild(xfrag);
            return xdoc.OuterXml;
        }

        public static string AppendMultipleSkills(string xmlMain, List<SkillsToAdd> skills)
        {
            XmlDocument xdoc = new XmlDocument();

            xdoc.LoadXml(xmlMain);
            xdoc.InnerXml = "<skillMap />"; //Reset skillMap to nothing before appending new skills
            foreach(SkillsToAdd nskill in skills)
            {
                string fragXML = "<skillCompetency><competencelevel>COMPETENCY_LEVEL</competencelevel><skillNameUriPair name=\"SKILL_NAME\"><refURL>REF_URL</refURL></skillNameUriPair></skillCompetency>";
                fragXML = fragXML.Replace("COMPETENCY_LEVEL", nskill.CompetencyLevel.ToString()).Replace("SKILL_NAME", nskill.Name).Replace("REF_URL", nskill.refURL);

                XmlDocumentFragment xfrag = xdoc.CreateDocumentFragment();
                xfrag.InnerXml = fragXML;

                xdoc.DocumentElement.AppendChild(xfrag);
            }
            return xdoc.InnerXml;
        }
        public static string UpdateUserSkillMap(UserToEdit user, string skillMap, string userOuterXml, string rootURL, string username, string password)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(userOuterXml);

            XmlNode node = xdoc.SelectSingleNode("/resource/skillMap");
            //Console.WriteLine("CURRENT SKILL MAP\n" + node.OuterXml + "\n\n");

            XmlDocument xskilldoc = new XmlDocument();
            xskilldoc.LoadXml(skillMap);

            XmlNode xNew = xskilldoc.SelectSingleNode("/skillMap");
            //Console.WriteLine("NEW SKILL MAP\n" + xNew.InnerXml + "\n\n");

            node.InnerXml = xNew.InnerXml;
            //Console.WriteLine("MODIFIED SKILL MAP\n" + node.OuterXml);
            //Console.WriteLine(xdoc.OuterXml);
            string requestXML = xdoc.OuterXml;
            string destinationURL = rootURL + "/resource/" + user.UserID;
            HttpWebResponse response = postXMLData(destinationURL, requestXML, username, password);
            if(response.StatusCode != HttpStatusCode.OK)
            {
                throw (new System.Exception(response.StatusCode + ": " + response.StatusDescription));
            }

            return response.StatusCode.ToString();
        }
        public static HttpWebResponse postXMLData(string destinationURL, string requestXML, string username, string password)
        {            
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(destinationURL);

            //Add Basic Authorization Headers
            String encoded = System.Convert.ToBase64String(System.Text.Encoding.GetEncoding("ISO-8859-1").GetBytes(username + ":" + password));
            request.Headers.Add("Authorization", "Basic " + encoded);

            //Add Standard Encoding (Do not add into Request Body Header)
            byte[] bytes;
            bytes = System.Text.Encoding.ASCII.GetBytes(requestXML);
            request.ContentType = "text/xml; encoding='utf-8'";
            request.ContentLength = bytes.Length;

            //Method is PUT, not POST
            request.Method = "PUT";
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();
            //Get Response and action accordingly
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                //Report nothing if successful
                Stream responseStream = response.GetResponseStream();
                string responseStr = new StreamReader(responseStream).ReadToEnd();
                Console.WriteLine(response.StatusCode.ToString() + "\n" + responseStr);
            }
            return response;
        }
        public static UserToEdit BuildUser(string fname, string lname, string uid, string refurl, string ext, string queue)
        {
            UserToEdit user = new UserToEdit();
            user.Firstname = fname;
            user.Lastname = lname;
            user.UserID = uid;
            user.refURL = refurl;
            user.Extension = ext;
            user.newQueue = queue;
            return user;
        }
        public static void BeginLog(UserToEdit user, TextWriter w)
        {
            w.Write("\r\nLog Entry : ");
            w.WriteLine($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
            w.WriteLine($"  :{user.Firstname} {user.Lastname} -- {user.UserID} [REFURL: {user.refURL}]");
            w.WriteLine($"  :");

        }
        public static void Log(string logMessage, TextWriter w)
        {
            w.WriteLine($"  :{logMessage}");
        }
        public static void EndLog(TextWriter w)
        {
            w.WriteLine("---------------------------------------------------------------------------------------------");
        }
        public static void DumpLog(StreamReader r)
        {
            string line;
            while ((line = r.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
        }
    }
}
