using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCCX_API_Testing
{
    class AgentData
    {
        public string agentName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Queue { get; set; }
        public AgentData(string sheetName, string sheetQueue)
        {
            agentName = sheetName;
            //Console.WriteLine(sheetName);
            FirstName = sheetName.Substring(0, sheetName.IndexOf(" "));
            LastName = sheetName.Substring(sheetName.IndexOf(FirstName) + 1);
            Queue = sheetQueue;
        }
    }
}
