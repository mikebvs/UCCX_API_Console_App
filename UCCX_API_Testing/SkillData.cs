using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCCX_API_Testing
{
    class SkillData
    {
        public string Name { get; set; }
        public Dictionary<string, int> SkillsAdded { get; set; }
        public List<string> SkillsRemoved { get; set; }
        public SkillData(string name, string toAdd, string toRemove)
        {
            //Initialize Name
            Name = name;

            //Create Add Dictionary
            List<string> addList = new List<string>();
            addList.AddRange(toAdd.Split(';'));
            Dictionary<string, int> addDictionary = new Dictionary<string, int>();
            foreach (string str in addList)
            {
                //Console.WriteLine(str);
                int firstParenth = str.IndexOf("(") + 1;
                int lastParenth = str.LastIndexOf(")");
                int difference = lastParenth - firstParenth;
                string valConvert = str.Substring(firstParenth, difference);
                string key = str.Substring(0, firstParenth - 1);
                int val = Convert.ToInt32(valConvert);
                addDictionary.Add(key, val);
            }
            //Initialize Add
            SkillsAdded = addDictionary;

            //Create Remove List
            List<string> removeList = new List<string>();
            removeList.AddRange(toRemove.Split(';'));
            //Initialize Remove
            SkillsRemoved = removeList;
        }
    }
}