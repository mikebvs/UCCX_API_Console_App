using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.Text;

namespace UCCX_API_Testing
{
    class Reader
    {
        string filePath { get; set; }
        public Reader(string file)
        {
            filePath = file;
        }
        public List<SkillData> ReadSkillData(int sheetIndex)
        {
            FileInfo file = new FileInfo(filePath);
            List<SkillData> skillData = new List<SkillData>();
            using (ExcelPackage package = new ExcelPackage(file))
            {

                StringBuilder sb = new StringBuilder();
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                //Console.WriteLine("ROWS: " + rowCount.ToString() + "\nCOLUMNS: " + colCount.ToString());

                for (int i = 2; i <= rowCount; i++)
                {
                    string name = String.Empty;
                    string add = String.Empty;
                    string remove = String.Empty;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (j == 1 && worksheet.Cells[i, j].Value != null)
                        {
                            name = worksheet.Cells[i, j].Value.ToString();
                        }
                        else if (j == 2 && worksheet.Cells[i, j].Value != null)
                        {
                            add = worksheet.Cells[i, j].Value.ToString();
                        }
                        else if (j == 3 && worksheet.Cells[i, j].Value != null)
                        {
                            remove = worksheet.Cells[i, j].Value.ToString();
                        }
                        else if (add == String.Empty || add == null)
                        {
                            continue;
                        }
                    }
                    if (add != String.Empty && add != null && add != "" && add.Length > 2)
                    {
                        //Console.WriteLine("Adding " + name);
                        SkillData skill = new SkillData(name, add, remove);
                        skillData.Add(skill);
                    }
                }
            }
            return skillData;
        }
        public List<AgentData> ReadAgentData(int sheetIndex)
        {

            FileInfo file = new FileInfo(filePath);

            List<AgentData> agentData = new List<AgentData>();
            using (ExcelPackage package = new ExcelPackage(file))
            {
                StringBuilder sb = new StringBuilder();
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int i = 2; i <= rowCount; i++)
                {
                    string sheetName = String.Empty;
                    string sheetQueue = String.Empty;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (j == 1 && worksheet.Cells[i, j].Value.ToString() != null)
                        {
                            sheetName = worksheet.Cells[i, j].Value.ToString();
                        }
                        else if (j == 2 && worksheet.Cells[i, j].Value.ToString() != null)
                        {
                            sheetQueue = worksheet.Cells[i, j].Value.ToString();
                        }
                        else if (sheetQueue == String.Empty || sheetQueue == null)
                        {
                            continue;
                        }
                    }
                    if (sheetQueue != String.Empty && sheetQueue != null)
                    {
                        AgentData agent = new AgentData(sheetName, sheetQueue);
                        agentData.Add(agent);
                    }
                }
            }
            return agentData;
        }
    }
}
