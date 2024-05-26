using System;
using System.IO;
using System.Reflection;
using OfficeOpenXml;
using UnityEditor;
using UnityEngine;

namespace DataSystem
{
    public class ExcelFile
    {
        public ScriptFile ScriptFile {get;private set;}
        public string Excel_File {get;private set;}
        public ExcelFile(string excelFile)
        {
            this.Excel_File = excelFile;
            
            ScriptFile = new ScriptFile(excelFile);
            //创建脚本
            ScriptFile.CreateScript(this); 
            ScriptFile.CreateLineDataScript(this);
            //保存刷新
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }
        
    }

    public class ScriptFile
    {
        public string FileName;
        public ScriptFile(string file)
        {
            this.FileName = Path.GetFileNameWithoutExtension(Path.GetFileName(file));
            Directory.CreateDirectory(FilePath.ScriptFilePath);
        }
        
        public void CreateScript(ExcelFile excelFile)
        {
            string filePath = FilePath.ScriptFilePath + FileName+"ScriptObject" + ".cs";
            
            FileInfo fileInfo = new FileInfo(excelFile.Excel_File);
            
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                var worksheets = excelPackage.Workbook.Worksheets;

                using (StreamWriter writer = new StreamWriter(filePath,false))
                {
                    // 写入文件头部
                    writer.WriteLine("using UnityEngine;");
                    writer.WriteLine("using System.Collections.Generic;");
                    writer.WriteLine("using ExcelManager;");
                    writer.WriteLine("namespace DataSystem");
                    writer.WriteLine("{");
                        
                    writer.WriteLine($"public class {FileName}ScriptObject : DataScriptObject");
                    writer.WriteLine("{");
                    writer.WriteLine("");
                    foreach (var worksheet in worksheets)
                    {
                        if(worksheet.Dimension == null) continue;
                        writer.WriteLine($"public List<{worksheet}> {worksheet}List;");
                    }
                    writer.WriteLine("}");
                
                    writer.WriteLine("}");
                    
                }
            }
            
        }

        public void CreateLineDataScript(ExcelFile excelFile)
        {
            FileInfo fileInfo = new FileInfo(excelFile.Excel_File);
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                var worksheets = excelPackage.Workbook.Worksheets;
                
                foreach (var worksheet in worksheets)
                {
                    string filePath = FilePath.ScriptFilePath + worksheet + ".cs";
                    if(worksheet.Dimension == null) continue;
                    
                    using (StreamWriter writer = new StreamWriter(filePath,false))
                    {
                        // 写入文件头部
                        writer.WriteLine("using System;");
                        writer.WriteLine("using UnityEngine;");
                        writer.WriteLine("using System.Collections.Generic;");
                        writer.WriteLine("using ExcelManager;");
                        writer.WriteLine("namespace DataSystem");
                        writer.WriteLine("{");
                
                        writer.WriteLine("[Serializable]");
                        writer.WriteLine($"public class {worksheet} : DataItem");
                        writer.WriteLine("{");
                        
                        for (int i = worksheet.Dimension.Start.Column; i <= worksheet.Dimension.End.Column; i++)
                        {  
                            if(worksheet.GetValue(2,i) == null ||worksheet.GetValue(3,i)==null ) continue;
                            
                            string fieldName = worksheet.GetValue(2,i).ToString();
                            string fieldtype = worksheet.GetValue(3,i).ToString();
                            
                            
                            writer.WriteLine("public " + fieldtype +" "+ fieldName +";");
                        }

                        writer.WriteLine("");
                        writer.WriteLine("}");
                        writer.WriteLine("}");
                        
                    }
                }
            }
            
        }
    }
    
}
