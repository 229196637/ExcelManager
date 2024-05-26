using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using ExcelManager;
using OfficeOpenXml;
using UnityEditor;
using UnityEngine;
using Object = UnityEngine.Object;

namespace DataSystem
{
    public class ExcelBuilder
    {
        //创建两个对应文件夹
        public void CreateDirectory()
        {
            Directory.CreateDirectory(FilePath.ExcelPath);
            Directory.CreateDirectory(FilePath.ExcelPath);
        }
        //创建脚本
        public void CreateScript()
        {
            Directory.CreateDirectory(FilePath.ExcelPath);
            
            string[] files = Directory.GetFiles(FilePath.ExcelPath);

            foreach (var file in files)
            {
                
                if(Path.GetExtension(file) != ".xlsx" || Path.GetFileName(file).StartsWith("~$"))
                    continue;
                
                //创建脚本
                ExcelFile excelFile = new ExcelFile(file);
                SetValues(excelFile);
            }
            
        }
        
        /// <summary>
        /// 根据脚本文件去创建
        /// 这里保存表单对应一个脚本
        /// </summary>
        public void CreateExcel(string excelName)
        {
            Directory.CreateDirectory(FilePath.ExcelPath);
            
            if(excelName == "")
            {
                Debug.LogError("请输入你想要创建的Excel名字");
                return;
            }
            
            Object[] selects = Selection.objects;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                foreach (var s in selects)
                {
                    //判断这个是否为cs脚本
                    Type t = GetTypeByName(FilePath.Namespace+s.name);
                    if (t == null)
                    {
                        Debug.LogError("目标文件不是脚本文件");
                        continue;
                    }
                
                    //判断这个是否继承 DataItem这个父类
                    Type parentClass = typeof(DataItem);
                    if(!t.IsSubclassOf(parentClass))
                    {
                        Debug.LogError("目标脚本文件父类并非指定父类DataItem");
                        continue;
                    }
                
                    //添加一个工作表，我们规定，工作表为这个数据类的名字，一表格对应一种数据类型 s.name为选中脚本的名字
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(s.name);
                    FieldInfo[] fieldInfos = t.GetFields();
                    
                    //处理表单
                    int index = 1;
                    foreach (var field in fieldInfos)
                    {
                        string fieldType = field.FieldType.ToString();
                        //处理后的string System.int32留下int32
                        //变量类型
                        string s1 = fieldType.Split(".")[1];
                        
                        //处理表单
                        //第一行统一为备注
                        worksheet.Cells[1,index].Value = "备注,自己写";
                        /*worksheet.Cells[1,index].Value = */
                        worksheet.Cells[2,index].Value = field.Name;
                        worksheet.Cells[3,index].Value = s1;
                        
                        
                        index += 1;
                    }
                    
                }
                
                if (excelPackage.Workbook.Worksheets.Count == 0) return;
                //保存数据
                string save_path = FilePath.ExcelPath+ excelName+ ".xlsx";
                FileInfo fileInfo = new FileInfo(save_path);
                excelPackage.SaveAs(fileInfo);
                //保存刷新
                AssetDatabase.SaveAssets();
                AssetDatabase.Refresh();
                Debug.Log("创建成功");
            }
        }
        
        /// <summary>
        /// 读取Scriptobject 如果没有则创建
        /// </summary>
        /// <param name="assetPath"></param>
        /// <param name="assetType"></param>
        /// <returns></returns>
        public Object LoadOrCreateScriptObject(string assetPath,Type assetType)
        {
            Directory.CreateDirectory(FilePath.ScriptObejctPath);
            
            var asset = AssetDatabase.LoadAssetAtPath(assetPath,assetType);
            
            if (asset == null)
            {
                asset = ScriptableObject.CreateInstance(assetType);
                AssetDatabase.CreateAsset((ScriptableObject)asset, assetPath);
                asset.hideFlags = HideFlags.NotEditable;
            }
            
            return asset;
        }

        /// <summary>
        /// 设置ScriptObject里面的值
        /// </summary>
        /// <param name="excelFile"></param>
        public void SetValues(ExcelFile excelFile)
        {
            FileInfo fileInfo = new FileInfo(excelFile.Excel_File);
            // 获取当前程序集

            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                var worksheets = excelPackage.Workbook.Worksheets;

                foreach (var worksheet in worksheets)
                {
                    if(worksheet.Dimension == null) continue;
                    
                    Type scirptObjectType = GetTypeByName(FilePath.Namespace+excelFile.ScriptFile.FileName+"ScriptObject");
                    if(scirptObjectType == null) continue;
                    
                    Object script = LoadOrCreateScriptObject("Assets/DataBase/DataObject/"+ excelFile.ScriptFile.FileName +".asset",scirptObjectType);
                    
                    FieldInfo field = script.GetType().GetField(worksheet+"List");
                    Type lineType = GetTypeByName(FilePath.Namespace+worksheet);
                    Type listType = typeof(List<>).MakeGenericType(lineType);
                    MethodInfo listAdd = listType.GetMethod("Add",new Type[]{lineType});
                    object list = Activator.CreateInstance(listType);
                
                    for (int i = worksheet.Dimension.Start.Row +3; i <= worksheet.Dimension.End.Row; i++)
                    {
                        object listObject = Activator.CreateInstance(lineType);
                        for (int j = worksheet.Dimension.Start.Column;j<= worksheet.Dimension.End.Row;j++)
                        {
                            if(worksheet.GetValue(2,j) == null ||worksheet.GetValue(3,j)==null ) continue;
                            
                            FieldInfo variable = lineType.GetField(worksheet.GetValue(2,j).ToString());
                            string tableValue = worksheet.GetValue(i,j).ToString();
                            variable.SetValue(listObject,Convert.ChangeType(tableValue,GetTypeByExcel(worksheet.GetValue(3,j).ToString())));
                        }
                        
                        listAdd.Invoke(list,new object[]{listObject});
                    }
                        
                    field.SetValue(script,list);
                    
                    
                    EditorUtility.SetDirty(script);
                }
            }
        }
        
        Type GetTypeByName(string className)
        {
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();

            foreach (Assembly assembly in assemblies)
            {
                Type type = assembly.GetType(className);
                if (type != null)
                {
                    return type;
                }
            }

            return null;
        }
        

        Type GetTypeByExcel(string excelType)
        {
            switch (excelType.ToLower())
            {
                case "string":
                    return typeof(string);
                    break;
                case "int":
                    return typeof(int);
                    break;
                case "int32":
                    return typeof(int);
                    break;
                case "float":
                    return typeof(float);
                    break;
                case "short":
                    return typeof(short);
                    break;
                case "bool":
                    return typeof(bool);
                    break;
                case "long":
                    return typeof(long);
                    break;
                
            }
            return null;
        }
        
    }

}