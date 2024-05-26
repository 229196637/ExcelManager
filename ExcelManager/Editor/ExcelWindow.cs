using System;
using UnityEditor;
using UnityEngine;

namespace DataSystem
{
    public class ExcelWindow : EditorWindow
    {
        private string excel_name = "";
        public static ExcelBuilder excelBuilder = new ExcelBuilder();
        [MenuItem("ExcelBuilder/BuilderWindow")]
        static void Window()
        {
            
            ExcelWindow excelWindow = GetWindow<ExcelWindow>();
            excelWindow.Show();
        }
        
        void OnGUI()
        {   
            GUILayout.Label("Excel生成器", GUILayout.Width(100), GUILayout.Height(30)); //提示语句
            GUILayout.Space(2);
            if (GUILayout.Button("CreateDirectory"))
            {
                excelBuilder.CreateDirectory();
            }
            GUILayout.Space(10);
            GUILayout.Label("根据Excel表格创建脚本文件:"+FilePath.ExcelPath);
            if (GUILayout.Button("CreateScript"))
            {
                excelBuilder.CreateScript();
            }
            GUILayout.Space(5);
            GUILayout.Label("重新读取Excel表格里面数据");
            if (GUILayout.Button("ReloadScriptObject"))
            {
                excelBuilder.CreateScript();
            }
            GUILayout.Space(5);
            GUILayout.Label("根据Excel表格创建脚本文件:"+FilePath.ScriptFilePath);
            excel_name = EditorGUILayout.TextField("请输入你想创建的Excel名字:",excel_name);
            GUILayout.Space(2);
            if (GUILayout.Button("CreateExcel"))
            {
                excelBuilder.CreateExcel(excel_name);
            }
            
        }
    }    
}
