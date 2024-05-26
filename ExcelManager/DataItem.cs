using System.Collections.Generic;
using UnityEditor;
using UnityEngine;

namespace ExcelManager
{
    public class DataItem
    {
    
    }
    
    public class DataScriptObject : ScriptableObject
    {
        readonly static string ScriptFilePath ="Assets/DataBase/DataObject/";
        
        public static T GetData<T>(string dataName) where T : DataScriptObject
        {
            string path = ScriptFilePath + dataName + ".asset";
            T data_instance = AssetDatabase.LoadAssetAtPath<T>(path);
            return data_instance;
        }
        
        
    }
}
