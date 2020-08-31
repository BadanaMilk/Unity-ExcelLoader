using System.Collections;
using System.Collections.Generic;
using UnityEditor;
using UnityEngine;

public class ExcelLoader_Setting : ScriptableObject
{
    [SerializeField]
    public string excelPath;
    [SerializeField]
    public string classPath;
    [SerializeField]
    public string dataPath;
    [SerializeField]
    public string csvPath;
    [SerializeField]
    public string defaultPath;

    public void SetDefaultPath()
    {
        int _index = Application.dataPath.IndexOf("Assets");
        defaultPath = Application.dataPath.Remove(_index, Application.dataPath.Length - _index);
    }

    public string GetExcelFullPath() { return defaultPath + excelPath; }
    public string GetClassFullPath() { return defaultPath + classPath; }
    public string GetDataFullPath() { return defaultPath + dataPath; }
    public string GetCsvFullPath() { return defaultPath + csvPath; }
}
