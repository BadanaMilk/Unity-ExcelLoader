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
}
