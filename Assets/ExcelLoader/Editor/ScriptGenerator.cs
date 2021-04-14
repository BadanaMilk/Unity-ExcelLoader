
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using UnityEditor;
using System.Linq;

namespace ExcelLoader
{
    public class ScriptGenerator
    {
        ExcelLoader_Setting setting;

        public string dataFileName;
        public string dataTableName;
             
        string dataTemplate;
        string nameSpaceText;
        const string memberFieldTemplate = "[UnityEngine.SerializeField] private $MemberType$ $MemberName$;\n\tpublic $MemberType$ Data_$MemberName$ { get { return $MemberName$; } }\n";
        const string dataParamsTemplate = "$MemberType$ _$MemberName$";
        const string memberInitTemplate = "$MemberName$ = _$MemberName$;";

        List<HeaderData> listHeader;

        public ScriptGenerator()
        {
            dataTemplate = File.ReadAllText(ExcelLoader_Editor.excelLoaderPath + "/Template/TableDataTemplate.txt", Encoding.UTF8);
        }

        public void SetScriptGenerator(string _sheetName, string _nameSpaceText, ExcelLoader_Setting _setting, List<HeaderData> _listHeader)
        {
            dataFileName = GetDataName(_sheetName);
            dataTableName = _sheetName;
            setting = _setting;
            listHeader = _listHeader;
            nameSpaceText = _nameSpaceText;
        }

        public void DataScriptGenerate()
        {
            string _csText = string.Copy(dataTemplate);

            _csText = _csText.Replace("$Namespace$", nameSpaceText);
            _csText = _csText.Replace("$TableData$", dataFileName);
            _csText = _csText.Replace("$TableKey$", listHeader[0].name);
            _csText = _csText.Replace("$TableName$", dataTableName);

            string[] _stringMember = new string[listHeader.Count];
            string[] _strDataParams = new string[listHeader.Count];
            string[] _strMemberInits = new string[listHeader.Count];
            for (int _index = 0; _index < listHeader.Count; _index++)
            {
                _stringMember[_index] = memberFieldTemplate.Replace("$MemberType$", listHeader[_index].TypeName()).Replace("$MemberName$", listHeader[_index].name);
                _strDataParams[_index] = dataParamsTemplate.Replace("$MemberType$", listHeader[_index].TypeName()).Replace("$MemberName$", listHeader[_index].name);
                _strMemberInits[_index] = memberInitTemplate.Replace("$MemberType$", listHeader[_index].TypeName()).Replace("$MemberName$", listHeader[_index].name);
            }
            
            _csText = _csText.Replace("$MembersField$", string.Join("\n\t", _stringMember));
            _csText = _csText.Replace("$DataParam$", string.Join(",\n\t\t", _strDataParams));
            _csText = _csText.Replace("$MembersInit$", string.Join("\n\t\t", _strMemberInits));

            string _fullPath = string.Format("{0}/{1}.cs", setting.GetClassFullPath(), dataFileName);
            string _folderPath = Path.GetDirectoryName(_fullPath);
            if (!Directory.Exists(_folderPath))
            {
                EditorUtility.DisplayDialog(
                    "경고",
                    "스크립트 파일을 넣는 폴더가 존재하지 않습니다.\n" + "경로 : " + _folderPath,
                    "OK"
                    );
                return;
            }

            using (var writer = new StreamWriter(_fullPath))
            {
                writer.Write(_csText);
                writer.Close();
            }
        }

        public Type GetDataType()
        {
            return GetType(dataFileName);
        }

        public Type GetTableType()
        {
            return GetType("ExcelLoader.DataContainer");
        }

        public static Type GetType(string typeName)
        {
            var type = Type.GetType(typeName);
            if (type != null) return type;
            foreach (var a in AppDomain.CurrentDomain.GetAssemblies())
            {
                type = a.GetType(typeName);
                if (type != null)
                    return type;
            }
            return null;
        }

        public static string GetDataName(string _sheetName)
        {
            return _sheetName + "Data";
        }
    }
}