﻿using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using UnityEngine;
using UnityEditor;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor.IMGUI.Controls;
using System.Text;
using System.Diagnostics;

namespace ExcelLoader
{
    public enum CellType
    {
        None,
        String,
        Byte,
        Short,
        Int,
        Long,
        Float,
        Double,
        Enum,
        Bool,
    }

    public enum eExcelLoaderType : byte
    {
        MultiSelect,
        SingleSelect
    }

    [InitializeOnLoadAttribute]
    public class ExcelLoader_Editor : EditorWindow
    {
        #region GUI View Data

        bool isCompiling = false;

        float treeViewStartPosY = 145f;
        float treeViewEndPosY = 265f;
        float treeViewPadding = 5f;
        float excelListViewWidth = 250f;

        Rect singleSelectListViewRect { get { return new Rect(treeViewPadding, treeViewStartPosY, excelListViewWidth, treeViewEndPosY); } }
        Rect multiSelectListViewRect { get { return new Rect(treeViewPadding, treeViewStartPosY, position.width - (treeViewPadding * 2), treeViewEndPosY); } }
        Rect sheetListViewRect { get { return new Rect(treeViewPadding + excelListViewWidth, treeViewStartPosY, position.width - (excelListViewWidth + (treeViewPadding * 2)), treeViewEndPosY); } }
     

        Vector2 scrollPosition;

        ExcelFileTreeView singleListView;       //단일 선택 엑셀 트리뷰 GUI
        ExcelFileTreeView multiListView;        //다중 선택 엑셀 트리뷰 GUI
        ExcelSheetTreeView sheetListView;       //시트 트리뷰 GUI
        TreeViewState singleListViewState;      //단일 선택 엑셀 트리뷰 상태 정보
        TreeViewState multiListViewState;       //다중 선택 엑셀 트리뷰 상태 정보
        TreeViewState sheetListViewState;       //시트 트리뷰 상태정보
        SearchField searchField;                //엑셀 검색 필드

        eExcelLoaderType currentLoadType;

        #endregion

        static ExcelLoader_Editor instance;

        public static string excelLoaderPath;
        ExcelLoader_Setting settingData;                        //엑셀 로드 폴더 세팅 정보
        List<string> listSearchedFiles = new List<string>();    //엑셀 파일 경로에서 찾은 엑셀 파일 리스트

        ScriptGenerator scriptGenerator;                        //스크립트 생성 클래스

        #region 단일 선택용 변수들
        IWorkbook selectWorkbook = null;                        //선택한 엑셀 파일
        ISheet selectSheet = null;                              //선택한 엑셀 시트
        List<HeaderData> listSelectSheetHeaders = new List<HeaderData>();//선택한 시트의 필드 정보
        List<string> listExcelSheets = new List<string>();      //선택한 엑셀파일의 시트 리스트

        int excelSelectID;                                      //선택한 엑셀의 트리뷰 Index ID
        int sheetSelectID;                                      //선택한 시트의 트리뷰 Index ID
        string lastSelectExcelName;                             //마지막으로 선택한 엑셀의 이름
        string lastSelectSheetName;                             //마지막으로 선택한 시트의 이름
        #endregion

        #region 다중 선택용 변수들
        List<int> listMultiSelects = new List<int>();           //선택한 엑셀 파일의 TreeView ID리스트
        #endregion


        [MenuItem("Tools/ExcelLoader", false)]
        public static void Open()
        {
            if (instance != null)
                DestroyImmediate(instance);
            instance = EditorWindow.GetWindow<ExcelLoader_Editor>();
            instance.Show();
        }

        [UnityEditor.Callbacks.DidReloadScripts]
        public static void OnRecompile()
        {
            ExcelLoader_Editor[] _editors = (ExcelLoader_Editor[])Resources.FindObjectsOfTypeAll(typeof(ExcelLoader_Editor));
            if (_editors.Length != 0)
            {
                instance = _editors[0];
                instance.RefreshEditor();
            }
        }

        private void Awake()
        {
            instance = this;
            string _filePath = new System.Diagnostics.StackTrace(true).GetFrame(0).GetFileName();
            excelLoaderPath = _filePath.Remove(0, _filePath.IndexOf("Assets")).Replace("\\", "/").Replace("ExcelLoader_Editor.cs", "");
            Init();
        }

        private void OnDestroy()
        {
            instance = null;
        }

        /// <summary>
        /// 엑셀 로더의 초기화 함수
        /// </summary>
        private void Init()
        {
            scriptGenerator = new ScriptGenerator();
            //엑셀 로더 세팅 정보를 저장할 폴더 경로
            string _settingPath = excelLoaderPath + "Setting";
            //세팅 폴더가 존재하는지 확인하고 없다면 폴더를 만들어준다.
            if (AssetDatabase.IsValidFolder(_settingPath) == false)
            {
                AssetDatabase.CreateFolder(excelLoaderPath.Remove(excelLoaderPath.LastIndexOf('/'), 1), "Setting");
            }
            //세팅 정보를 로드한다. 없다면 생성
            settingData = AssetDatabase.LoadAssetAtPath<ExcelLoader_Setting>(string.Format("{0}/ExcelLoaderSetting.asset", _settingPath));
            if (settingData == null)
            {
                settingData = ScriptableObject.CreateInstance<ExcelLoader_Setting>();
                AssetDatabase.CreateAsset(settingData, string.Format("{0}/ExcelLoaderSetting.asset", _settingPath));
            }
            settingData.SetDefaultPath();
            EditorUtility.SetDirty(settingData);
            AssetDatabase.SaveAssets();

            //GUI를 위해 트리뷰를 생성
            singleListViewState = new TreeViewState();
            singleListView = new ExcelFileTreeView(singleListViewState, settingData, ref listSearchedFiles, OnClickSingleSelectExcelList);
            multiListViewState = new TreeViewState();
            multiListView = new ExcelFileTreeView(multiListViewState, settingData, ref listSearchedFiles, null);
            sheetListViewState = new TreeViewState();
            sheetListView = new ExcelSheetTreeView(sheetListViewState, settingData, ref listExcelSheets, OnClickSheetList);
            searchField = new SearchField();
        }

        /// <summary>
        /// 에디터 갱신
        /// </summary>
        private void RefreshEditor()
        {
            string _filePath = new System.Diagnostics.StackTrace(true).GetFrame(0).GetFileName();
            excelLoaderPath = _filePath.Remove(0, _filePath.IndexOf("Assets")).Replace("\\", "/").Replace("ExcelLoader_Editor.cs", "");

            scriptGenerator = new ScriptGenerator();

            multiListViewState = new TreeViewState();
            multiListView = new ExcelFileTreeView(multiListViewState, settingData, ref listSearchedFiles, null);
            multiListViewState.selectedIDs = listMultiSelects;
            
            singleListViewState = new TreeViewState();
            singleListView = new ExcelFileTreeView(singleListViewState, settingData, ref listSearchedFiles, OnClickSingleSelectExcelList);
            singleListView.SetSelection(new List<int>() { excelSelectID });

            sheetListViewState = new TreeViewState();
            sheetListView = new ExcelSheetTreeView(sheetListViewState, settingData, ref listExcelSheets, OnClickSheetList);

            searchField = new SearchField();
            if (currentLoadType == eExcelLoaderType.SingleSelect)
            {
                sheetListView.SetSelection(new List<int>() { sheetSelectID });
                OnClickSingleSelectExcelList(excelSelectID, lastSelectExcelName);
                OnClickSheetList(sheetSelectID, lastSelectSheetName);
            }
        }

        /// <summary>
        /// GUI 갱신
        /// </summary>
        private void RefreshGUI()
        {
            sheetListView.Reload();
            singleListView.Reload();
        }

        private void OnGUI()
        {
            isCompiling = EditorApplication.isCompiling;
            GUI.enabled = !isCompiling;
            scrollPosition = EditorGUILayout.BeginScrollView(scrollPosition);
            EditorGUILayout.BeginVertical();

            DrawSelectFolder("엑셀 테이블 경로", ref settingData.excelPath);
            DrawSelectFolder("테이블 클래스 경로", ref settingData.classPath);
            DrawSelectFolder("데이터 저장 경로", ref settingData.dataPath);
            DrawSelectFolder("CSV 저장 경로", ref settingData.csvPath);

            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button("엑셀 파일 검색", GUILayout.Width(120)))
            {
                listSearchedFiles.Clear();
                string[] _arrData = Directory.GetFiles(settingData.GetExcelFullPath());
                for (int _index = 0; _index < _arrData.Length; _index++)
                {
                    string _extention = GetExtensionString(_arrData[_index]);
                    if (_extention == "xls" || _extention == "xlsx")
                        listSearchedFiles.Add(_arrData[_index]);
                }
                multiListView.Reload();
                singleListView.Reload();
            }
            if (GUILayout.Button("네임스페이스 편집", GUILayout.Width(120)))
            {
                if (File.Exists(excelLoaderPath + "Setting/ExcelLoaderNamespace.txt") == false)
                {
                    FileStream _stream = File.Create(excelLoaderPath + "Setting/ExcelLoaderNamespace.txt");
                    _stream.Close();
                    AssetDatabase.Refresh();
                }
                Process.Start(Application.dataPath.Replace("Assets", excelLoaderPath) + "Setting/ExcelLoaderNamespace.txt");
            }
            EditorGUILayout.EndHorizontal();
            eExcelLoaderType _prevLoadType = currentLoadType;
            currentLoadType = (eExcelLoaderType)GUILayout.Toolbar((int)currentLoadType, new string[2] {"다중 선택", "단일 선택"});
            EditorGUILayout.EndVertical();
            //현재 로드 타입이 이전에 타입과 다르면 선택된 내용을 지워준다.
            if (_prevLoadType != currentLoadType)
            {
                selectWorkbook = null;
                selectSheet = null;
                listSelectSheetHeaders.Clear();
                excelSelectID = 0;
                sheetSelectID = 0;
                lastSelectExcelName = string.Empty;
                lastSelectSheetName = string.Empty;
                listExcelSheets.Clear();
                sheetListView.Reload();

                multiListView.SetSelection(new List<int>());
                multiListView.searchString = string.Empty;
                if (currentLoadType == eExcelLoaderType.MultiSelect)
                {
                    searchField.downOrUpArrowKeyPressed -= singleListView.SetFocusAndEnsureSelectedItem;
                    searchField.downOrUpArrowKeyPressed += multiListView.SetFocusAndEnsureSelectedItem;
                }
                else
                {
                    searchField.downOrUpArrowKeyPressed -= multiListView.SetFocusAndEnsureSelectedItem;
                    searchField.downOrUpArrowKeyPressed += singleListView.SetFocusAndEnsureSelectedItem;
                }
            }

            switch (currentLoadType)
            {
                case eExcelLoaderType.MultiSelect:
                    {
                        multiListView.searchString = searchField.OnGUI(new Rect(treeViewPadding, treeViewStartPosY - 20, excelListViewWidth, 20), multiListView.searchString);
                        multiListView.OnGUI(multiSelectListViewRect);

                        GUILayout.Space(treeViewEndPosY + 30);
                        Color _deaultColor = GUI.color;
                        GUI.color = Color.green;
                        if (GUILayout.Button("전체 선택", GUILayout.Width(100)))
                        {
                            multiListView.SelectAllRows();
                        }
                        GUI.color = _deaultColor;
                        GUILayout.BeginHorizontal();
                        GUI.enabled = !isCompiling && multiListViewState.selectedIDs.Count > 0;
                        if (GUILayout.Button("CS 파일 생성/갱신", GUILayout.Width(position.width / 2 - 5)))
                        {
                            EditorUtility.DisplayProgressBar("Work...", "CS파일 생성중...", 0);
                            IList<int> _selects = multiListView.GetSelection();
                            listMultiSelects = new List<int>(_selects);
                            string _namespace = LoadNamespaceText();
                            for (int _index = 0; _index < _selects.Count; _index++)
                            {
                                TreeViewItem _item = multiListView.GetItem(_selects[_index]);
                                EditorUtility.DisplayProgressBar("Work...", string.Format("CS파일 생성중({0})...", _item.displayName), (float)(_index + 1) / _selects.Count);

                                string _sheetName;
                                string _excelFilePath;
                                GetExcelFilePathAndSheetName(_item, out _excelFilePath, out _sheetName);

                                LoadExcel(_excelFilePath);
                                ISheet _tableSheet;
                                List<HeaderData> _headers = LoadSheet(selectWorkbook, _item.displayName, _sheetName, out _tableSheet);
                                string _log = string.Format("File Name = {0}, Sheet Name = {1}", _item.displayName, _sheetName);
                                if (_tableSheet == null)
                                {
                                    UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 엑셀 파일에 해당 시트가 존재하지 않습니다. {0}", _log);
                                    continue;
                                }
                                if (_headers.Count < 1)
                                {
                                    UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 엑셀 파일에 필드가 존재하지 않습니다. {0}", _log);
                                }
                                else
                                {
                                    scriptGenerator.SetScriptGenerator(_sheetName, _namespace, settingData, _headers);
                                    scriptGenerator.DataScriptGenerate();
                                }
                            }
                            EditorUtility.ClearProgressBar();
                            AssetDatabase.Refresh();
                        }

                        //현재 선택한 엑셀파일들중에 CS파일이 생성되지 않은 엑셀이 있는 경우 GUI를 disable한다.
                        bool _guiEnable = !isCompiling && multiListViewState.selectedIDs.Count > 0;
                        if (_guiEnable == true)
                        {
                            IList<int> _selects = multiListView.GetSelection();
                            for (int _index = 0; _index < _selects.Count; _index++)
                            {
                                TreeViewItem _item = multiListView.GetItem(_selects[_index]);
                                string _sheetName;
                                string _excelFilePath;
                                GetExcelFilePathAndSheetName(_item, out _excelFilePath, out _sheetName);
                                if (File.Exists(string.Format("{0}/{1}.cs", settingData.GetClassFullPath(), ScriptGenerator.GetDataName(_sheetName))) == false)
                                {
                                    _guiEnable = false;
                                    break;
                                }
                            }
                        }
                        GUI.enabled = _guiEnable;

                        if (GUILayout.Button("바이너리,CSV 생성/갱신", GUILayout.Width(position.width / 2 - 5)))
                        {
                            IList<int> _selects = multiListView.GetSelection();
                            listMultiSelects = new List<int>(_selects);
                            EditorUtility.DisplayProgressBar("Work...", "바이너리 작성중...", 0);
                            for (int _index = 0; _index < _selects.Count; _index++)
                            {
                                TreeViewItem _item = multiListView.GetItem(_selects[_index]);
                                EditorUtility.DisplayProgressBar("Work...", string.Format("바이너리 작성중({0})...", _item.displayName), (float)(_index + 1) / _selects.Count);

                                string _sheetName;
                                string _excelFilePath;
                                GetExcelFilePathAndSheetName(_item, out _excelFilePath, out _sheetName);

                                LoadExcel(_excelFilePath);
                                ISheet _tableSheet;
                                List<HeaderData> _headers = LoadSheet(selectWorkbook, _item.displayName, _sheetName, out _tableSheet);

                                string _log = string.Format("File Name = {0}, Sheet Name = {1}", _item.displayName, _sheetName);
                                if (_tableSheet == null)
                                {
                                    UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 엑셀 파일에 해당 시트가 존재하지 않습니다. {0}", _log);
                                    continue;
                                }
                                Type _tableType = ScriptGenerator.GetType("ExcelLoader.DataContainer");
                                Type _dataType = ScriptGenerator.GetType(string.Format("{0}Data", _sheetName));
                                if (_dataType == null)
                                {
                                    UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 엑셀 파일에 해당하는 CS파일이 존재하지 않습니다. {0}", _log);
                                    continue;
                                }
                                if (_headers.Count < 1)
                                {
                                    UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 엑셀 파일에 필드가 존재하지 않습니다. {0}", _log);
                                }
                                else
                                {
                                    if(WriteBinary(_headers, _tableSheet, _tableType, _dataType, settingData.GetDataFullPath(), _sheetName))
                                        WriteCSV(_headers, _tableSheet, _dataType, settingData.GetCsvFullPath(), _sheetName);
                                }
                            }
                            EditorUtility.ClearProgressBar();
                            AssetDatabase.Refresh();
                            RefreshGUI();
                        }
                        GUI.enabled = !isCompiling;
                        GUILayout.EndHorizontal();
                    }
                    break;
                case eExcelLoaderType.SingleSelect:
                    {
                        singleListView.searchString = searchField.OnGUI(new Rect(treeViewPadding, treeViewStartPosY - 20, excelListViewWidth, 20), singleListView.searchString);
                        singleListView.OnGUI(singleSelectListViewRect);
                        sheetListView.OnGUI(sheetListViewRect);

                        GUILayout.Space(treeViewEndPosY + 20 + 20);

                        using (new GUILayout.HorizontalScope(EditorStyles.toolbar))
                        {
                            GUILayout.Label("Member", GUILayout.MinWidth(100));
                            GUILayout.FlexibleSpace();
                            string[] names = { "Type", "Array" };
                            int[] widths = { 55, 40 };
                            for (int i = 0; i < names.Length; i++)
                            {
                                GUILayout.Label(new GUIContent(names[i]), GUILayout.Width(widths[i]));
                            }
                        }

                        using (new EditorGUILayout.VerticalScope("box"))
                        {
                            GUILayout.BeginVertical();
                            if (listSelectSheetHeaders != null)
                            {
                                foreach (HeaderData header in listSelectSheetHeaders)
                                {
                                    GUILayout.BeginHorizontal();

                                    EditorGUILayout.LabelField(header.name, GUILayout.MinWidth(100));
                                    GUILayout.FlexibleSpace();

                                    EditorGUILayout.EnumPopup(header.type, GUILayout.Width(60));
                                    GUILayout.Space(20);

                                    EditorGUILayout.Toggle(header.arrayGroup > 0, GUILayout.Width(20));
                                    GUILayout.Space(10);
                                    GUILayout.EndHorizontal();
                                }
                            }
                            EditorGUILayout.EndVertical();
                        }

                        GUILayout.BeginHorizontal();
                        GUI.enabled = !isCompiling && listSelectSheetHeaders.Count > 0;
                        if (GUILayout.Button("CS 파일 생성/갱신", GUILayout.Width(position.width / 2 - 5)))
                        {
                            scriptGenerator.DataScriptGenerate();
                            AssetDatabase.Refresh();
                        }
                        GUI.enabled = !isCompiling && (selectSheet == null ? false : File.Exists(string.Format("{0}/{1}.cs", settingData.GetClassFullPath(), scriptGenerator.dataFileName)));
                        if (GUILayout.Button("바이너리,CSV 생성/갱신", GUILayout.Width(position.width / 2 - 5)))
                        {
                            if (WriteBinary(
                                listSelectSheetHeaders,
                                selectSheet,
                                scriptGenerator.GetTableType(),
                                scriptGenerator.GetDataType(),
                                settingData.GetDataFullPath(),
                                scriptGenerator.dataTableName))
                            {
                                WriteCSV(listSelectSheetHeaders, selectSheet, scriptGenerator.GetDataType(), settingData.GetCsvFullPath(), scriptGenerator.dataTableName);
                            }
                            AssetDatabase.Refresh();
                            RefreshGUI();
                        }
                        GUI.enabled = !isCompiling;
                        GUILayout.EndHorizontal();

                        //if (GUILayout.Button("테이블 로드 테스트"))
                        //{
                        //    string _path = settingData.dataPath + '/' + scriptGenerator.dataTableName + ".bytes";
                        //    _path = _path.Remove(0, _path.IndexOf("Assets"));
                        //    DataContainer _table = DataContainer.LoadTable(AssetDatabase.LoadAssetAtPath<TextAsset>(_path));
                        //}
                    }
                    break;
                default:
                    break;
            }
            EditorGUILayout.EndScrollView();
        }

        /// <summary>
        /// 폴더 선택 GUI 그리는 함수
        /// </summary>
        /// <param name="_pathName"></param>
        /// <param name="_path"></param>
        void DrawSelectFolder(string _pathName, ref string _path)
        {
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField(string.Format("{0}\t:", _pathName), GUILayout.Width(150));
            GUI.enabled = false;
            EditorGUILayout.TextField(settingData.defaultPath + _path);
            GUI.enabled = !isCompiling;
            if (GUILayout.Button("..", GUILayout.Width(30)))
            {
                string _selectPath = EditorUtility.OpenFolderPanel("Select Folder", settingData.defaultPath + _path, "");
                if (string.IsNullOrEmpty(_selectPath) == false)
                {
                    _path = _selectPath.Replace(settingData.defaultPath, "");
                    EditorUtility.SetDirty(settingData);
                    AssetDatabase.SaveAssets();
                }
            }
            EditorGUILayout.EndHorizontal();
        }

        #region TreeView OnClick Event
        /// <summary>
        /// 단일 선택 엑셀 트리뷰 클릭 이벤트
        /// </summary>
        /// <param name="_selectID"></param>
        /// <param name="_excelPath"></param>
        void OnClickSingleSelectExcelList(int _selectID, string _excelPath)
        {
            excelSelectID = _selectID;
            lastSelectExcelName = _excelPath;
            listExcelSheets.Clear();

            if (string.IsNullOrEmpty(_excelPath))
                return;

            _excelPath = _excelPath.Replace('\\', '/');
            EditorUtility.DisplayProgressBar("Excel Load", "", 0f);
            LoadExcel(_excelPath);
            EditorUtility.ClearProgressBar();
            for (int _index = 0; _index < selectWorkbook.NumberOfSheets; _index++)
            {
                listExcelSheets.Add(selectWorkbook.GetSheetName(_index));
            }
            sheetListView.Reload();
        }

        /// <summary>
        /// 시트 트리뷰 클릭 이벤트
        /// </summary>
        /// <param name="_sheetName"></param>
        void OnClickSheetList(int _selectID, string _sheetName)
        {
            sheetSelectID = _selectID;
            lastSelectSheetName = _sheetName;

            if (string.IsNullOrEmpty(_sheetName))
                return;

            var _listHeader = LoadSheet(selectWorkbook, lastSelectExcelName, _sheetName, out selectSheet);
            if (_listHeader != null)
                listSelectSheetHeaders = _listHeader;

            scriptGenerator.SetScriptGenerator(selectSheet.SheetName, LoadNamespaceText(), settingData, listSelectSheetHeaders);
        }
        #endregion TreeView OnClick Event

        /// <summary>
        /// 트리뷰아이템에서 엑셀 파일 경로와 시트이름을 얻는다.
        /// </summary>
        /// <param name="_item"></param>
        /// <param name="_excelFilePath"></param>
        /// <param name="_sheetName"></param>
        void GetExcelFilePathAndSheetName(TreeViewItem _item, out string _excelFilePath, out string _sheetName)
        {
            ExcelLoaderTreeView.ExcelLoaderTreeViewItem _treeViewItem = _item as ExcelLoaderTreeView.ExcelLoaderTreeViewItem;
            _excelFilePath = _treeViewItem.itemName;
            string[] _splits = _excelFilePath.Split('\\');
            _splits = _splits[_splits.Length - 1].Split('.');
            _sheetName = _splits[0];
        }

        /// <summary>
        /// 해당 경로의 엑셀파일을 로드한다
        /// </summary>
        /// <param name="path"></param>
        void LoadExcel(string path)
        {
            try
            {
                using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    string extension = GetExtensionString(path);

                    if (extension == "xls")
                        selectWorkbook = new HSSFWorkbook(fileStream);
                    else if (extension == "xlsx")
                    {
#if UNITY_EDITOR_OSX
                        throw new Exception("xlsx is not supported on OSX.");
#else
                        selectWorkbook = new XSSFWorkbook(fileStream);
#endif
                    }
                    else
                    {
                        throw new Exception("Wrong file.");
                    }
                }
            }
            catch (Exception e)
            {
                UnityEngine.Debug.LogError(e.Message);
            }
        }

        /// <summary>
        /// 지금 로드되있는 엑셀파일에서 해당 시트의 정보를 로드한다.
        /// </summary>
        /// <param name="_sheetName"></param>
        List<HeaderData> LoadSheet(IWorkbook _workBook, string _workBookName, string _sheetName, out ISheet _sheet)
        {
            _sheet = _workBook.GetSheet(_sheetName);
            List<HeaderData> _headers = new List<HeaderData>();

            if (_sheet == null)
                return new List<HeaderData>();

            IRow title = _sheet.GetRow(0);
            if (title == null)
                return new List<HeaderData>();

            for (int i = 0; i < title.LastCellNum; i++)
            {
                var _cell = title.GetCell(i);
                if (_cell == null || string.IsNullOrEmpty(_cell.StringCellValue))
                {
                    UnityEngine.Debug.LogWarningFormat("ExcelLoader Error : 셀이 비어있습니다. {0}칼럼", i);
                    continue;
                }
                string _header = _cell.StringCellValue;
                string[] _headerData = _header.Split('-');
                if (_headerData.Length < 2)
                {
                    UnityEngine.Debug.LogWarningFormat("ExcelLoader Warning : 이 셀에 자료형 구분자가 존재하지 않습니다. {0}칼럼", i);
                    continue;
                }
                else
                {
                    char[] _type = _headerData[1].ToLower().ToCharArray();
                    _type[0] = char.ToUpper(_type[0]);
                    CellType _eType = (CellType)Enum.Parse(typeof(CellType), new string(_type));
                    if (_eType == CellType.None)
                    {
                        UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 엑셀에 잘못된 타입 이름이 들어가있습니다. 테이블={0}, 시트={1}, 필드명={2}", _workBookName, _sheetName, _headerData[1]);
                        return null;
                    }
                    int _arrayGroupID = 0;
                    if (_headerData.Length > 2)
                    {
                        string _arrayGroup = _headerData[2];
                        _arrayGroup = _arrayGroup[_arrayGroup.Length - 1].ToString();
                        if (int.TryParse(_arrayGroup, out _arrayGroupID) == false)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 필드에 배열로 지정했는데 그룹 넘버없습니다. 테이블={0}, 시트={1}, 필드명={2}", _workBookName, _sheetName, _headerData[1]);
                            return null;
                        }
                    }

                    HeaderData _headData = new HeaderData() { type = _eType, name = _headerData[0].Replace(" ", ""), cellColumnIndex = _cell.ColumnIndex };

                    //배열 그룹이 있다면
                    if (_arrayGroupID > 0)
                    {
                        //배열 그룹 아이디가 같은 헤더가 있는지 찾는다.
                        int _arrayGroupHead = _headers.FindIndex(_item => _item.arrayGroup == _arrayGroupID);
                        //있을 경우 그 헤더에 칼럼을 추가해준다.
                        if (_arrayGroupHead != -1)
                        {
                            _headers[_arrayGroupHead].AddArrayData(_cell.ColumnIndex);
                        }
                        //없는 경우 배열데이터로 헤더를 세팅한다.
                        else
                        {
                            _headData.SetArrayData(_arrayGroupID);
                            _headers.Add(_headData);
                        }
                    }
                    else
                    {
                        _headers.Add(_headData);
                    }
                }
            }

            return _headers;
        }

        /// <summary>
        /// 네임스페이스 텍스트를 얻는다.
        /// </summary>
        /// <returns></returns>
        string LoadNamespaceText()
        {
            string _filePath = excelLoaderPath + "Setting/ExcelLoaderNamespace.txt";
            if (File.Exists(_filePath) == false)
            {
                var _file = File.CreateText(_filePath);
                _file.Close();
            }
            return File.ReadAllText(_filePath, Encoding.UTF8);
        }
        /// <summary>
        /// 경로에서 확장자 문자열을 얻는함수
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        string GetExtensionString(string path)
        {
            string ext = Path.GetExtension(path);
            string[] arg = ext.Split(new char[] { '.' });
            return arg[1];
        }

        /// <summary>
        /// Cell 값을 원하는 Type으로 변환하는 함수
        /// </summary>
        /// <param name="_cell"></param>
        /// <param name="_type"></param>
        /// <returns></returns>
        object ConvertFrom(ICell _cell, Type _type)
        {
            object _value = null;

            if (_cell.CellType == NPOI.SS.UserModel.CellType.Blank)
                return _value;

            if (_type == typeof(byte) || _type == typeof(float) || _type == typeof(double) || _type == typeof(short) || _type == typeof(int) || _type == typeof(long))
            {
                if (_cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                {
                    _value = _cell.NumericCellValue;
                }
                else if (_cell.CellType == NPOI.SS.UserModel.CellType.String)
                {
                    if (_type == typeof(byte))
                    {
                        byte _parseValue = 0;
                        if (byte.TryParse(_cell.StringCellValue, out _parseValue) == false)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 잘못된 타입의 값이 들어가있습니다. 시트={0}, 행={1}, 열={2},", _cell.Sheet.SheetName, _cell.RowIndex, _cell.ColumnIndex);
                        }
                        _value = _parseValue;
                    }
                    else if (_type == typeof(float))
                    {
                        float _parseValue = 0;
                        if (float.TryParse(_cell.StringCellValue, out _parseValue) == false)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 잘못된 타입의 값이 들어가있습니다. 시트={0}, 행={1}, 열={2},", _cell.Sheet.SheetName, _cell.RowIndex, _cell.ColumnIndex);
                        }
                        _value = _parseValue;
                    }
                    else if (_type == typeof(double))
                    {
                        double _parseValue = 0;
                        if (double.TryParse(_cell.StringCellValue, out _parseValue) == false)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 잘못된 타입의 값이 들어가있습니다. 시트={0}, 행={1}, 열={2},", _cell.Sheet.SheetName, _cell.RowIndex, _cell.ColumnIndex);
                        }
                        _value = _parseValue;
                    }
                    else if (_type == typeof(short))
                    {
                        short _parseValue = 0;
                        if (short.TryParse(_cell.StringCellValue, out _parseValue) == false)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 잘못된 타입의 값이 들어가있습니다. 시트={0}, 행={1}, 열={2},", _cell.Sheet.SheetName, _cell.RowIndex, _cell.ColumnIndex);
                        }
                        _value = _parseValue;
                    }
                    else if (_type == typeof(int))
                    {
                        int _parseValue = 0;
                        if (int.TryParse(_cell.StringCellValue, out _parseValue) == false)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 잘못된 타입의 값이 들어가있습니다. 시트={0}, 행={1}, 열={2},", _cell.Sheet.SheetName, _cell.RowIndex, _cell.ColumnIndex);
                        }
                        _value = _parseValue;
                    }
                    else if (_type == typeof(long))
                    {
                        long _parseValue = 0;
                        if (long.TryParse(_cell.StringCellValue, out _parseValue) == false)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 잘못된 타입의 값이 들어가있습니다. 시트={0}, 행={1}, 열={2},", _cell.Sheet.SheetName, _cell.RowIndex, _cell.ColumnIndex);
                        }
                        _value = _parseValue;
                    }
                }
                else if (_cell.CellType == NPOI.SS.UserModel.CellType.Formula)
                {
                    if (_type == typeof(float))
                        _value = Convert.ToSingle(_cell.NumericCellValue);
                    if (_type == typeof(double))
                        _value = Convert.ToDouble(_cell.NumericCellValue);
                    if (_type == typeof(short))
                        _value = Convert.ToInt16(_cell.NumericCellValue);
                    if (_type == typeof(int))
                        _value = Convert.ToInt32(_cell.NumericCellValue);
                    if (_type == typeof(long))
                        _value = Convert.ToInt64(_cell.NumericCellValue);
                }
            }
            else if (_type == typeof(string) || _type.IsArray)
            {
                if (_cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                    _value = _cell.NumericCellValue;
                else
                    _value = _cell.StringCellValue;
            }
            else if (_type == typeof(bool))
            {
                try
                {
                    _value = _cell.BooleanCellValue;
                }
                catch( Exception e)
                {
                    bool temp;
                    if( Boolean.TryParse(_cell.StringCellValue, out temp) )
                    {
                        _value = temp;
                    }
                    else
                    {
                        throw e;
                    }
                }
                
            }
            else if (_type.IsGenericType && _type.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                var nc = new NullableConverter(_type);
                return nc.ConvertFrom(_value);
            }
            else if (_type.IsEnum)
            {
                if (_cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                {
                    int _intValue = 0;
                    if (Int32.TryParse(_cell.NumericCellValue.ToString(), out _intValue) == false)
                    {
                        UnityEngine.Debug.LogError("ExcelLoader Error : Enum타입이 잘못 기입되어있습니다.");
                    }
                    else
                        _value = Enum.ToObject(_type, _intValue);
                }
                else
                {
                    _value = Enum.Parse(_type, _cell.StringCellValue, true);
                }
                return _value;
            }

            if (_type.IsArray)
            {
                if (_type.GetElementType() == typeof(float))
                    return ConvertToArray<float>((string)_value);
                else if (_type.GetElementType() == typeof(double))
                    return ConvertToArray<double>((string)_value);
                else if (_type.GetElementType() == typeof(short))
                    return ConvertToArray<short>((string)_value);
                else if (_type.GetElementType() == typeof(int))
                    return ConvertToArray<int>((string)_value);
                else if (_type.GetElementType() == typeof(long))
                    return ConvertToArray<long>((string)_value);
                else if (_type.GetElementType() == typeof(string))
                    return ConvertToArray<string>((string)_value);
            }

            return Convert.ChangeType(_value, _type);
        }

        /// <summary>
        /// 엑셀 테이블을 테이블 클래스와 지정된 데이터 타입으로 저장해서 바이너리로 변환하는 함수
        /// </summary>
        /// <param name="_tableType">테이블 타입</param>
        /// <param name="_dataType">데이터 타입</param>
        /// <param name="_savepath">저장 경로</param>
        /// <param name="_filename">파일 이름</param>
        bool WriteBinary(List<HeaderData> _headerData, ISheet _sheet, Type _tableType, Type _dataType, string _savepath, string _filename)
        {
            Type _listType = typeof(List<>).MakeGenericType(_dataType);
            var _listInstance = Activator.CreateInstance(_listType);
            IList _list = (IList)_listInstance;

            List<PropertyInfo> _dataPropertyInfo = _dataType.GetProperties().ToList();
            _dataPropertyInfo = _dataPropertyInfo.FindAll(_item => _headerData.Find(_header => _header.GetMemberName() == _item.Name) != null);
            Dictionary<object, iTableDataBase> _checkKeyData = new Dictionary<object, iTableDataBase>();
            foreach (IRow _row in _sheet)
            {
                if (_row.RowNum < 1)
                {
                    continue;
                }

                ICell _keyCell = _row.GetCell(0);
                if (_keyCell == null || _keyCell.CellType == NPOI.SS.UserModel.CellType.Blank)
                    continue;

                object[] _propertyDatas = new object[_headerData.Count];
                for (int _index = 0; _index < _headerData.Count; _index++)
                {
                    ICell _cell = _row.GetCell(_headerData[_index].cellColumnIndex);

                    if (_cell == null || _cell.CellType == NPOI.SS.UserModel.CellType.Blank)
                        continue;

                    PropertyInfo _property = _dataPropertyInfo.Find(_item => _item.Name == _headerData[_index].GetMemberName());
                    if (_property == null)
                    {
                        EditorUtility.DisplayDialog("오류", string.Format("데이터 클래스에 헤더에 맞는 변수가 없습니다.\n시트={0}\n변수명={1}", _cell.Sheet.SheetName, _headerData[_index].name), "확인");
                        return false;
                    }
                    if (_headerData[_index].arrayGroup > 0)
                    {
                        List<int> _listArrayColurm = _headerData[_index].GetArrayColurms();
                        Type _elementType = _property.PropertyType.GetElementType();
                        var _arrayDatas = Array.CreateInstance(_elementType, _listArrayColurm.Count);
                        for (int _index_2 = 0; _index_2 < _listArrayColurm.Count; _index_2++)
                        {
                            ICell _arrayDataCell = _row.GetCell(_listArrayColurm[_index_2]);
                            _arrayDatas.SetValue(ConvertFrom(_arrayDataCell, _elementType), _index_2);
                        }
                        _propertyDatas[_index] = _arrayDatas;
                    }
                    else
                    {
                        _propertyDatas[_index] = ConvertFrom(_cell, _property.PropertyType);
                    }
                }
                iTableDataBase _data = (iTableDataBase)Activator.CreateInstance(_dataType, _propertyDatas);
                _list.Add(_data);
                if (_checkKeyData.ContainsKey(_data.GetKey()) == false)
                {
                    _checkKeyData.Add(_data.GetKey(), _data);
                }
                else
                {
                    EditorUtility.DisplayDialog("오류", string.Format("같은 키값이 존재합니다.\n테이블명={0}\n 중복 키값={1}", _sheet.SheetName, _data.GetKey()), "확인");
                    return false;
                }
            }

            var _var = Activator.CreateInstance(_tableType, _list, _dataType);
            DataContainer _table = (DataContainer)_var;
            StreamWriter sWriter = new StreamWriter(_savepath + string.Format("/{0}.bytes", _filename));            
            BinaryFormatter bin = new BinaryFormatter();
            bin.Serialize(sWriter.BaseStream, _table);
            sWriter.Close();
            return true;
        }

        void WriteCSV(List<HeaderData> _headerData, ISheet _sheet, Type _dataType, string _savepath, string _filename)
        {
            List<PropertyInfo> _dataPropertyInfo = _dataType.GetProperties().ToList();
            _dataPropertyInfo = _dataPropertyInfo.FindAll(_item => _headerData.Find(_header => _header.GetMemberName() == _item.Name) != null);

            using (var wtr = new StreamWriter(_savepath + string.Format("/{0}.csv", _filename), false, Encoding.UTF8))
            {
                foreach (IRow _row in _sheet)
                {
                    if (_row.RowNum < 1)
                        continue;

                    ICell _keyCell = _row.GetCell(0);
                    if (_keyCell == null || _keyCell.CellType == NPOI.SS.UserModel.CellType.Blank)
                        continue;

                    string[] _stringCells = new string[_headerData.Count];
                    for (int _index = 0; _index < _headerData.Count; _index++)
                    {
                        ICell _cell = _row.GetCell(_headerData[_index].cellColumnIndex);

                        if (_cell == null || _cell.CellType == NPOI.SS.UserModel.CellType.Blank)
                            continue;

                        PropertyInfo _property = _dataPropertyInfo.Find(_item => _item.Name == _headerData[_index].GetMemberName());
                        if (_property == null)
                        {
                            UnityEngine.Debug.LogErrorFormat("ExcelLoader Error : 데이터 클래스에 헤더에 맞는 변수가 없습니다. 시트={0}, 변수명={1}", _cell.Sheet.SheetName, _headerData[_index].name);
                            return;
                        }

                        object _data;
                        if (_headerData[_index].arrayGroup > 0)
                        {
                            List<int> _listArrayColurm = _headerData[_index].GetArrayColurms();
                            Type _elementType = _property.PropertyType.GetElementType();
                            string[] _arrayDatas = new string[_listArrayColurm.Count];
                            for (int _index_2 = 0; _index_2 < _listArrayColurm.Count; _index_2++)
                            {
                                ICell _arrayDataCell = _row.GetCell(_listArrayColurm[_index_2]);
                                _arrayDatas[_index_2] = ConvertFrom(_arrayDataCell, _elementType).ToString();
                            }
                            _data = string.Join(",", _arrayDatas);
                        }
                        else
                        {
                            _data = ConvertFrom(_cell, _property.PropertyType);
                        }

                        _stringCells[_index] = _data.ToString();
                    }
                    wtr.WriteLine(string.Join(",", _stringCells));
                }
            }
        }

        /// <summary>
        /// 스트링 값을 제네릭 타입의 배열로 만들어서 반환하는 함수
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        T[] ConvertToArray<T>(string value)
        {
            object[] temp = value.Split(',');
            T[] result = temp.Select(e => Convert.ChangeType(e, typeof(T)))
                                 .Select(e => (T)e).ToArray();
            return result;
        }
    }

    /// <summary>
    /// 엑셀 테이블의 칼럼 데이터
    /// </summary>
    public class HeaderData
    {
        public CellType type;
        public string name;
        public int cellColumnIndex;
        public int arrayGroup { get; private set; }
        List<int> listArrayDataColumnIndex;

        /// <summary>
        /// 칼럼의 변수 타입명
        /// </summary>
        /// <returns></returns>
        public string TypeName()
        {
            string _typeName;
            if (type == CellType.Enum)
            {
                _typeName = 'e' + name;
            }
            else
            {
                _typeName = type.ToString().ToLower();
            }

            if (arrayGroup > 0)
                _typeName = _typeName + "[]";

            return _typeName;
        }

        /// <summary>
        /// 칼럼 명
        /// </summary>
        /// <returns></returns>
        public string GetMemberName()
        {
            return "Data_" + name;
        }

        /// <summary>
        /// 그룹 아이디를 통한 배열 데이터
        /// </summary>
        /// <param name="_groupID"></param>
        public void SetArrayData(int _groupID)
        {
            arrayGroup = _groupID;
            listArrayDataColumnIndex = new List<int>();
            listArrayDataColumnIndex.Add(cellColumnIndex);

            int _removeNumberIndex = 0;
            for (int _index = name.Length - 1; _index >= 0; _index--)
            {
                if (char.IsNumber(name[_index]))
                {
                    _removeNumberIndex = _index;
                }
                else
                    break;
            }
            name = name.Substring(0, _removeNumberIndex) + 's';
        }

        /// <summary>
        /// 배열로 묶인 칼럼 인덱스 추가.
        /// </summary>
        /// <param name="_colurmIndex"></param>
        public void AddArrayData(int _colurmIndex)
        {
            listArrayDataColumnIndex.Add(_colurmIndex);
        }

        /// <summary>
        /// 배열로 묶인 칼럼 인덱스 리스트를 얻는다.
        /// </summary>
        /// <returns></returns>
        public List<int> GetArrayColurms() { return listArrayDataColumnIndex; }
    }
}