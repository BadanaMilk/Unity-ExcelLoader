using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using UnityEditor.IMGUI.Controls;
using System;
using System.IO;
using UnityEditor.UI;

namespace ExcelLoader
{
    public class ExcelLoaderTreeView : TreeView
    {
        protected ExcelLoader_Setting settingInfo;

        protected List<string> arrayData;
        protected Action<int, string> onClickEvent;

        protected Texture2D icon_on;
        protected Texture2D icon_off;

        public ExcelLoaderTreeView(TreeViewState _viewState, ExcelLoader_Setting _settingInfo, ref List<string> _listData, Action<int, string> _onClickEvent) : base(_viewState)
        {
            onClickEvent = _onClickEvent;
            arrayData = _listData;
            settingInfo = _settingInfo;

            rowHeight = 20f;
            showAlternatingRowBackgrounds = true;
            showBorder = true;
            useScrollView = true;
            depthIndentWidth = 0;
        }

        public TreeViewItem GetItem(int _id)
        {
            return rootItem.children.Find(_item => _item.id == _id);
        }

        protected override TreeViewItem BuildRoot()
        {
            TreeViewItem _rootItem = new TreeViewItem { id = 0, depth = -1, displayName = "Root" };
            List<TreeViewItem> _listItem = new List<TreeViewItem>();
            int _count = arrayData == null ? 0 : arrayData.Count;
            for (int _index = 0; _index < _count; _index++)
            {
                ExcelLoaderTreeViewItem _item = GetTreeViewItem(_index + 1, 1, arrayData[_index]);
                SetIcon(_item);
                _listItem.Add(_item);
            }
            SetupParentsAndChildrenFromDepths(_rootItem, _listItem);
            return _rootItem;
        }

        protected virtual void SetIcon(ExcelLoaderTreeViewItem _item) { }

        protected virtual ExcelLoaderTreeViewItem GetTreeViewItem(int _id, int _depth, string _name) 
        {
            ExcelLoaderTreeViewItem _item = new ExcelLoaderTreeViewItem(_id, _depth, _name);
            return _item;
        }

        public override void OnGUI(Rect rect)
        {
            base.OnGUI(rect);
        }

        protected override void SingleClickedItem(int id)
        {
            base.SingleClickedItem(id);
            ExcelLoaderTreeViewItem _result = (ExcelLoaderTreeViewItem)FindItem(id, rootItem);
            if (onClickEvent != null && _result != null)
                onClickEvent(id, _result.itemName);
        }

        protected override void RowGUI(RowGUIArgs args)
        {
            var bundleItem = (args.item as ExcelLoaderTreeViewItem);
            CenterRectUsingSingleLineHeight(ref args.rowRect);
            if (args.item.icon == null)
                extraSpaceBeforeIconAndLabel = 16f;
            else
                extraSpaceBeforeIconAndLabel = 0f;

            base.RowGUI(args);
        }

        public class ExcelLoaderTreeViewItem : TreeViewItem
        {
            public string itemName;

            public ExcelLoaderTreeViewItem(int _id, int _depth, string _itemName)
            {
                id = _id;
                depth = _depth;
                itemName = _itemName;
                displayName = _itemName;
            }
        }
    }

    public class ExcelFileTreeView : ExcelLoaderTreeView
    {
        public ExcelFileTreeView(TreeViewState _viewState, ExcelLoader_Setting _settingInfo, ref List<string> _listData, Action<int, string> _onClickEvent) 
            : base(_viewState, _settingInfo, ref _listData, _onClickEvent)
        {
            icon_on = EditorGUIUtility.Load(string.Format("{0}Icon/icon_excel.png", ExcelLoader_Editor.excelLoaderPath)) as Texture2D;
            Reload();
        }
        protected override void SetIcon(ExcelLoaderTreeViewItem _item)
        {
            _item.icon = icon_on;
        }

        protected override ExcelLoaderTreeViewItem GetTreeViewItem(int _id, int _depth, string _name)
        {
            ExcelLoaderTreeViewItem _item = new ExcelLoaderTreeViewItem(_id, _depth, _name);
            _item.displayName = _name.Remove(0, _name.IndexOf('\\') + 1);
            return _item;
        }
    }

    public class ExcelSheetTreeView : ExcelLoaderTreeView
    {
        Texture2D icon_cs_off;

        public ExcelSheetTreeView(TreeViewState _viewState, ExcelLoader_Setting _settingInfo, ref List<string> _listData, Action<int, string> _onClickEvent)
            : base(_viewState, _settingInfo, ref _listData, _onClickEvent)
        {

            icon_cs_off = EditorGUIUtility.Load(string.Format("{0}Icon/icon_cs_off.png", ExcelLoader_Editor.excelLoaderPath)) as Texture2D;
            icon_on = EditorGUIUtility.Load(string.Format("{0}Icon/icon_data_on.png", ExcelLoader_Editor.excelLoaderPath)) as Texture2D;
            icon_off = EditorGUIUtility.Load(string.Format("{0}Icon/icon_data_off.png", ExcelLoader_Editor.excelLoaderPath)) as Texture2D;
            Reload();
        }

        protected override void SetIcon(ExcelLoaderTreeViewItem _item)
        {
            if (File.Exists(settingInfo.GetClassFullPath()+ string.Format("/{0}Data.cs", _item.itemName)))
            {
                if (File.Exists(settingInfo.GetDataFullPath() + string.Format("/{0}.bytes", _item.itemName)))
                    _item.icon = icon_on;
                else
                    _item.icon = icon_off;
            }
            else
                _item.icon = icon_cs_off;
        }
    }
}