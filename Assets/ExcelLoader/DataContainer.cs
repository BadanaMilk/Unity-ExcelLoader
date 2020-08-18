using UnityEngine;
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;

namespace ExcelLoader
{
    [Serializable]    
    public class DataContainer
    {
        [SerializeField]
        Type dataType;

        [SerializeField]
        List<iTableDataBase> listDatas;

        public Type mDataType { get { return dataType; } }

        protected Dictionary<object, iTableDataBase> dictionaryDatas;
        
        public DataContainer(IList _dataList, Type _dataType)
        {
            dataType = _dataType;
            listDatas = _dataList.Cast<iTableDataBase>().ToList();
        }

        /// <summary>
        /// 가지고 있는 데이터 리스트를 딕셔너리로 저장
        /// </summary>
        void Initialize()
        {
            dictionaryDatas = new Dictionary<object, iTableDataBase>();
            for (int _index = 0; _index < listDatas.Count; _index++)
            {
                dictionaryDatas.Add(listDatas[_index].GetKey(), listDatas[_index]);
            }
            //런타임에서 사용하는건 딕셔너리라서 필요없는 리스트는 클리어해준다.
            listDatas.Clear();
        }

        /// <summary>
        /// 키값을 통해 값을 얻는 함수.
        /// </summary>
        /// <typeparam name="TData">테이블 데이터 타입</typeparam>
        /// <param name="_key">키값</param>
        /// <returns>테이블값</returns>
        public TData GetValueFromKey<TData>(object _key) where TData : iTableDataBase
        {
            if (dictionaryDatas.ContainsKey(_key))
                return (TData)dictionaryDatas[_key];
            else
                return default;
        }

        /// <summary>
        /// 키값을 통해 값을 얻는 함수
        /// </summary>
        /// <typeparam name="TData">테이블 데이터 타입</typeparam>
        /// <param name="_key">키값</param>
        /// <param name="_data">데이터</param>
        /// <returns>값의 존재여부</returns>
        public bool TryGetValueFromKey<TData>(object _key, out TData _data) where TData : iTableDataBase
        {
            iTableDataBase _baseData = null;
            if (dictionaryDatas.TryGetValue(_key, out _baseData))
            {
                _data = (TData)_baseData;
                return true;
            }
            else
            {
                _data = default;
                return false;
            }
        }

        /// <summary>
        /// 테이블을 검색해서 원하는 값을 얻는 함수.
        /// </summary>
        /// <typeparam name="TData">테이블 데이터 타입</typeparam>
        /// <param name="_predicate">검색 비교 조건</param>
        /// <param name="_result">테이블 데이터</param>
        /// <returns>값의 존재여부</returns>
        public bool TryGetSearchValue<TData>(Predicate<TData> _predicate, out TData _result) where TData : iTableDataBase
        {
            var keys = from entry in dictionaryDatas
                       where _predicate((TData)entry.Value)
                       select entry.Key;

            _result = default;
            if (keys.Count() > 0)
            {
                _result = (TData)dictionaryDatas[keys.First()];
                return true;
            }
            return false;
        }

        /// <summary>
        /// 바이너리로 저장한 테이블 에셋을 실사용 테이블 클래스 오브젝트로 로드하는 함수
        /// </summary>
        /// <param name="_asset"></param>
        /// <returns></returns>
        public static DataContainer LoadTable(TextAsset _asset)
        {
            MemoryStream _ms = new MemoryStream(_asset.bytes);
            BinaryFormatter bin = new BinaryFormatter();
            DataContainer _table = (DataContainer)bin.Deserialize(_ms);
            _ms.Close();
            _table.Initialize();
            return _table;
        }
    }

    public interface iTableDataBase
    {
        object GetKey();

        string GetTableName();
    }
}