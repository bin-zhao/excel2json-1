using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace excel2json
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class JsonExporter
    {
//         Dictionary<string, Dictionary<string, object>> m_data_dict;
        List<object> m_data_array;

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="sheet">ExcelReader创建的一个表单</param>
        /// <param name="headerRows">表单中的那几行是表头</param>
        public JsonExporter(DataTable sheet, int headerRows, bool lowcase)
        {
            if (sheet.Columns.Count <= 0)
                return;
            if (sheet.Rows.Count <= 0)
                return;

//             m_data_dict = new Dictionary<string, Dictionary<string, object>>();
            m_data_array = new List<object>();

            //--以第一列为ID，转换成ID->Object的字典
            int firstDataRow = headerRows - 1;
            DataRow clientDataTypeRow = sheet.Rows[1];  // 客户端解析数据的类型行
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];
                string ID = row[sheet.Columns[0]].ToString();
                if (ID.Length <= 0)
                    continue;

                var rowData = new Dictionary<string, object>();
                foreach (DataColumn column in sheet.Columns)
                {
                    object value = row[column];
                    object dataType = clientDataTypeRow[column];
                    value = StringValue2TypeValue(value, dataType);
                    // 表头自动转换成小写
                    string fieldName = column.ToString();
                    if (lowcase)    
                        fieldName = fieldName.ToLower();

                    if (!string.IsNullOrEmpty(fieldName))
                        rowData[fieldName] = value;
                }

//                 m_data_dict[ID] = rowData;
                m_data_array.Add(rowData);
            }
        }

        // 获取数组元素类型
        private string GetArrayItemType(string dataTypeStr)
        {
            string itemTypeStr;
            int posBeg = dataTypeStr.IndexOf('(');
            int posEnd = dataTypeStr.IndexOf(')');
            if (posBeg == -1 || posEnd == -1)
            {
                itemTypeStr = "int";
            }
            else
            {
                itemTypeStr = dataTypeStr.Substring(posBeg + 1, posEnd - posBeg - 1);
            }
            return itemTypeStr;
        }

        // 获取字典key-value类型
        private string[] GetDictItemType(string dataTypeStr)
        {
            string itemTypeStr;
            int posBeg = dataTypeStr.IndexOf('(');
            int posEnd = dataTypeStr.IndexOf(')');
            if (posBeg == -1 || posEnd == -1)
            {
                itemTypeStr = "int:int";
            }
            else
            {
                itemTypeStr = dataTypeStr.Substring(posBeg + 1, posEnd - posBeg - 1);
            }
            return itemTypeStr.Split(':');
        }

        // 字符串转换为指定类型
        private object StringValue2TypeValue(object value, object dataType)
        {
            string dataTypeStr = Convert.ToString(dataType);
            if (string.Compare(dataTypeStr, "int") == 0)
            {
                // 整形
                return Convert.ToInt32(value);
            }
            else if (string.Compare(dataTypeStr, "double") == 0)
            {
                // 浮点数
                return Convert.ToDouble(value);
            }
            else if (dataTypeStr.StartsWith("array"))
            {
                // 数组
                string valueStr = Convert.ToString(value);
                string[] valueStrArray = valueStr.Split(';');
                string itemTypeStr = GetArrayItemType(dataTypeStr);
                // 不会用C#模板，暂时先用if判断
                if (itemTypeStr.CompareTo("int") == 0)
                {
                    // 整形数组
                    List<int> intList = new List<int>();
                    for (int i = 0; i < valueStrArray.Length; ++i)
                    {
                        if (valueStrArray[i].Length <= 0)
                        {
                            // TODO 不应该跳过，要异常处理
                            continue;
                        }
                        intList.Add(Convert.ToInt32(valueStrArray[i]));
                    }
                    return intList;
                }
                else if (itemTypeStr.CompareTo("double") == 0)
                {
                    // 浮点数数组
                    List<double> doubleList = new List<double>();
                    for (int i = 0; i < valueStrArray.Length; ++i)
                    {
                        if (valueStrArray[i].Length <= 0)
                        {
                            // TODO 不应该跳过，要异常处理
                            continue;
                        }
                        doubleList.Add(Convert.ToDouble(valueStrArray[i]));
                    }
                    return doubleList;
                }
                else
                {
                    // 字符串数组（默认）
                    List<string> strList = new List<string>();
                    for (int i = 0; i < valueStrArray.Length; ++i)
                    {
                        if (valueStrArray[i].Length <= 0)
                        {
                            // TODO 不应该跳过，要异常处理
                            continue;
                        }
                        strList.Add(valueStrArray[i]);
                    }
                    return strList;
                }
            }
            else if (dataTypeStr.StartsWith("map"))
            {
                // 字典
                string valueStr = Convert.ToString(value);
                string[] valueStrArray = valueStr.Split(';');
                string[] itemTypeStr = GetDictItemType(dataTypeStr);
                // 不会用C#模板，暂时先用if判断
                if (itemTypeStr[0].CompareTo("int") == 0 && itemTypeStr[1].CompareTo("int") == 0)
                {
                    // 整形-整形字典
                    List<List<int>> intintDict = new List<List<int>>();
                    for (int i = 0; i < valueStrArray.Length; ++i)
                    {
                        string[] pairStr = valueStrArray[i].Split('_');
                        List<int> pair = new List<int>();
                        pair.Add(Convert.ToInt32(pairStr[0]));
                        pair.Add(Convert.ToInt32(pairStr[1]));
                        intintDict.Add(pair);
                    }
                    return intintDict;
                }
            }
            return value;
//             if (value.GetType() == typeof(double))
//             {
//                 // double
//                 // 去掉数值字段的“.0”
//                 double num = (double)value;
//                 if ((int)num == num)
//                     value = (int)num;
//             }
        }

        // 获取array版输出文件路径
        private string GetFilePathArray(string filePath)
        {
//             string fileDir = Path.GetDirectoryName(filePath);
//             string fileName = Path.GetFileNameWithoutExtension(filePath);
//             string fileExt = Path.GetExtension(filePath);
//             return Path.Combine(fileDir, fileName + "_array" + fileExt);
            return filePath;
        }

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            if (m_data_array == null)
                throw new Exception("JsonExporter内部数据为空。");
//             if (m_data_dict == null)

            //-- 转换为JSON字符串
//             string json = JsonConvert.SerializeObject(m_data_dict, Formatting.Indented);
            string json_array = JsonConvert.SerializeObject(m_data_array, Formatting.Indented);

            //-- 保存文件
//             using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
//             {
//                 using (TextWriter writer = new StreamWriter(file, encoding))
//                     writer.Write(json);
//             }
            string filePathArray = GetFilePathArray(filePath);
            using (FileStream file = new FileStream(filePathArray, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(json_array);
            }
        }
    }
}
