using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;

namespace FoodInventory.API.Utilities
{
    public class Import<T> where T : class, new()
    {
        private DataSet _dataSet;
        private HttpPostedFile _file;
        private List<string> _expectedColumns;
        private List<string> _missingHeaders;
        public List<T> _invalidRows;
        public List<T> _validRows;
        private Type type = typeof(T);
        public Import(HttpPostedFile file)
        {
            _invalidRows = new List<T>();
            _validRows = new List<T>();
            _file = file;
            _dataSet = GetDataSet();
        }

        private DataSet GetDataSet()
        {
            DataSet result = null;
            using (Stream stream = _file.InputStream)
            {
                IExcelDataReader reader;

                if (_file.FileName.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (_file.FileName.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                else
                {
                    throw new Exception("File Not Supported");
                }

                reader.IsFirstRowAsColumnNames = true;

                result = reader.AsDataSet();
            }

            return result;
        }

        private void GetRequiredColumns()
        {
            _expectedColumns = new List<string>();
            foreach (PropertyInfo prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                _expectedColumns.Add(prop.Name);
            }
        }
        private void CheckFileColumns()
        {
            int idx = 0;
            DataRow headerDataRow = _dataSet.Tables[0].Rows[idx];
            _missingHeaders = new List<string>();

            for (var c = 0; c < _expectedColumns.Count; c++)
            {
                if (!headerDataRow.Table.Columns.Contains(_expectedColumns[c]))
                {
                    _missingHeaders.Add(_expectedColumns[c]);
                }
            }


            if (_missingHeaders.Count > 0)
            {
                StringBuilder message = new StringBuilder();
                message.Append("Missing headers: ");
                foreach (string header in _missingHeaders)
                {
                    message.Append(header);
                    message.Append(",");
                }

                throw new Exception(message.ToString());
            }
        }
        public void ValidateItems()
        {
            GetRequiredColumns();
            CheckFileColumns();

            for (var r = 0; r < _dataSet.Tables[0].Rows.Count; r++)
            {
                DataRow row = _dataSet.Tables[0].Rows[r];
                T product = new T();
                bool validRow = true;

                for (var c = 0; c < _expectedColumns.Count; c++)
                {
                    string value = row.Table.Rows[r][_expectedColumns[c]].ToString();
                    if (value != null)
                    {
                        PropertyInfo pInfo = product.GetType().GetProperty(_expectedColumns[c]);
                        Type t = pInfo.PropertyType;
                        
                        // check if Nullable<T> or not
                        if (Nullable.GetUnderlyingType(pInfo.PropertyType) == null)
                        {
                            var convertedValue = ConvertToType(value, t);
                            if (convertedValue != null)
                                pInfo.SetValue(product, convertedValue, null);
                            else
                            {
                                validRow = false;
                                break;
                            }
                        }
                        else
                        {
                            // TODO: need more generics??
                            if (t == typeof(Nullable<DateTime>))
                            {
                                Nullable<DateTime> date = null;
                                if (value != "")
                                {
                                    double dd;
                                    bool validDate = double.TryParse(value, out dd);
                                    if(!validDate)
                                    {
                                        validRow = false;
                                        break;
                                    }

                                    date = DateTime.FromOADate(dd);
                                    pInfo.SetValue(product, date);
                                }
                                else
                                {
                                    pInfo.SetValue(product, date);
                                }
                            }
                        }
                    }
                }

                if (validRow)
                    _validRows.Add(product);
                else
                    _invalidRows.Add(product);
            }
        }
        
        public static object ConvertToType(string value, Type type)
        {
            if (type == null) return null;
            System.ComponentModel.TypeConverter conv = System.ComponentModel.TypeDescriptor.GetConverter(type);
            if (conv.CanConvertFrom(typeof(string)))
            {
                try
                {
                    return conv.ConvertFrom(value);
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }
    }
}