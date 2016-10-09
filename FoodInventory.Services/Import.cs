using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodInventory.Services
{
    public class Import<T> where T : class, new()
    {
        private DataSet _dataSet;
        private HttpPostedFile _file;
        private List<string> _expectedColumns;
        private List<string> _missingHeaders;
        private List<T> _invalidRows;
        private Type type = typeof(T);
        public Import()
        {
            GetRequiredColumns();
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
        }
        public IEnumerable<T> GetItems(HttpPostedFile file)
        {
            _file = file;
            _dataSet = GetDataSet();
            CheckFileColumns();

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

            List<T> obj = new List<T>();

            for (var r = 0; r < _dataSet.Tables[0].Rows.Count; r++)
            {
                DataRow row = _dataSet.Tables[0].Rows[r];
                T product = new T();

                for (var c = 0; c < _expectedColumns.Count; c++)
                {
                    string value = row.Table.Rows[r][_expectedColumns[c]].ToString();
                    if (value != null)
                    {
                        PropertyInfo pInfo = product.GetType().GetProperty(_expectedColumns[c]);
                        if (Nullable.GetUnderlyingType(pInfo.PropertyType) == null)
                        {
                            pInfo.SetValue(product, Convert.ChangeType(value, pInfo.PropertyType), null);
                        }
                        else
                        {
                            // TODO:
                            if (pInfo.PropertyType == typeof(Nullable<DateTime>))
                            {
                                DateTime date;
                                var convertedVal = DateTime.TryParse(value, out date) ? (Nullable<DateTime>)date : null;
                                pInfo.SetValue(product, convertedVal);
                            }
                        }
                    }
                }

                obj.Add(product);
            }

            return obj.AsEnumerable();
        }
    }
}
