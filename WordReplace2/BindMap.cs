using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.Serialization;

namespace WordReplace
{
    class BindDoesNotExistException : ApplicationException
    {
        public BindDoesNotExistException() { }

        public BindDoesNotExistException(string message) : base(message) { }

        public BindDoesNotExistException(string message, Exception inner) : base(message, inner) { }

        protected BindDoesNotExistException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    class BindMap
    {
        private DataRow Row;
        private DataColumnCollection Cols;

        public BindMap(DataRow row, DataColumnCollection cols)
        {
            Row = row;
            Cols = cols;
        }

        public string Get(string key)
        {
            var ind = Cols.IndexOf(key);
            if (ind == -1)
            {
                string msg = String.Format("DataColumn: {0} doesn't exists", key);
                throw new BindDoesNotExistException(msg);
            }
            var col = Cols[ind];
            try
            {
                var newValue = Row[col];
                if (newValue.GetType() == typeof(DateTime))
                {
                    DateTime dt = (DateTime)newValue;
                    if (dt.TimeOfDay.TotalMilliseconds.Equals(0.0))
                        newValue = dt.ToString("d");
                }
                return newValue.ToString();
            }
            catch (ArgumentException)
            {
                string msg = String.Format("DataColumn: {0} not found in the row", key);
                throw new BindDoesNotExistException(msg);
            }
        }

        public string Get(string key, string DefaultValue)
        {
            try
            {
                return Get(key);
            }
            catch (BindDoesNotExistException)
            {
                return DefaultValue;
            }
        }
    }
}
