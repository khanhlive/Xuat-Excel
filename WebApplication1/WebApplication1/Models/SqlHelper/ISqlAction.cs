using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApplication1.Models.SqlHelper
{
    public interface ISqlAction<T,KeyType>
    {
        void Insert(T item);
        void Insert();
        DataTable GetList();
        string Update(T item);
        string Update();
        string Delete();
        string Delete(T item);
        T GetItem(KeyType key);
        T GetItem();
    }
}
