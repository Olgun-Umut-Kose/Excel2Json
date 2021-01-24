
using System.Data.Common;


namespace Excel2Json.Abstract
{
    public interface IQuery
    {
        void SetObjects(DbDataRecord row);

        IQuery InitObjects(DbDataRecord row);
        IQuery GetObjects();
        
    }
}
