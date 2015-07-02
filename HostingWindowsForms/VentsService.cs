using System.Collections.Generic;
using ConecctorOneC;

namespace HostingWindowsForms
{
    public class VentsService : IVentsService
    {
        List<Connection.ClassifierMeasure> IVentsService.GetList()
        {
            var conOneC = new Connection();

            conOneC.ConnectionString();

            var getList = conOneC.ClassifierMeasureList();

            return getList;
            
        }
    }
}
