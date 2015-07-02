using System.ServiceModel;
using ConecctorOneC;
using System.Collections.Generic;

namespace HostingWindowsForms
{
    [ServiceContract]
    public interface IVentsService
    {
        [OperationContract]
        List<Connection.ClassifierMeasure> GetList();
    }
}
