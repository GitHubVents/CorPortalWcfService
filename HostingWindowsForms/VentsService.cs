using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using ConecctorOneC;

namespace HostingWindowsForms
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in both code and config file together.
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
