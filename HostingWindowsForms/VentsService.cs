using System.Collections.Generic;
using ConecctorOneC;
using System.ServiceModel;
using System;
using System.Windows.Forms;

namespace HostingWindowsForms
{
    [ServiceContract]
    public interface I1cService
    {
        [OperationContract]
        bool AuthenticationUser(string UserName, string Password);

        [OperationContract]
        List<Connection.ClassifierMeasure> GetClassifierMeasureList();

        [OperationContract]
        List<Connection.Nomenclature> SearchNomenclatureByName(string name);
    }
    [ServiceContract]
    public interface IEpdmService
    {
        //[WebInvoke(Method = "GET",
        //           ResponseFormat = WebMessageFormat.Json,
        //           UriTemplate = "SearchDoc/{name}")]
        [OperationContract]
        List<EPDM.SearchColumnName> SearchDoc(string name);

        //[WebInvoke(Method = "GET",
        //           ResponseFormat = WebMessageFormat.Json,
        //           //UriTemplate = "BOM?filePath = {filePath}&config = {config}")]
        //           UriTemplate = "BOM/{filePath}/{config}")]
        [OperationContract]
        IList<EPDM.BomCells> Bom(string filePath, string config);

        [OperationContract]
        IEnumerable<string> GetConfiguration(string filePath);

        [OperationContract]
        string GetLink(string fileName);
    }
    [ServiceContract]
    public interface ITaskService
    {
        [OperationContract]
        void AddTaskList(List<Data.SqlQuery.TaskParam> list);
    }
    [ServiceContract]
    public interface ITaskTest
    {
        [OperationContract]
        void Test();
    }   
    public class VentsService : I1cService, IEpdmService, ITaskService, ITaskTest
    {
        #region Authentication
        public bool AuthenticationUser(string UserName, string Password)
        {
            var status = new Authorization.Authentication();
            var statusReturn = status.AuthenticateUser(UserName, Password);

            return statusReturn;
        }
        #endregion
        #region EPDM
            #region Fields

            const int BomId = 8;

            #endregion
            public IEnumerable<string> GetConfiguration(string filePath)
            {
                var epdmClass = new EPDM();

                var enumString = epdmClass.GetConfiguration(filePath);

                return enumString;
            }
            public List<EPDM.SearchColumnName> SearchDoc(string name)
            {
                var epdmClass = new EPDM();
                var searchList = epdmClass.SearchDoc(name);

                return searchList;
            }
            public IList<EPDM.BomCells> Bom(string filePath, string config)
            {
                var bomClass = new EPDM
                {
                    BomId = BomId,
                    AssemblyPath = filePath,
                };
                return bomClass.BomList(filePath, config);
            }           
            public string GetLink(string fileName)
            {
                var strLink = new EPDM();
                var str = strLink.GetLink(fileName);
                return str;
            }
            #region Task      
            public void AddTaskList(List<Data.SqlQuery.TaskParam> list)
            {
                var sql = new Data.SqlQuery();
                sql.AddTaskList(list);
            }
            #endregion
        #endregion
        #region 1C
        public List<Connection.ClassifierMeasure> GetClassifierMeasureList()
        {
            var conOneC = new Connection();
            //conOneC.ConnectionString();
            var getList = conOneC.ClassifierMeasureList();
            return getList;
        }
        public List<Connection.Nomenclature> SearchNomenclatureByName(string name)
        {
            var conOneC = new Connection();
            //conOneC.ConnectionString();
            var getList = conOneC.SearchNomenclatureByName(name);
            return getList;
        }
        #endregion
        public void Test()
        {
            MessageBox.Show("!");
        }

    }
}