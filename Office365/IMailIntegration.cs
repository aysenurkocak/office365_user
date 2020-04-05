using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace Office365
{
    [ServiceContract]
    public interface IMailIntegration
    {
        [OperationContract]
        [WebGet(
        BodyStyle = WebMessageBodyStyle.Wrapped,
        RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json,
        UriTemplate = "/GetUser/{mailNickname}")]
        User GetUser(string mailNickname);

        [OperationContract]
        [WebInvoke(
         Method = "POST",
         BodyStyle = WebMessageBodyStyle.Bare,
         RequestFormat = WebMessageFormat.Json,
         ResponseFormat = WebMessageFormat.Json,
         UriTemplate = "/UserCreate")]
         Result UserCreate(User user);

        [OperationContract]
        [WebGet(
        BodyStyle = WebMessageBodyStyle.Wrapped,
        RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json,
        UriTemplate = "/UserDelete/{mail}")]
        Result UserDelete(string mail);




    }
}
