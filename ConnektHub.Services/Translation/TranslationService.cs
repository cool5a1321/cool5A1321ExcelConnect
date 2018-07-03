using Prospecta.ConnektHub.Services.HttpService;
namespace Prospecta.ConnektHub.Services.Translation
{
    public class TranslationService : ITranslationService
    {
        private IHttpRequest _httpRequest;
        public TranslationService(IHttpRequest httpRequest)
        {
            _httpRequest = httpRequest;
        }
        public string GetFieldIdNDescriptionInEnglish(string moduleId)
        {
            //var url = "restUserValidation/userValidation?userId=" + userName + "&password=" + password;
            //return _httpRequest.HttpGet(url);
            string jsonString = string.Empty;
            jsonString = "{\"result\":[{\"fieldId\": \"OPS_TAB\",\"fieldDescri\": \"OPS Table\"},{\"fieldId\": \"QA1_TAB\",\"fieldDescri\": \"QA1 table\"},{\"fieldId\": \"qa_col2\",\"fieldDescri\": \"qa_col2\"},{\"fieldId\": \"QA_TAB\",\"fieldDescri\": \"QA Table\"}]}";
            return jsonString;
        }
    }
}