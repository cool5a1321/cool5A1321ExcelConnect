using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Prospecta.ConnektHub.Models;
using Prospecta.ConnektHub.Services.Translation;
using System.Collections.Generic;

namespace Prospecta.ConnektHub.Controllers
{
    public class TranslationController
    {
        private ITranslationService _translationService;
        #region Constructors
        public TranslationController(ITranslationService translationService)
        {
            _translationService = translationService;
        }
        #endregion
        #region Public Methods
        public List<TranslationData> GetFieldIdNDescriptionInEnglish(string moduleId)
        {
            var lstTranslationData = new List<TranslationData>();
            var jsonString = _translationService.GetFieldIdNDescriptionInEnglish(moduleId);
            if (!string.IsNullOrEmpty(jsonString))
            {
                dynamic dictFieldIdNDescription = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

                var data = dictFieldIdNDescription["result"];
                
                foreach (var item in data)
                {
                    var key = item["fieldId"].ToString();
                    var value = item["fieldDescri"];
                    
                    var translationDataItem = new TranslationData { FieldId = key, FieldDescription = value };
                    lstTranslationData.Add(translationDataItem);
                }
            }
            return lstTranslationData;
        }
        #endregion

    }
}
