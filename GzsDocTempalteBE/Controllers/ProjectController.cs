using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web.Http;
using Novacode;

namespace GzsDocTempalteBE
{
    public class ProjectController : ApiController
    {
        private static readonly string CurrentDir = System.AppDomain.CurrentDomain.BaseDirectory;
        private static readonly Dictionary<string, PurposeModel> PurposeDictionary = new Dictionary<string, PurposeModel>
        {
            {"01.03", new PurposeModel { Code = "А 01.03", Name = "для ведення особистого селянського господарства"} },
            {"02.01", new PurposeModel { Code = "B 02.01", Name = "для будівництва і обслуговування житлового будинку, господарських будівель і споруд (присадибна ділянка)"} }
        };
        private static readonly Dictionary<string, string> LocalHeadTypeDictionary = new Dictionary<string, string>
        {
            {"с.", "Сільський"},
            {"смт.", "Селищний"},
            {"м.", "Міський"}
        };

        // GET api/project
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // POST api/project
        public HttpResponseMessage Post([FromBody]ProjectModel project)
        {
            var templateFile = CurrentDir + "template.docx";
            var resultFile = CurrentDir + "result.docx";
            var replaceDictionary = new Dictionary<string, string>
            {
                {"FullNameWho", $"{project.LastNameWho} {project.FirstNameWho} {project.MiddleNameWho}"},
                {"FullNameWhom", $"{project.LastNameWhom} {project.FirstNameWhom} {project.MiddleNameWhom}"},
                {"ShortName", $"{project.LastNameWho} {project.FirstNameWho[0]}.{project.MiddleNameWho[0]}."},
                {"Passport", project.Passport},
                {"IssuedDate", project.IssuedDate},
                {"IssuedAuthority", project.IssuedAuthority},
                {"OwnerAddress", project.OwnerAddress},
                {"SexGot", project.OwnerSex == "male" ? "отримав" : "отримала" },
                {"SexWhich", project.OwnerSex == "male" ? "який" : "яка" },
                {"SexInformed", project.OwnerSex == "male" ? "повідомлений" : "повідомлена" },

                {"PurposeCode", PurposeDictionary[project.Purpose].Code},
                {"PurposeFull", PurposeDictionary[project.Purpose].Name},
                {"PropertyArea", project.PropertyArea},
                {"PropertyAddress", GetPropertyAddress(project)},
                {"PropertyRegion", ReplaceWordEnd(project.SettlementRegion, 2)},
                {"Settlement", $"{project.SettlementType} {project.SettlementName}"},
                {"PropertyLocation", project.PropertyLocation},
                {"PropertyOrientation", ReplaceWordEnd(project.PropertyOrientation, 1, "ій")},
                {"PropertyNeighbours", project.PropertyNeighbours},
                {"BorderSignsCount", project.BorderSignsCount},

                {"LocalHeadType", LocalHeadTypeDictionary[project.SettlementType]},
                {"LocalGovernmentHead", project.LocalGovernmentHead},
                {"LocalGovernmentWhose", GetLocalGovernmentWhose(project.LocalGovernmentName)},
                {"SessionResolution", $"від {project.ResolutionDate} року №{project.ResolutionNumber}"},
                {"ResolutionNumber", project.ResolutionNumber},

                {"DeveloperName", project.DeveloperName},
                {"DeveloperEngineerName", project.DeveloperEngineerName},
                {"ContractDate", project.ContractDate},
                {"ContractNumber", project.ContractNumber}
            };

            using (var document = DocX.Load(templateFile))
            {
                foreach (var replaceItem in replaceDictionary)
                {
                    document.ReplaceText($"$[{replaceItem.Key}]", replaceItem.Value, false, RegexOptions.IgnoreCase);
                }
                document.SaveAs(resultFile);
            }
            
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(resultFile, FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.Add("Content-Disposition", "attachment;filename=\"project.docx\"");
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            return result;
        }

        private static string GetPropertyAddress(ProjectModel project)
        {
            var result = $"{project.SettlementType} {project.SettlementName}";
            if (project.PropertyAddressType != "" && project.PropertyAddressName != "")
            {
                result += $", {project.PropertyAddressType} {project.PropertyAddressName}";
            }
            if (project.PropertyAddressBuilding != "")
            {
                result += $", {project.PropertyAddressBuilding}";
            }
            if (project.PropertyAddressBlock != "")
            {
                result += $"/{project.PropertyAddressBlock}";
            }
            return result;
        }

        private static string GetLocalGovernmentWhose(string localGovernmentName)
        {
            var words = localGovernmentName.Split(' ');
            words[0] = ReplaceWordEnd(words[0], 1, "ої");
            words[1] = ReplaceWordEnd(words[1], 1, "ої");
            words[2] = ReplaceWordEnd(words[2], 1, "и");
            return string.Join(" ", words);
        }
        private static string ReplaceWordEnd(string word, int charCount, string replacement = "")
        {
            return word.Remove(word.Length - charCount) + replacement;
        }
    }
}
