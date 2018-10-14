using SPO.ClientManager.Model;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPO.ClientManager
{
    public class FixLookupSiteColumn
    {
        public static void UpdateLookUpSiteColumn()
        {
            var context = AuthHelper.GetClientContext();
            Web oWeb = context.Web;

            var listOfLookUpSiteColumnDetails = GetLookUpSiteColumnInfo(context, oWeb);
            foreach(var x in listOfLookUpSiteColumnDetails)
            {
                Field oField = context.Web.Fields.GetByInternalNameOrTitle(x.Name);
                context.Load(oField, f => f.SchemaXml);

                context.ExecuteQuery();

                oField.SchemaXml = "<Field Type='Lookup' WebId='" + oWeb.Id + "' Name ='" + x.Name + "' StaticName='" + x.Name + "' DisplayName='" + x.DisplayName + "' List = '" + x.List + "' ShowField = '" + x.ShowField + "' /> ";
                context.Load(oField);
                context.ExecuteQuery();
                Console.WriteLine(x.Name  + " SchemaXml has been updated with WebId");
                Console.WriteLine(x.Name  + " = " + oField.SchemaXml);
                Console.WriteLine("************************************");
                Console.WriteLine();
            }
        }

        private static List<SiteColumnProperty> GetLookUpSiteColumnInfo(ClientContext context, Web oWeb)
        {
            context.Load(oWeb, w => w.Id);

            List LegalEntityList = context.Web.Lists.GetByTitle("TEP_LegalEntity");
            context.Load(LegalEntityList, le => le.Id);

            List businessUnitList = context.Web.Lists.GetByTitle("TEP_BusinessUnit");
            context.Load(businessUnitList, s => s.Id);

            List divisionList = context.Web.Lists.GetByTitle("TEP_Division");
            context.Load(divisionList, d => d.Id);

            List taxJurisdictionList = context.Web.Lists.GetByTitle("TEP_TaxJurisdiction");
            context.Load(taxJurisdictionList, tj => tj.Id);

            List keyProcessList = context.Web.Lists.GetByTitle("TEP_KeyProcess");
            context.Load(keyProcessList, kp => kp.Id);

            List processList = context.Web.Lists.GetByTitle("TEP_Process");
            context.Load(processList, p => p.Id);

            List regionList = context.Web.Lists.GetByTitle("TEP_Region");
            context.Load(regionList, r => r.Id);

            List statusTrackerList = context.Web.Lists.GetByTitle("TEP_StatusTracker");
            context.Load(statusTrackerList, st => st.Id);

            List subProcessList = context.Web.Lists.GetByTitle("TEP_SubProcess");
            context.Load(subProcessList, sp => sp.Id);

            List formList = context.Web.Lists.GetByTitle("TEP_Form");
            context.Load(formList, f => f.Id);

            List taxTtpeList = context.Web.Lists.GetByTitle("TEP_Taxtype");
            context.Load(taxTtpeList, tt => tt.Id);

            List geographyLevel1List = context.Web.Lists.GetByTitle("TEP_GeographyLevel1");
            context.Load(geographyLevel1List, gl1 => gl1.Id);

            List geographyLevel2List = context.Web.Lists.GetByTitle("TEP_GeographyLevel2");
            context.Load(geographyLevel2List, gl2 => gl2.Id);

            List geographyLevel3List = context.Web.Lists.GetByTitle("TEP_GeographyLevel3");
            context.Load(geographyLevel3List, gl3 => gl3.Id);

            List geographyLevel4List = context.Web.Lists.GetByTitle("TEP_GeographyLevel4");
            context.Load(geographyLevel4List, gl4 => gl4.Id);

            List issueTrackerList = context.Web.Lists.GetByTitle("TEP_IssueTracker");
            context.Load(issueTrackerList, it => it.Id);

            List currencyList = context.Web.Lists.GetByTitle("TEP_Currency");
            context.Load(currencyList, c => c.Id);

            List statusTrackerCheckListQuestionList = context.Web.Lists.GetByTitle("TEP_StatusTrackerChecklistQuestions");
            context.Load(statusTrackerCheckListQuestionList, q => q.Id);

            List lwRegistrationTypeList = context.Web.Lists.GetByTitle("TEP_LERegistrationType");
            context.Load(lwRegistrationTypeList, r => r.Id);

            context.ExecuteQuery();

            var listOfLookUpSiteColumnDetails = new List<SiteColumnProperty>()
            {
                new SiteColumnProperty() { DisplayName = "Division", Name = "TEP_Division", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=divisionList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Filing Parent", Name = "TEP_FilingParent", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, List=LegalEntityList.Id, WebId=oWeb.Id},
                new SiteColumnProperty() { DisplayName = "Jurisdiction", Name = "TEP_Jurisdiction", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=taxJurisdictionList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "KeyProcess", Name = "TEP_KeyProcess", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=keyProcessList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Process", Name = "TEP_Process", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=processList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "TEP_Region", Name = "TEP_Region", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=regionList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Related GeneralTaskID", Name = "TEP_RelatedGeneralTaskID", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="ID", List=statusTrackerList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Sub Process", Name = "TEP_SubProcess", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=subProcessList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Tax Form", Name = "TEP_TaxForm", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=formList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Tax Type", Name = "TEP_TaxType", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=taxTtpeList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Business Unit", Name = "TEP_BusinessUnit", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=businessUnitList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Geography Level2", Name = "TEP_GeographyLevel2", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=geographyLevel2List.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Geography Level3", Name = "TEP_GeographyLevel3", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=geographyLevel3List.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Geography Level4", Name = "TEP_GeographyLevel4", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=geographyLevel4List.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Legal Entity", Name = "TEP_LegalEntity", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=LegalEntityList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Related Issues", Name = "TEP_RelatedIssues", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=issueTrackerList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "RelatedTaskID", Name = "TEP_RelatedTaskID", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="ID", List=statusTrackerList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Country", Name = "TEP_Country", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=taxJurisdictionList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "Currency", Name = "TEP_Currency", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=currencyList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "QuestionID", Name = "TEP_QuestionID", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="ID", List=statusTrackerCheckListQuestionList.Id, WebId=oWeb.Id },
                new SiteColumnProperty() { DisplayName = "LE Registration Type", Name = "TEP_RegistrationType", Format = "", Type = "Lookup", Group = "_TEP_Columns", IsRequired=false, ShowField="Title", List=lwRegistrationTypeList.Id, WebId=oWeb.Id },
            };

            return listOfLookUpSiteColumnDetails;
        }
    }
}
