using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPO.ClientManager.Model
{
    public class Data
    {

        public Dictionary<string, Dictionary<string, List<string>>> GetListInfo()
        {
            //this dictionory object is use to mentain the List/Library Name, Content Type and Site Column details

            Dictionary<string, Dictionary<string, List<string>>> listDictObj = new Dictionary<string, Dictionary<string, List<string>>>
            {
                {
                    "Task", new Dictionary<string, List<string>>{{ "CT_Task", new List<string>
                    {
                        "SC_TaskName","SC_Entity","SC_KeyProcess","SC_Process",
                        "SC_SubProcess","SC_Jurisdiction","SC_Year",
                        "SC_GeographyLevel1","SC_Region",
                        "SC_DueDate","SC_Period","SC_TaskNumber","SC_DateStarted",
                        "SC_ExtensionDate","SC_FinalStatus",
                        "SC_AssignedTo","SC_Approver","SC_ApproverStatus","SC_Frequency"
                    } } }
                },
                {
                    "Issue", new Dictionary<string, List<string>>{{ "CT_Issue", new List<string>
                    {
                        "SC_AssignedTo","SC_IssueStatus","SC_Priority","SC_Descriptions",
                        "SC_DueDate","SC_Comments","SC_Category","SC_RelatedTaskID"
                    } }}
                },
                {
                    "Contacts", new Dictionary<string, List<string>>{{ "CT_Contacts", new List<string>
                    {
                            "SC_ContactName","SC_ContactRole","SC_ContactEmail","SC_ContactPhone","SC_ContactPhoto",
                    } }}
                },
                {
                    "Exception",new Dictionary<string, List<string>>{{ "CT_Exception", new List<string>
                    {
                           "SC_FileName","SC_MethodName","SC_Exception",
                    } }}
                },
                {
                    "Jurisdiction",new Dictionary<string, List<string>>{{ "CT_Jurisdiction", new List<string>
                    {
                        "SC_Region", "SC_City", "SC_State/Province","SC_Country","SC_JurisdictionType",
                        "SC_GeographyLevel1","SC_GeographyLevel2"
                    } }}
                },


                {
                    "Entity",new Dictionary<string, List<string>>{{ "CT_Entity", new List<string>{ } }}
                },
                {
                    "HelpRequests",new Dictionary<string, List<string>>{{ "CT_HelpRequests", new List<string>
                    {
                        "SC_Descriptions","SC_SubmittedBy"
                    } }}
                },
                {
                    "Suggestions",new Dictionary<string, List<string>>{{ "CT_Suggestions", new List<string>
                    {
                        "SC_Descriptions","SC_SubmittedBy"
                    } }}
                },
                {
                    "Announcements",new Dictionary<string, List<string>>{{ "CT_Announcements", new List<string>
                    {
                            "SC_Descriptions","SC_SortOrderNumber","SC_LogoURL","SC_StartDate","SC_EndDate",
                    } }}
                },
                {
                    "KeyProcess",new Dictionary<string, List<string>>{{ "CT_KeyProcess", new List<string>{}}}
                },
                {
                    "Process",new Dictionary<string, List<string>>{{ "CT_Process", new List<string>
                    {
                        "SC_KeyProcess"
                    } }}
                },
                {
                    "SubProcess",new Dictionary<string, List<string>>{{ "CT_SubProcess", new List<string>
                    {
                        "SC_Process"
                    } }}
                },
                {
                    "Region",new Dictionary<string, List<string>>{{ "CT_Region", new List<string>{}}}
                },
                {
                    "GeographyLevel1",new Dictionary<string, List<string>>{{ "CT_GeographyLevel1", new List<string>
                    {
                        "SC_Region", "SC_CountryCode"
                    } }}
                },
                {
                    "GeographyLevel2",new Dictionary<string, List<string>>{{ "CT_GeographyLevel2", new List<string>
                    {
                        "SC_GeographyLevel1"
                    } }}
                },
                {
                    "Payments",new Dictionary<string, List<string>>{{ "CT_Payments", new List<string>
                    {
                        "SC_RelatedTaskID","SC_Currency","SC_PaymentDate","SC_PaymentAmount"
                    } }}
                },
                {
                    "Currency",new Dictionary<string, List<string>>{{ "CT_Currency", new List<string>{}}}
                }
            };

            return listDictObj;
            //return new Dictionary<string, Dictionary<string, List<string>>>();
        }

        public Dictionary<string, Dictionary<string, List<string>>> GetLibInfo()
        {
            Dictionary<string, Dictionary<string, List<string>>> libDictObj = new Dictionary<string, Dictionary<string, List<string>>>
            {
                {
                    "Documents", new Dictionary<string, List<string>>{{ "CT_Documents", new List<string>
                    {
                        "SC_Descriptions","SC_DownloadDate","SC_DownloadedBy","SC_Jurisdiction","SC_KeyProcess",
                        "SC_Entity","SC_Process","SC_SubProcess","SC_Year"
                    } }}
                },
                {
                    "ContactPhotos", new Dictionary<string, List<string>>{{ "CT_ContactPhotos", new List<string>{}}}
                },
                {
                    "FinalDocuments", new Dictionary<string, List<string>>{{ "CT_FinalDocuments", new List<string>
                    {
                        "SC_Descriptions","SC_Jurisdiction","SC_KeyProcess","SC_Process","SC_SubProcess",
                        "SC_Year","SC_Entity"

                    } }}
                },
            };

            return libDictObj;
        }
    }
}