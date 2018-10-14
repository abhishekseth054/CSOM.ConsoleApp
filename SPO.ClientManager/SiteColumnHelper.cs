using SPO.ClientManager.Model;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SPO.ClientManager
{
    public class SiteColumnHelper
    {
        public static void ValidateAndCreateSiteColumn()
        {
            var clientContext = AuthHelper.GetClientContext();
            Web oWeb = clientContext.Web;

            clientContext.Load(oWeb, w => w.Id);

            List EntityList = clientContext.Web.Lists.GetByTitle("Entity");
            clientContext.Load(EntityList, e => e.Id);



            List JurisdictionList = clientContext.Web.Lists.GetByTitle("Jurisdiction");
            clientContext.Load(JurisdictionList, j => j.Id);

            List keyProcessList = clientContext.Web.Lists.GetByTitle("KeyProcess");
            clientContext.Load(keyProcessList, kp => kp.Id);

            List processList = clientContext.Web.Lists.GetByTitle("Process");
            clientContext.Load(processList, p => p.Id);

            List subProcessList = clientContext.Web.Lists.GetByTitle("SubProcess");
            clientContext.Load(subProcessList, sp => sp.Id);

            List regionList = clientContext.Web.Lists.GetByTitle("Region");
            clientContext.Load(regionList, r => r.Id);

            List taskList = clientContext.Web.Lists.GetByTitle("Task");
            clientContext.Load(taskList, t => t.Id);

            List geographyLevel1List = clientContext.Web.Lists.GetByTitle("GeographyLevel1");
            clientContext.Load(geographyLevel1List, gl1 => gl1.Id);

            List geographyLevel2List = clientContext.Web.Lists.GetByTitle("GeographyLevel2");
            clientContext.Load(geographyLevel2List, gl2 => gl2.Id);

            List issueList = clientContext.Web.Lists.GetByTitle("Issue");
            clientContext.Load(issueList, i => i.Id);

            List currencyList = clientContext.Web.Lists.GetByTitle("Currency");
            clientContext.Load(currencyList, c => c.Id);

            clientContext.ExecuteQuery();


            string frequencyChoices = "<CHOICES>"
                                      + "    <CHOICE>Annually</CHOICE>"
                                      + "    <CHOICE>Biennially</CHOICE>"
                                      + "    <CHOICE>Bi-Annually</CHOICE>"
                                      + "    <CHOICE>Semi-Annually</CHOICE>"
                                      + "    <CHOICE>Monthly</CHOICE>"
                                      + "    <CHOICE>Quarterly</CHOICE>"
                                      + "    <CHOICE>Quarterly</CHOICE>"
                                      + "    <CHOICE>One-Time</CHOICE>"
                                      + "    <CHOICE>N/A</CHOICE>"
                                      + "    <CHOICE>Varies</CHOICE>"
                                      + "    <CHOICE>Fees</CHOICE>"
                                      + "    <CHOICE>Occasional</CHOICE>"
                                      + "    <CHOICE>F2</CHOICE>"
                                      + "</CHOICES>";



            string YearChoices = "<CHOICES>"
                                    + "    <CHOICE>2010</CHOICE>"
                                    + "    <CHOICE>2011</CHOICE>"
                                    + "    <CHOICE>2012</CHOICE>"
                                    + "    <CHOICE>2013</CHOICE>"
                                    + "    <CHOICE>2014</CHOICE>"
                                    + "    <CHOICE>2015</CHOICE>"
                                    + "    <CHOICE>2016</CHOICE>"
                                    + "    <CHOICE>2017</CHOICE>"
                                    + "    <CHOICE>2018</CHOICE>"
                                    + "    <CHOICE>2019</CHOICE>"
                                    + "    <CHOICE>2020</CHOICE>"
                                    + "</CHOICES>";

            string statusesChoices = "<CHOICES>"
                                    + "    <CHOICE>Not Started</CHOICE>"
                                    + "    <CHOICE>In Progress</CHOICE>"
                                    + "    <CHOICE>Completed</CHOICE>"
                                    + "</CHOICES>";

            string finalStatusChoices = "<CHOICES>"
                                        + "    <CHOICE>Not Started</CHOICE>"
                                        + "    <CHOICE>Pending Client Data</CHOICE>"
                                        + "    <CHOICE>Client Data Received</CHOICE>"
                                        + "    <CHOICE>In Preparation</CHOICE>"
                                        + "    <CHOICE>In Review</CHOICE>"
                                        + "    <CHOICE>Pending Approval</CHOICE>"
                                        + "    <CHOICE>Pending EY Approval</CHOICE>"
                                        + "    <CHOICE>Ready For Client</CHOICE>"
                                        + "    <CHOICE>Pending Client Approval</CHOICE>"
                                        + "    <CHOICE>Ready For Filing</CHOICE>"
                                        + "    <CHOICE>Completed and not yet filed</CHOICE>"
                                        + "    <CHOICE>Completed</CHOICE>"
                                        + "</CHOICES>";

            string taskPeriodChoices = "<CHOICES>"
                                        + "    <CHOICE>Q1</CHOICE>"
                                        + "    <CHOICE>Q2</CHOICE>"
                                        + "    <CHOICE>Q3</CHOICE>"
                                        + "    <CHOICE>Q4</CHOICE>"
                                        + "    <CHOICE>Annual</CHOICE>"
                                        + "    <CHOICE>Monthly</CHOICE>"
                                        + "    <CHOICE>January</CHOICE>"
                                        + "    <CHOICE>February</CHOICE>"
                                        + "    <CHOICE>March</CHOICE>"
                                        + "    <CHOICE>April</CHOICE>"
                                        + "    <CHOICE>May</CHOICE>"
                                        + "    <CHOICE>June</CHOICE>"
                                        + "    <CHOICE>July</CHOICE>"
                                        + "    <CHOICE>August</CHOICE>"
                                        + "    <CHOICE>September</CHOICE>"
                                        + "    <CHOICE>October</CHOICE>"
                                        + "    <CHOICE>November</CHOICE>"
                                        + "    <CHOICE>December</CHOICE>"
                                        + "    <CHOICE>Biennially</CHOICE>"
                                        + "    <CHOICE>Semi-Annual 1</CHOICE>"
                                        + "    <CHOICE>Semi-Annual 2</CHOICE>"
                                        + "    <CHOICE>Bi-Monthly 1</CHOICE>"
                                        + "    <CHOICE>Bi-Monthly 2</CHOICE>"
                                        + "</CHOICES>";

            string issueStatusChoices = "<CHOICES>"
                                        + "    <CHOICE>Active</CHOICE>"
                                        + "    <CHOICE>Resolved</CHOICE>"
                                        + "    <CHOICE>Closed</CHOICE>"
                                        + "</CHOICES>";

            string priorityChoices = "<CHOICES>"
                                        + "    <CHOICE>Critical</CHOICE>"
                                        + "    <CHOICE>High</CHOICE>"
                                        + "    <CHOICE>Medium</CHOICE>"
                                        + "    <CHOICE>Low</CHOICE>"
                                        + "</CHOICES>";

            string jurisdictionTypeChoices = "<CHOICES>"
                                            + "    <CHOICE>City</CHOICE>"
                                            + "    <CHOICE>Country</CHOICE>"
                                            + "    <CHOICE>State\\Province</CHOICE>"
                                            + "    <CHOICE>Municipality</CHOICE>"
                                            + "</CHOICES>";

            string categoryChoices = "<CHOICES>"
                                    + "<CHOICE>Category 1</CHOICE>"
                                    + "<CHOICE>Category 2</CHOICE>"
                                    + "<CHOICE>Category 3</CHOICE>"
                                    + "</CHOICES>";

            var listOfSiteColumnProperty = new List<SiteColumnProperty>
        {
            //Custom_Columns Group Site Column  
                
            new SiteColumnProperty() { DisplayName = "Task Name", Name = "SC_TaskName", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Entity", Name = "SC_Entity", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=EntityList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "KeyProcess", Name = "SC_KeyProcess", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=keyProcessList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "Process", Name = "SC_Process", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=processList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "Sub Process", Name = "SC_SubProcess", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=subProcessList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "Jurisdiction", Name = "SC_Jurisdiction", Format = "", Type = "Boolean", Group = "Custom_Columns", IsRequired=false, DefaultValue="0"},
            new SiteColumnProperty() { DisplayName = "Year", Name = "SC_Year", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= YearChoices },
            new SiteColumnProperty() { DisplayName = "Geography Level1", Name = "SC_GeographyLevel1", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=geographyLevel1List.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "Geography Level2", Name = "SC_GeographyLevel2", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=geographyLevel2List.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "Region", Name = "SC_Region", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=regionList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "DueDate", Name = "SC_DueDate", Format = "DateOnly", Type = "DateTime", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Period", Name = "SC_Period", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= taskPeriodChoices },
            new SiteColumnProperty() { DisplayName = "Task Number", Name = "SC_TaskNumber", Format = "", Type = "Number", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Date Started", Name = "SC_DateStarted", Format = "DateOnly", Type = "DateTime", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Extension Date", Name = "SC_ExtensionDate", Format = "DateOnly", Type = "DateTime", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Final Status", Name = "SC_FinalStatus", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= finalStatusChoices },
            new SiteColumnProperty() { DisplayName = "Assigned To", Name = "SC_AssignedTo", Format = "", Type = "User", Group = "Custom_Columns", IsRequired=false, UserSelectionMode="PeopleAndGroups"},
            new SiteColumnProperty() { DisplayName = "Approver", Name = "SC_Approver", Format = "", Type = "User", Group = "Custom_Columns", IsRequired=false, UserSelectionMode="PeopleAndGroups"},
            new SiteColumnProperty() { DisplayName = "Approver Status", Name = "SC_ApproverStatus", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= statusesChoices },
            new SiteColumnProperty() { DisplayName = "Frequecny", Name = "SC_Frequency", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= frequencyChoices},
            new SiteColumnProperty() { DisplayName = "Issue Status", Name = "SC_IssueStatus", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= issueStatusChoices },
            new SiteColumnProperty() { DisplayName = "Priority", Name = "SC_Priority", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= priorityChoices },
            new SiteColumnProperty() { DisplayName = "Description", Name = "SC_Descriptions", Format = "", Type = "Note", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Comments", Name = "SC_Comments", Format = "", Type = "Note", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Category", Name = "SC_Category", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= categoryChoices },
            new SiteColumnProperty() { DisplayName = "RelatedTaskID", Name = "SC_RelatedTaskID", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="ID", List=taskList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "Contact Name", Name = "SC_ContactName", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired = true },
            new SiteColumnProperty() { DisplayName = "Contact Role", Name = "SC_ContactRole", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired = true  },
            new SiteColumnProperty() { DisplayName = "Contact Email", Name = "SC_ContactEmail", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired = true  },
            new SiteColumnProperty() { DisplayName = "Contact Phone", Name = "SC_ContactPhone", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired = true  },
            new SiteColumnProperty() { DisplayName = "Contact Photo", Name = "SC_ContactPhoto", Format = "Hyperlink", Type = "URL", Group = "Custom_Columns", IsRequired = true  },
            new SiteColumnProperty() { DisplayName = "File Name", Name = "SC_FileName", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired=false },
            new SiteColumnProperty() { DisplayName = "Method Name", Name = "SC_MethodName", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired=false },
            new SiteColumnProperty() { DisplayName = "Exception", Name = "SC_Exception", Format = "", Type = "Note", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "City", Name = "SC_City", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired=false },
            new SiteColumnProperty() { DisplayName = "State/Province", Name = "SC_State/Province", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired=false },
            new SiteColumnProperty() { DisplayName = "Country", Name = "SC_Country", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=JurisdictionList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "JurisdictionType", Name = "SC_JurisdictionType", Format = "Dropdown", Type = "Choice", Group = "Custom_Columns", IsRequired=false, ChoicesDetails= jurisdictionTypeChoices },
            new SiteColumnProperty() { DisplayName = "SubmittedBy", Name = "SC_SubmittedBy", Format = "", Type = "User", Group = "Custom_Columns", IsRequired=false, UserSelectionMode="PeopleOnly"},
            new SiteColumnProperty() { DisplayName = "Sort Order Number", Name = "SC_SortOrderNumber", Format = "", Type = "Number", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Start Date", Name = "SC_StartDate", Format = "DateOnly", Type = "DateTime", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "End Date", Name = "SC_EndDate", Format = "DateOnly", Type = "DateTime", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Logo URL", Name = "SC_LogoURL", Format = "Hyperlink", Type = "URL", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Country Code", Name = "SC_CountryCode", Format = "", Type = "Text", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Currency", Name = "SC_Currency", Format = "", Type = "Lookup", Group = "Custom_Columns", IsRequired=false, ShowField="Title", List=currencyList.Id, WebId=oWeb.Id },
            new SiteColumnProperty() { DisplayName = "Payment Date", Name = "SC_PaymentDate", Format = "DateOnly", Type = "DateTime", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Payment Amount", Name = "SC_PaymentAmount", Format = "", Type = "Number", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "DownloadDate", Name = "SC_DownloadDate", Format = "DateOnly", Type = "DateTime", Group = "Custom_Columns", IsRequired=false},
            new SiteColumnProperty() { DisplayName = "Downloaded By", Name = "SC_DownloadedBy", Format = "", Type = "User", Group = "Custom_Columns", IsRequired=false, UserSelectionMode="PeopleOnly"},
        };

            ValidateAndCreate(clientContext, oWeb, listOfSiteColumnProperty);

        }

        private static void ValidateAndCreate(ClientContext clientContext, Web oWeb, List<SiteColumnProperty> listOfSiteColumnProperty)
        {
            foreach (SiteColumnProperty siteColumnProp in listOfSiteColumnProperty)
            {
                int count = Helper.IsExist_Helper(clientContext, siteColumnProp.Name, "field");

                if (siteColumnProp.Type == "Lookup")
                {
                    if (count == 0)
                    {
                        CreateSiteColumnForLookUPType(clientContext, oWeb, siteColumnProp);
                    }
                    else
                    {
                        DeleteSiteColumn(clientContext, oWeb, siteColumnProp.Name, new List<SiteColumnProperty>());
                        CreateSiteColumnForLookUPType(clientContext, oWeb, siteColumnProp);
                    }
                }
                else if (siteColumnProp.Type == "User")
                {
                    if (count == 0)
                    {
                        CreateSiteColumnForUserType(clientContext, oWeb, siteColumnProp);
                    }
                    else
                    {
                        DeleteSiteColumn(clientContext, oWeb, siteColumnProp.Name, new List<SiteColumnProperty>());
                        CreateSiteColumnForUserType(clientContext, oWeb, siteColumnProp);
                    }
                }
                else
                {
                    if (count == 0)
                    {
                        CreateSiteColumn(clientContext, oWeb, siteColumnProp);
                    }
                    else
                    {
                        DeleteSiteColumn(clientContext, oWeb, siteColumnProp.Name, new List<SiteColumnProperty>());
                        CreateSiteColumn(clientContext, oWeb, siteColumnProp);
                    }
                }
            }
        }

        //Create LookUp Column
        private static void CreateSiteColumnForLookUPType(ClientContext clientContext, Web oWeb, SiteColumnProperty siteColumnProp)
        {
            // //define the relationship with the lookup field, in that case the field needs to be indexed:
            // string schemaLookupField = "<Field Type='Lookup' Name='Country' StaticName='Country' DisplayName='Country Name' List = '{B5E2D800F-E739-401F-983F-B40984B70273}' ShowField = 'Title' RelationshipDeleteBehavior = 'Restrict' Indexed = 'TRUE' /> ";
            //string schemaLookupField = "<Field Type='Lookup' Name='Country' StaticName='Country' DisplayName='Country Name' List = 'Countries' ShowField = 'Title' RelationshipDeleteBehavior = 'Restrict' Indexed = 'TRUE' /> ";
            string schemaLookupField = "<Field Type='Lookup' DisplayName='" + siteColumnProp.DisplayName.Trim() + "' WebId ='" + siteColumnProp.WebId + "' Name ='" + siteColumnProp.Name + "' StaticName='" + siteColumnProp.Name + "' List = '" + siteColumnProp.List + "' ShowField = '" + siteColumnProp.ShowField + "' Group='" + siteColumnProp.Group + "' /> ";
            Field lookupField = oWeb.Fields.AddFieldAsXml(schemaLookupField, true, AddFieldOptions.AddFieldInternalNameHint);
            clientContext.Load(lookupField);
            clientContext.ExecuteQuery();
            Console.WriteLine(siteColumnProp.Name + ": Created");

        }

        // Create User Type Site Column
        private static void CreateSiteColumnForUserType(ClientContext clientContext, Web oWeb, SiteColumnProperty siteColumnProp)
        {
            string schemaUserField = "<Field Type='User' Name='" + siteColumnProp.Name + "' StaticName='" + siteColumnProp.Name + "' DisplayName='" + siteColumnProp.Name + "' UserSelectionMode='" + siteColumnProp.UserSelectionMode + "' Group='" + siteColumnProp.Group + "'/>";
            Field userField = oWeb.Fields.AddFieldAsXml(schemaUserField, true, AddFieldOptions.AddFieldInternalNameHint);

            clientContext.Load(userField);
            clientContext.ExecuteQuery();
            Console.WriteLine(siteColumnProp.Name + ": Created");

        }

        public static void CreateSiteColumn(ClientContext clientContext, Web oWeb, SiteColumnProperty listdetauls)
        {
            string schema = "<Field DisplayName='" + listdetauls.DisplayName.Trim() +
                                "' Name='" + listdetauls.Name.Trim() +
                                "' Group='" + listdetauls.Group +
                                "' Type='" + listdetauls.Type +
                                "' Format='" + listdetauls.Format +
                                "' List = '" + listdetauls.List +
                                "' ShowField = '" + listdetauls.ShowField +
                                "' WebId = '" + listdetauls.WebId +
            // NumLines = '" + listdetauls.NumLines + 
            // RichText='" + listdetauls.RichText +
            // RichTextMode = '" + listdetauls.RichTextMode + 
            // IsolateStyles='" + listdetauls.IsolateStyles +
            // Sortable='" + listdetauls.Sortable +
            "'><Default>" + listdetauls.DefaultValue + "</Default>" + listdetauls.ChoicesDetails + "</Field>";

            Field field = oWeb.Fields.AddFieldAsXml(schema, true, AddFieldOptions.AddFieldInternalNameHint);

            field.Required = listdetauls.IsRequired;
            field.Update();
            clientContext.Load(field);
            clientContext.ExecuteQuery();

            Console.WriteLine(listdetauls.Name + ": Created");
        }

        private static void DeleteSiteColumn(ClientContext clientContext, Web oWeb, string siteColumn, List<SiteColumnProperty> listOfSiteColumnProperty)
        {
            if (!String.IsNullOrEmpty(siteColumn))
            {
                oWeb.Fields.GetByInternalNameOrTitle(siteColumn).DeleteObject();
                Console.WriteLine(siteColumn + ": Deleted");
            }

            clientContext.ExecuteQuery();
        }

        public static Guid GetSiteColumnIDByName(ClientContext clientContext, Web oWeb, string SiteColumnName)
        {
            var siteColumn = clientContext.LoadQuery(oWeb.Fields.Where(sc => sc.InternalName == SiteColumnName));
            clientContext.ExecuteQuery();
            var ageSiteColumn = siteColumn.FirstOrDefault();
            var ageSiteColumnId = ageSiteColumn.Id.ToString();

            return new Guid(ageSiteColumnId);
        }
    }
}
