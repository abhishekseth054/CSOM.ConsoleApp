﻿using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using SPO.ClientManager.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SPO.ClientManager
{
    public class ListHelper
    {
        #region List and Library Creation and Deletion

        public static void ValidateAndCreateListAndLibrry()
        {
            // ClientContext - Get the context for the SharePoint Site    
            var clientContext = AuthHelper.GetClientContext();
            clientContext.ExecuteQuery();

            var listDetails = new Data().GetListInfo().Keys;
            var libraryDetails = new Data().GetLibInfo().Keys;


            // Create List
            foreach (string listName in listDetails)
            {
                if (!String.IsNullOrEmpty(listName.Trim()))
                {
                    string templateType = "List";
                    int count = Helper.IsExist_Helper(clientContext, listName, "list");
                    if (count == 0)
                    {
                        CreateListLibrary(clientContext, listName, templateType);
                    }
                    else
                    {
                        DeleteListAndLibrary(clientContext, listName);
                        CreateListLibrary(clientContext, listName, templateType);
                    }
                }
            }

            //Create Library
            foreach (string libName in libraryDetails)
            {
                if (!String.IsNullOrEmpty(libName.Trim()))
                {
                    string templateType = "Library";

                    int count = Helper.IsExist_Helper(clientContext, libName, "list");

                    if (count == 0)
                    {
                        CreateListLibrary(clientContext, libName, templateType);
                    }
                    else
                    {
                        DeleteListAndLibrary(clientContext, libName);
                        CreateListLibrary(clientContext, libName, templateType);
                    }
                }
            }
        }

        public static void CreateListLibrary(ClientContext clientContext, string listName, string templateType)
        {
            Web oWebsite = clientContext.Web;
            ListCreationInformation listCreationInfo = new ListCreationInformation();
            listCreationInfo.Title = listName;
            listCreationInfo.Description = listName + " Created By CSOM";
            if (templateType == "List")
            {
                listCreationInfo.TemplateType = (int)ListTemplateType.GenericList;
            }
            else if (templateType == "Library")
            {
                listCreationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
            }

            List oList = oWebsite.Lists.Add(listCreationInfo);
            clientContext.Load(oList, o => o.EnableVersioning);
            clientContext.ExecuteQuery();
            Console.WriteLine(listName + " Created");

            if (templateType == "List")
            {
                oList.UpdateListVersioning(true, false, true);
                oList.MajorVersionLimit = 5;
            }
            else if (templateType == "Library")
            {
                oList.UpdateListVersioning(true, true, true);
                oList.MajorVersionLimit = 5;
                oList.MajorWithMinorVersionsLimit = 3;
                oList.DraftVersionVisibility = DraftVisibilityType.Author;
            }
            oList.Update();
            clientContext.ExecuteQuery();
            Console.WriteLine(listName + " Versioning enabled ");

        }

        public static void DeleteListAndLibrary(ClientContext clientContext, string listLibName)
        {
            Web oWebsite = clientContext.Web;

            List oList = oWebsite.Lists.GetByTitle(listLibName);
            oList.DeleteObject();
            clientContext.ExecuteQuery();
            clientContext.Web.Update();

            Console.WriteLine(listLibName + " Deleted");
        }

        public static void ValidateAndAssociateContenTypeToList()
        {
            var clientContext = AuthHelper.GetClientContext();
            Web oweb = clientContext.Web;

            var listDetails = new Data().GetListInfo();
            var libraryDetails = new Data().GetLibInfo();

            //Push the data from libraryDetails to listDetails
            libraryDetails.ToList().ForEach(x => listDetails.Add(x.Key, x.Value));

            listDetails.ToDictionary(x => x.Key);

            var contentTypeDetails = listDetails.Values.ToList();
            
            //iterate from the listDetails and associate content Type with the list 
            //value is Content type and key is List name

            for (int i = 0, j = 0; i < listDetails.Keys.Count(); i++, j++)
            {
                var listName = listDetails.Keys.ToList()[i];

                foreach (var contentTypeName in contentTypeDetails[j].Keys)
                {
                    var contentTypeID = ContentTypeHelper.GetContentTypeIdByName(clientContext, contentTypeName);

                    bool isAssociated = oweb.ContentTypeExistsByName(listName, contentTypeName);

                    if (!isAssociated)
                    {
                        AssociateContentTypeToList(clientContext, oweb, listName, contentTypeID, contentTypeName);
                        SetDefaultContentType(clientContext, oweb, listName, contentTypeID, contentTypeName);
                    }
                    else
                    {
                        Console.WriteLine(contentTypeName + " Content Type already Associated to " + listName + " List/Library");
                    }

                    break;
                }
            }
        }

        private static void SetDefaultContentType(ClientContext clientContext, Web oweb, string listName, string contentTypeID, string contentTypeName)
        {
            List aeList = clientContext.Web.Lists.GetByTitle(listName);
            var currentCtOrder = aeList.ContentTypes;
            clientContext.Load(currentCtOrder, coll => coll.Include(
                                    ct => ct.Name,
                                    ct => ct.Id));
            clientContext.ExecuteQuery();

            IList<ContentTypeId> reverseOrder = (from ct in currentCtOrder where ct.Name.Equals(contentTypeName, StringComparison.OrdinalIgnoreCase) select ct.Id).ToList();
            aeList.RootFolder.UniqueContentTypeOrder = reverseOrder;
            aeList.RootFolder.Update();
            aeList.Update();
            clientContext.ExecuteQuery();

            Console.WriteLine(contentTypeName + " is now default contnet Type for the " + listName + " List/Librry");
        }

        private static void AssociateContentTypeToList(ClientContext clientContext, Web oweb, string listName, string contentTypeID, string contentTypeName)
        {
            ContentType ct = oweb.ContentTypes.GetById(contentTypeID);
            clientContext.Load(oweb);
            clientContext.Load(ct);
            clientContext.ExecuteQuery();

            List list = oweb.Lists.GetByTitle(listName);
            list.ContentTypesEnabled = true;
            list.ContentTypes.AddExistingContentType(ct);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            Console.WriteLine(contentTypeName + " Content Type Associated to " + listName + " List/Library");
        }

        #endregion


        #region Not Using

        //public static void CreateListColumn()
        //{
        //    var clientContext = Helper.GetClientContext();

        //    List oList = clientContext.Web.Lists.GetByTitle("Test LIST");

        //    //Number DataType Column
        //    Field numberField = oList.Fields.AddFieldAsXml("<Field DisplayName='Age' Type='Number' />", true, AddFieldOptions.DefaultValue);
        //    FieldNumber fieldNumber = clientContext.CastTo<FieldNumber>(numberField);
        //    fieldNumber.MaximumValue = 100;
        //    fieldNumber.MinimumValue = 35;
        //    fieldNumber.Update();
        //    clientContext.Load(fieldNumber);

        //    // //Single Line Of Text DataType Column
        //    Field textField = oList.Fields.AddFieldAsXml("<Field DisplayName='SingleLine' Type='Text' />", true, AddFieldOptions.DefaultValue);
        //    FieldText fieldText = clientContext.CastTo<FieldText>(textField);
        //    fieldText.Update();
        //    clientContext.Load(fieldText);

        //    //Multi Line Of Text DataType Column
        //    Field multiLineField = oList.Fields.AddFieldAsXml("<Field DisplayName='MultiLine' Type='Note' />", true, AddFieldOptions.DefaultValue);
        //    FieldMultiLineText fieldmultiLineText = clientContext.CastTo<FieldMultiLineText>(multiLineField);
        //    fieldmultiLineText.Update();
        //    clientContext.Load(fieldmultiLineText);

        //    //Multi Line Rich Text DataType Column
        //    Field multiLineRichTextField = oList.Fields.AddFieldAsXml("<Field DisplayName='Multi Line RichText' Type='Note' />", true, AddFieldOptions.DefaultValue);
        //    FieldMultiLineText fieldmultiLineRichText = clientContext.CastTo<FieldMultiLineText>(multiLineRichTextField);
        //    fieldmultiLineRichText.AllowHyperlink = true;
        //    fieldmultiLineRichText.RichText = true;
        //    fieldmultiLineRichText.Update();
        //    fieldmultiLineRichText.UpdateAndPushChanges(true);
        //    clientContext.Load(fieldmultiLineRichText);

        //    //An enhanced multi line text field
        //    string schemaRichTextField = "<Field Type='Note' Name='EnhancedmultiLine' StaticName='EnhancedmultiLine' DisplayName = 'Enhanced multiLine' NumLines = '6' RichText = 'TRUE' RichTextMode = 'FullHtml' IsolateStyles = 'TRUE' Sortable = 'FALSE' /> ";
        //    Field multilineenhancedTextField = oList.Fields.AddFieldAsXml(schemaRichTextField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(multilineenhancedTextField);

        //    // DropDown Choice
        //    string schemaChoiceFieldDDL = "<Field Type='Choice' DisplayName='ChoiceDDL' Name='ChoiceDDL' StaticName='ChoiceDDL' Format = 'Dropdown' >"
        //                               + "<Default>Option 2</Default>"
        //                               + "<CHOICES>"
        //                               + "    <CHOICE>Option 2</CHOICE>"
        //                               + "    <CHOICE>Option 3</CHOICE>"
        //                               + "</CHOICES>"
        //                               + "</Field>";

        //    Field choiceDDLField = oList.Fields.AddFieldAsXml(schemaChoiceFieldDDL, true, AddFieldOptions.AddFieldInternalNameHint);
        //    FieldChoice fieldChoice = clientContext.CastTo<FieldChoice>(choiceDDLField);
        //    fieldChoice.Required = true;
        //    fieldChoice.Update();
        //    clientContext.Load(fieldChoice);

        //    //Radio buttons
        //    string schemaRadioChoiceField = "<Field Type='Choice' Name='ChoiceRadio' StaticName='ChoiceRadio' DisplayName = 'Choice Radio' Format = 'RadioButtons' > "
        //         + "<Default>Opt Radio 3</Default>"
        //         + "<CHOICES>"
        //         + "    <CHOICE>Opt Radio 1</CHOICE>"
        //         + "    <CHOICE>Opt Radio 2</CHOICE>"
        //         + "    <CHOICE>Opt Radio 3</CHOICE>"
        //         + "</CHOICES>"
        //         + "</Field>";
        //    Field choiceField = oList.Fields.AddFieldAsXml(schemaRadioChoiceField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(choiceField);

        //    //Checkboxes
        //    string schemaMultiChoiceField = "<Field Type='MultiChoice' Name='ChoiceMulti' StaticName='ChoiceMulti' DisplayName = 'Choice Multi' > "
        //           + "<Default>MultiChoice 2</Default>"
        //           + "<CHOICES>"
        //           + "    <CHOICE>MultiChoice 1</CHOICE>"
        //           + "    <CHOICE>MultiChoice 2</CHOICE>"
        //           + "    <CHOICE>MultiChoice 3</CHOICE>"
        //           + "</CHOICES>"
        //           + "</Field>";
        //    Field choiceMultiChoiceField = oList.Fields.AddFieldAsXml(schemaMultiChoiceField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(choiceMultiChoiceField);

        //    //Fill In option
        //    string schemaFillInChoiceField = "<Field Type='Choice' DisplayName='Fill In Choice' Name='FillInChoice' StaticName='FillInChoice' Format = 'Dropdown' FillInChoice = 'TRUE' > "
        //             + "<Default>My Choices Data will come here</Default>"
        //             + "<CHOICES>"
        //             + "    <CHOICE>FillInChoice 1</CHOICE>"
        //             + "    <CHOICE>FillInChoice 2</CHOICE>"
        //             + "    <CHOICE>FillInChoice 3</CHOICE>"
        //             + "</CHOICES>"
        //             + "</Field>";
        //    Field choiceFillInChoiceField = oList.Fields.AddFieldAsXml(schemaFillInChoiceField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(choiceFillInChoiceField);

        //    //Picture field

        //    string schemaPictureField = "<Field Type='URL' Name='EmployeePicture' StaticName='EmployeePicture' DisplayName = 'Employee Picture' Format = 'Image' /> ";
        //    Field pictureField = oList.Fields.AddFieldAsXml(schemaPictureField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(pictureField);

        //    //URL field

        //string schemaUrlField = "<Field Type='URL' Name='BlogUrl' StaticName='BlogUrl' DisplayName='Blog URL' Format='Hyperlink'/>";
        //Field urlField = oList.Fields.AddFieldAsXml(schemaUrlField, true, AddFieldOptions.AddFieldInternalNameHint);
        //clientContext.Load(urlField);
        //clientContext.ExecuteQuery();

        //    // Lookup field

        //    List countryList = clientContext.Web.Lists.GetByTitle("Countries");
        //    clientContext.Load(countryList, c => c.Id);
        //    clientContext.ExecuteQuery();


        //    // //define the relationship with the lookup field, in that case the field needs to be indexed:
        //    // string schemaLookupField = "<Field Type='Lookup' Name='Country' StaticName='Country' DisplayName='Country Name' List = '{B5E2D800F-E739-401F-983F-B40984B70273}' ShowField = 'Title' RelationshipDeleteBehavior = 'Restrict' Indexed = 'TRUE' /> ";
        //    //string schemaLookupField = "<Field Type='Lookup' Name='Country' StaticName='Country' DisplayName='Country Name' List = 'Countries' ShowField = 'Title' RelationshipDeleteBehavior = 'Restrict' Indexed = 'TRUE' /> ";
        //    string schemaLookupField = "<Field Type='Lookup' Name='Country' StaticName='Country' DisplayName='Country Name' List = '" + countryList.Id + "' ShowField = 'Title' /> ";
        //    Field lookupField = oList.Fields.AddFieldAsXml(schemaLookupField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    lookupField.Update();
        //    clientContext.Load(lookupField);

        //    //// multi-select lookup field
        //    string schemaMultiLookupField = "<Field Type='LookupMulti' Name='Country' StaticName='Country' DisplayName='Country' List = '" + countryList.Id + "' ShowField = 'Title' Mult = 'TRUE' /> ";
        //    //string schemaMultiLookupField = "<Field Type='LookupMulti' Name='Country' StaticName='Country' DisplayName='Country' List = 'Countries' ShowField = 'Title' Mult = 'TRUE' /> ";
        //    Field lookupFieldmulti = oList.Fields.AddFieldAsXml(schemaMultiLookupField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    lookupFieldmulti.Update();
        //    clientContext.Load(lookupFieldmulti);


        //    ////Ref: https://karinebosch.wordpress.com/my-articles/creating-fields-using-csom/

        //    // //User Field
        //    string schemaUserField = "<Field Type='User' Name='UserName' StaticName='UserName' DisplayName='User Name' />";
        //    Field userField = oList.Fields.AddFieldAsXml(schemaUserField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(userField);

        //    ////User Field multiple
        //    //string schemaUserGroupField = "<Field Type='UserMulti' Name='EmployeeName' StaticName='EmployeeName' DisplayName='Employee Name' UserSelectionMode = 'PeopleOnly' UserSelectionScope = '7' Mult = 'TRUE' /> ";
        //    string schemaUserGroupField = "<Field Type='UserMulti' Name='EmployeeName' StaticName='EmployeeName' DisplayName='Employee Name' UserSelectionMode = 'PeopleAndGroups' UserSelectionScope = '7' Mult = 'TRUE' /> ";
        //    Field userGroupField = oList.Fields.AddFieldAsXml(schemaUserGroupField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(userGroupField);

        //    ////boolean field

        //    string schemaBooleanField = "<Field Type='Boolean' Name='Married' StaticName='Married' DisplayName='Married'> <Default>1</Default> </Field>";
        //    Field booleanField = oList.Fields.AddFieldAsXml(schemaBooleanField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(booleanField);

        //    ////DateTime Field

        //    //Date only field
        //    string schemaBirthDate = "<Field Type='DateTime' Name='BirthDate' StaticName='BirthDate' DisplayName = 'Birth date' Format = 'DateOnly'> <Default>[Today]</Default></Field>";
        //    Field birthDateField = oList.Fields.AddFieldAsXml(schemaBirthDate, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(birthDateField);

        //    ////Date and time field
        //    string schemaArrivalField = "<Field Type='DateTime' Name='ArrivalDateTime' StaticName='ArrivalDateTime' DisplayName = 'Arrival' Format = 'DateTime'> <Default>[Now]</Default></Field>";
        //    Field DateTimeField = oList.Fields.AddFieldAsXml(schemaArrivalField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(DateTimeField);

        //    ////hidden field

        //    string schemaHiddenTextField = "<Field Type='Text' Name='HiddenField' StaticName='HiddenField' DisplayName='Hidden Field' Hidden='TRUE' />";
        //    Field hiddenTextField = oList.Fields.AddFieldAsXml(schemaHiddenTextField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(hiddenTextField);

        //    //indexed field // Not Working as of now

        //    //Field f = oList.Fields.GetByInternalNameOrTitle("ID");
        //    //clientContext.Load(f);
        //    //clientContext.ExecuteQuery();
        //    //f.Indexed = true;
        //    //f.Update();

        //    //Managed Metadata field

        //    Guid termStoreId = Guid.Empty;
        //    Guid termSetId = Guid.Empty;
        //    GetTaxonomyFieldInfo(clientContext, out termStoreId, out termSetId);

        //    // Single selection Taxonomy field
        //    string schemaTaxonomyField = "<Field Type='TaxonomyFieldType' Name='TaxonomyField' StaticName='TaxonomyField' DisplayName = 'Taxonomy Field' /> ";
        //    Field taxonomyFieldSingle = oList.Fields.AddFieldAsXml(schemaTaxonomyField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(taxonomyFieldSingle);

        //    // Retrieve the field as a Taxonomy Field
        //    TaxonomyField taxonomyField = clientContext.CastTo<TaxonomyField>(taxonomyFieldSingle);
        //    taxonomyField.SspId = termStoreId;
        //    taxonomyField.TermSetId = termSetId;
        //    taxonomyField.TargetTemplate = String.Empty;
        //    taxonomyField.AnchorId = Guid.Empty;
        //    taxonomyField.Update();

        //    // Multi selection Taxonomy field

        //    string schemaTaxonomyFieldMulti = "<Field Type='TaxonomyFieldTypeMulti' Name='TaxonomyFieldMulti' StaticName='TaxonomyFieldMulti' DisplayName = 'Taxonomy Field Multi' Mult = 'TRUE' /> ";
        //    Field taxonomyFieldMulti = oList.Fields.AddFieldAsXml(schemaTaxonomyFieldMulti, false, AddFieldOptions.AddFieldInternalNameHint);
        //    clientContext.Load(taxonomyFieldMulti);

        //    // Retrieve the field as a Taxonomy Field
        //    TaxonomyField taxonomyField1 = clientContext.CastTo<TaxonomyField>(taxonomyFieldMulti);
        //    taxonomyField1.SspId = termStoreId;
        //    taxonomyField1.TermSetId = termSetId;
        //    taxonomyField1.TargetTemplate = String.Empty;
        //    taxonomyField1.AnchorId = Guid.Empty;
        //    taxonomyField1.Update();


        //    clientContext.ExecuteQuery();

        //    //Calculated field

        //    // Not Working

        //    //string formula = "<Formula>=Age&amp;\"\"&amp;SingleLine&amp;\"(id:\"&amp;ID&amp;\"\"</Formula>"
        //    //                  + "<FieldRefs>"
        //    //                  + "<FieldRef Name='Age' />"
        //    //                  + "<FieldRef Name='SingleLine' />"
        //    //                  + "<FieldRef Name='ID' />"
        //    //                  + "</FieldRefs>";

        //    //string schemaCalculatedField = "<Field Type='Calculated' Name='CalculatedField' StaticName='CalculatedField' DisplayName = 'Calculated Field' ResultType = 'Text' Required = 'TRUE' ReadOnly = 'TRUE' > " + formula + " </ Field > ";
        //    //Field fullNameField = oList.Fields.AddFieldAsXml(schemaCalculatedField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    //clientContext.ExecuteQuery();

        //    string fieldXml = "<Field Name='CalculatedField_Year' StaticName='CalculatedField_Year' DisplayName='CalculatedField Year' Type='Text' ReadOnly = 'TRUE'>"
        //                       + "<DefaultFormula>=CONCATENATE(YEAR(Today))</DefaultFormula>"
        //                       + "</Field>";
        //    Field field = oList.Fields.AddFieldAsXml(fieldXml, true, AddFieldOptions.DefaultValue);
        //    clientContext.ExecuteQuery();
        //}

        //private static void GetTaxonomyFieldInfo(ClientContext clientContext, out Guid termStoreId, out Guid termSetId)
        //{
        //    termStoreId = Guid.Empty;
        //    termSetId = Guid.Empty;

        //    TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
        //    TermStore termStore = session.GetDefaultSiteCollectionTermStore();
        //    TermSetCollection termSets = termStore.GetTermSetsByName("Header", 1033);

        //    clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
        //    clientContext.Load(termStore, ts => ts.Id);
        //    clientContext.ExecuteQuery();

        //    termStoreId = termStore.Id;
        //    termSetId = termSets.FirstOrDefault().Id;
        //}

        //public static void DeleteListColumn()
        //{
        //    var clientContext = Helper.GetClientContext();

        //    List oList = clientContext.Web.Lists.GetByTitle(listName);
        //    Field f = oList.Fields.GetByInternalNameOrTitle("SingleLine");
        //    f.DeleteObject();
        //    clientContext.Load(f);
        //    clientContext.ExecuteQuery();
        //}

        #endregion
    }
}
