using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
using System.Text;

using Microsoft.SharePoint.Client.Utilities;

namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    await LoadUserFromEmailOrName(ctx,"le nguyen");

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        #region Exercise 1

        // 1 Using CSOM create a List name “CSOM Test"
        static async Task CreateList(ClientContext ctx)
        {
            Web web = ctx.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = "CSOM Test";
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;
            List list = web.Lists.Add(creationInfo);
            list.Description = "CSOM Test - List Description";

            list.Update();
            await ctx.ExecuteQueryAsync();
        }

        // 2 Create term set “city-{yourname}” in dev tenant
        static async Task CreateTermSet(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get TermStore
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get TermGroup
            TermGroup termGroup = termStore.Groups.GetByName("Site Collection - sharepointassignment.sharepoint.com-sites-ngocleprecio");
            // Int variable - new term lcid
            int lcid = 1033;
            // Guid - new term guid
            Guid newTermId = Guid.NewGuid();
            // Create TermSet
            TermSet newTermSet = termGroup.CreateTermSet("city-ngocle", newTermId, lcid);

            await ctx.ExecuteQueryAsync();

        }

        // 3 Create 2 terms “Ho Chi Minh” and “Stockholm” in termset “city-{yourname}”
        static async Task CreateTerm(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get TermStore
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get TermGroup
            TermGroup termGroup = termStore.Groups.GetByName("Site Collection - sharepointassignment.sharepoint.com-sites-ngocleprecio");
            // Get TermSet
            TermSet termSet = termGroup.TermSets.GetByName("city-ngocle");
            // Int variable - new term lcid
            int lcid = 1033;
            // Guid - new term guid
            Guid newTermId = Guid.NewGuid();
            // Create Termg
            Term newTerm = termSet.CreateTerm("Ho Chi Minh", lcid, newTermId);

            newTermId = Guid.NewGuid();
            // Create Term
            newTerm = termSet.CreateTerm("Stockholm", lcid, newTermId);

            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success");

        }

        // 4 Create site fields “about” type text and field “city” type taxonomy
        static async Task CreateAboutSiteFields(ClientContext ctx)
        {
            Web web = ctx.Web;
            Guid guid = Guid.NewGuid();
            string fieldAsXml = $"<Field ID='{guid}' Name='About' DisplayName='About' Type='Text' Hidden='False' Group='Custom Columns' Description='About Text Field' />";
            web.Fields.AddFieldAsXml(fieldAsXml, true, AddFieldOptions.DefaultValue);
            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success");
        }

        static async Task CreateCityTaxonomySiteFields(ClientContext ctx)
        {
            Web web = ctx.Web;
            Guid guid = Guid.NewGuid();
            string fieldAsXml = $"<Field ID='{guid}' Name='City' DisplayName='City' Type='TaxonomyFieldType' Hidden='False' Group='Custom Columns' Description='City Taxonomy Field' />";
            Field field = web.Fields.AddFieldAsXml(fieldAsXml, true, AddFieldOptions.DefaultValue);
            await ctx.ExecuteQueryAsync();

            (Guid, Guid) result = await GetTaxonomyFieldInfo(ctx);
            Guid termStoreId = result.Item1;
            Guid termSetId = result.Item2;

            // Cast as Taxonomy Field
            TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();

            await ctx.ExecuteQueryAsync();
        }

        static async Task<(Guid, Guid)> GetTaxonomyFieldInfo(ClientContext clientContext)
        {
            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;

            TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            TermSetCollection termSets = termStore.GetTermSetsByName("city-ngocle", 1033);

            clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            clientContext.Load(termStore, ts => ts.Id);
            await clientContext.ExecuteQueryAsync();
            termStoreId = termStore.Id;
            termSetId = termSets.FirstOrDefault().Id;
            return (termStoreId, termSetId);
        }

        // 5 Create site content type “CSOM Test content type” then add this content type to list “CSOM test”,
        // add fields “about” and “city” to this content type
        static async Task CreateContentType(ClientContext ctx)
        {
            // Create Content Type
            ContentTypeCollection contentTypeColl = ctx.Web.ContentTypes;
            string ctName = "CSOM Test content type";
            ContentTypeCreationInformation ctCreation = new ContentTypeCreationInformation();
            ctCreation.Name = ctName;
            ctCreation.Description = "Custom Content Type created using CSOM";
            ctCreation.Group = "List Content Types";
            ContentType ct = contentTypeColl.Add(ctCreation);
            ctx.Load(ct);

            // Add to list
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            list.ContentTypes.AddExistingContentType(ct);
            await ctx.ExecuteQueryAsync();
        }

        static async Task AddFieldsToContentType(ClientContext ctx)
        {
            string cityGuid = await GetCityFieldGuid(ctx);
            Guid id = new Guid(cityGuid); // City Field GUID becasue there is another site column with similar name in sites
            Field fieldCityAdded = ctx.Web.Fields.GetById(id);
            Field fieldAboutAdded = ctx.Web.Fields.GetByTitle("About");
            ContentType ct = ctx.Web.ContentTypes.GetById("0x0100954DE3873C92F840AD95A8AEF6E420CF");

            ct.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldCityAdded,
            });
            ct.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldAboutAdded
            });
            ct.Update(true);
            await ctx.ExecuteQueryAsync();
        }

        // 6 In list “CSOM test” set “CSOM Test content type” as default content type
        static async Task SetDefaultContentTypeToList(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            ContentTypeCollection currentCtOrder = list.ContentTypes;
            ctx.Load(currentCtOrder);
            await ctx.ExecuteQueryAsync();
            IList<ContentTypeId> reverseOrder = new List<ContentTypeId>();
            foreach (ContentType ct in currentCtOrder)
            {
                if (ct.Name.Equals("CSOM Test content type"))
                {
                    reverseOrder.Add(ct.Id);
                }
            }
            list.RootFolder.UniqueContentTypeOrder = reverseOrder;
            list.RootFolder.Update();
            list.Update();
            await ctx.ExecuteQueryAsync();
        }

        // 7 Create 5 list items to list with some value in field “about” and “city”
        static async Task CreateListItem(ClientContext ctx, int numberOfItems, bool isSetDefaultAbout, bool isSetDefaultCity)
        {
            TermCollection terms = await GetTerms("city-ngocle", ctx);

            List oList = ctx.Web.Lists.GetByTitle("CSOM Test");

            for (int i = 0; i < numberOfItems; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                ListItem oListItem = oList.AddItem(itemCreateInfo);

                oListItem["Title"] = "Item " + i.ToString();
                if (!isSetDefaultAbout)
                {
                    oListItem["About"] = "About " + i.ToString();
                }
                if (!isSetDefaultCity)
                {
                    Field cityField = oList.Fields.GetByTitle("City");
                    TaxonomyField txField = ctx.CastTo<TaxonomyField>(cityField);
                    TaxonomyFieldValue tValue = new TaxonomyFieldValue();

                    //Calculate index for term in term set

                    int x = i % terms.Count;

                    tValue.Label = terms[x].Name;
                    tValue.TermGuid = terms[x].Id.ToString();
                    tValue.WssId = -1;
                    txField.SetFieldValueByValue(oListItem, tValue);
                }
                // Assign values to Cities Field
                Field citiesField = oList.Fields.GetByTitle("Cities");
                TaxonomyField citiesTaxoField = ctx.CastTo<TaxonomyField>(citiesField);

                citiesTaxoField.SetFieldValueByTermCollection(oListItem, terms, 1033);

                oListItem.Update();
            }


            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Success");
        }

        // 8 Update site field “about” set default value for it to “about default” then create 2 new list items
        static async Task SetDefaultValueAboutField(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            Field field = list.Fields.GetByTitle("About");
            field.DefaultValue = "about default";
            field.Update();
            await ctx.ExecuteQueryAsync();
            await CreateListItem(ctx, 2, true, false);
            Console.WriteLine("Success");
        }

        // 9 Update site field “city” set default value for it to “Ho Chi Minh” then create 2 new list items.
        static async Task SetDefaultValueCityField(ClientContext ctx)
        {

            TermCollection terms = await GetTerms("city-ngocle",ctx);

            Term hcmTerm = terms.Where(term => term.Name == "Ho Chi Minh").First();

            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            Field cityfield = list.Fields.GetByTitle("City");
            TaxonomyField field = ctx.CastTo<TaxonomyField>(cityfield);
            TaxonomyFieldValue defaultValue = new TaxonomyFieldValue();
            defaultValue.Label = hcmTerm.Name;
            defaultValue.TermGuid = hcmTerm.Id.ToString();
            defaultValue.WssId = -1;

            //retrieve validated taxonomy field value
            var validatedValue = field.GetValidatedString(defaultValue);
            await ctx.ExecuteQueryAsync();
            field.DefaultValue = validatedValue.Value;
            field.Update();
            await ctx.ExecuteQueryAsync();

            // Create 2 items with default value for About and City field
            await CreateListItem(ctx, 2, true, true);
        }

        #endregion

        #region Exercise 2

        // 1 Write CAML query to get list items where field “about” is not “about default”

        static async Task GetListItemWhereNotDefautAbout(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            ListItemCollection items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View><Query>
                                <Where>
                                    <Neq>
                                         <FieldRef Name='About' />
                                         <Value Type='Text'>about default</Value>
                                    </Neq>
                                </Where>
                            </Query></View>"
            });
            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
            foreach (ListItem item in items)
            {
                //Console.WriteLine($"{((TaxonomyFieldValue)item["City"]).Label} - {item["About"].GetType()} " +
                //    $"- {item["Name"]} - {item["ID"]}");
                Console.WriteLine($"{((TaxonomyFieldValue)item["City"]).GetType()} " +
                    $"- {item["About"]} - {item["ID"]}");
            }
        }

        // 2 Create List View by CSOM order item newest in top and only show list item where “city” field has
        // value “Ho Chi Minh”, View Fields: Id, Name, City, About
        static async Task CreateListView(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");

            ViewCollection viewCollection = list.Views;
            ctx.Load(viewCollection);

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = "CSOMAssigment";
            viewCreationInformation.ViewTypeKind = ViewType.None;
            viewCreationInformation.Query = @"
                                                <OrderBy>
                                                    <FieldRef Name='ID' Ascending='FALSE'/>
                                                </OrderBy>
                                                <Where>
                                                    <Eq>
                                                            <FieldRef Name='City' />
                                                            <Value Type='Text'>Ho Chi Minh</Value>
                                                    </Eq>
                                                </Where>
                                              ";
            string CommaSeparateColumnNames = "ID,Name,City,About";
            viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');
            View listView = viewCollection.Add(viewCreationInformation);

            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success");

        }

        // 3 Write function update list items in batch, try to update 2 items every time and update field
        // “about” which have value “about default” to “Update script”. (CAML)
        static async Task Update2Items(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");

            ListItemCollection items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View><Query>
                                <Where>
                                    <Eq>
                                         <FieldRef Name='About' />
                                         <Value Type='Text'>about default</Value>
                                    </Eq>
                                </Where>
                            </Query><RowLimit>2</RowLimit></View>"
            });
            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
            foreach (ListItem item in items)
            {
                item["About"] = "Update script";
                item.Update();
            }

            await ctx.ExecuteQueryAsync();

            foreach (ListItem item in items)
            {
                Console.WriteLine($"{item["About"]} - {item["ID"]}");
            }
        }

        // 4 Create new field “author” type people in list “CSOM Test” then migrate all list items to set
        // user admin to field “CSOM Test Author”
        static async Task CreateAuthorFieldAndMigrateData(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            Guid guid = Guid.NewGuid();
            string fieldAsXml = $"<Field ID='{guid}' Name='Author0' DisplayName='Author' Type='User' Hidden='False' Group='Custom Columns' Description='CSOM Test Author Description' />";
            Field field = list.Fields.AddFieldAsXml(fieldAsXml, true, AddFieldOptions.DefaultValue);
            ctx.Load(field);

            await ctx.ExecuteQueryAsync();

            // Set all items Author = userAdmin
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = list.GetItems(query);
            UserCollection users = ctx.Web.SiteUsers;
            ctx.Load(items);
            ctx.Load(users);
            await ctx.ExecuteQueryAsync();
            User userAdmin = null;
            foreach (User user in users)
            {
                if (user.IsSiteAdmin)
                {
                    userAdmin = user;
                }
            }
            foreach (ListItem item in items)
            {
                item["Author0"] = userAdmin;
                item.Update();
            }
            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success Update");
        }

        #endregion

        #region Exercises Advance

        // 1 Create Taxonomy Field which allow multi values, with name “cities” map to your termset
        static async Task CreateCitiesTaxonomySiteFields(ClientContext ctx)
        {
            Web web = ctx.Web;
            Guid guid = Guid.NewGuid();
            string fieldAsXml = $"<Field ID='{guid}' Name='Cities' DisplayName='Cities' Type='TaxonomyFieldType' Hidden='False' Group='Custom Columns' Description='Cities Multi Values Taxonomy Field' />";
            Field field = web.Fields.AddFieldAsXml(fieldAsXml, true, AddFieldOptions.DefaultValue);
            await ctx.ExecuteQueryAsync();

            (Guid, Guid) result = await GetTaxonomyFieldInfo(ctx);
            Guid termStoreId = result.Item1;
            Guid termSetId = result.Item2;

            // Cast as Taxonomy Field
            TaxonomyField citiesField = ctx.CastTo<TaxonomyField>(field);
            citiesField.SspId = termStoreId;
            citiesField.TermSetId = termSetId;
            citiesField.TargetTemplate = String.Empty;
            citiesField.AnchorId = Guid.Empty;
            citiesField.AllowMultipleValues = true;
            citiesField.Update();

            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success");
        }

        // 2 Add field “cities” to content type “CSOM Test content type” make sure don’t need update list
        // but added field should be available in your list “CSOM test”
        static async Task AddCitiesFieldToContentType(ClientContext ctx)
        {
            Field fieldCitiesAdded = ctx.Web.Fields.GetByTitle("Cities");
            ContentType ct = ctx.Web.ContentTypes.GetById("0x0100954DE3873C92F840AD95A8AEF6E420CF");

            ct.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldCitiesAdded
            });
            ct.Update(true);
            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success");
        }

        // 3 Add 3 list item to list “CSOM test” and set multi value to field “cities” 
        static async Task Add3ListItemsWithMultiValueCities(ClientContext ctx)
        {
            await CreateListItem(ctx, 3, true, false);
            Console.WriteLine("Success");
        }

        // 4 Create new List type Document lib name “Document Test” add content type “CSOM Test content type”
        // to this list.
        static async Task CreateListTypeDocument(ClientContext ctx)
        {
            Web web = ctx.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = "Document Test";
            creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
            List list = web.Lists.Add(creationInfo);
            list.Description = "Document Test List Description";

            list.Update();
            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Document List Created");

            // Add Content type to list
            ContentType ct = ctx.Web.ContentTypes.GetById("0x0100954DE3873C92F840AD95A8AEF6E420CF");
            list.ContentTypes.AddExistingContentType(ct);
            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Content type Added To list");
        }

        // 5 Create Folder “Folder 1” in root of list “Document Test” then create “Folder 2” inside “Folder 1”,
        // Create 3 list items in “Folder 2” with value “Folder test” in field “about”. Create 2 flies in
        // “Folder 2” with value “Stockholm” in field “cities”.
        static async Task CreateFolderAndItemInDocumentLib(ClientContext ctx)
        {
            TermCollection terms = await GetTerms("city-ngocle", ctx);

            List docLib = ctx.Web.Lists.GetByTitle("Document Test");
            // Create Folder 1
            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
            listItemCreationInformation.UnderlyingObjectType = FileSystemObjectType.Folder;
            listItemCreationInformation.LeafName = "Folder 1";
            ListItem folder1 = docLib.AddItem(listItemCreationInformation);
            folder1["Title"] = "Folder 1";
            folder1.Update();

            // Create Folder 2
            Folder folder2 = folder1.Folder.Folders.Add("Folder 2");
            folder1.Update();

            // Create 3 files with value “Folder test” in field “about”
            for (int i = 1; i <= 3; i++)
            {
                FileCreationInformation fileCreationInformation = new FileCreationInformation();
                fileCreationInformation.Url = $"file_about_{i}.txt";
                string somestring = "hello there";
                byte[] toBytes = Encoding.ASCII.GetBytes(somestring);
                fileCreationInformation.Content = toBytes;
                File addedFile = folder2.Files.Add(fileCreationInformation);
                ListItem item = addedFile.ListItemAllFields;
                item["About"] = "Folder test";
                item.Update();
            }

            // Create 2 files with value “Stockholm” in field “cities”.
            for (int i = 1; i <= 2; i++)
            {
                FileCreationInformation fileCreationInformation = new FileCreationInformation();
                fileCreationInformation.Url = $"file_cities_{i}.txt";
                string somestring = "hello there";
                byte[] toBytes = Encoding.ASCII.GetBytes(somestring);
                fileCreationInformation.Content = toBytes;
                File addedFile = folder2.Files.Add(fileCreationInformation);
                ListItem item = addedFile.ListItemAllFields;
                item["Cities"] = terms[1].Id.ToString();
                
                item.Update();
            }

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Success");
        }

        // 6 Write CAML get all list item just in “Folder 2” and have value “Stockholm” in “cities” field
        static async Task GetListItemFolder2(ClientContext ctx)
        {
            List docLib = ctx.Web.Lists.GetByTitle("Document Test");
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View><Query>
                                <Where>
                                    <Eq>
                                         <FieldRef Name='Cities' />
                                         <Value Type='Text'>Stockholm</Value>
                                    </Eq>
                                </Where>
                            </Query></View>";
            query.FolderServerRelativeUrl = "/sites/ngocleprecio/Document Test/Folder 1/Folder 2";
            ListItemCollection listItems = docLib.GetItems(query);
            ctx.Load(listItems);
            await ctx.ExecuteQueryAsync();
            foreach (ListItem item in listItems)
            {
                Console.WriteLine(item.FieldValues["FileLeafRef"]);
            }
        }

        // 7 Create List Item in “Document Test” by upload a file Document.docx
        static async Task UploadFileToDocumentLib(ClientContext ctx)
        {
            string filePath = @"D:\instruction\training\Sharepoint\Document.docx";

            if (!System.IO.File.Exists(filePath))
            {
                throw new System.IO.FileNotFoundException(filePath);
            }
            string fileName = System.IO.Path.GetFileName(filePath);

            byte[] fileStream = System.IO.File.ReadAllBytes(filePath);
            List doclib = ctx.Web.Lists.GetByTitle("Document Test");
            Folder docFolder = doclib.RootFolder;
            FileCreationInformation fileCreationInformation = new FileCreationInformation();
            fileCreationInformation.Url = fileName;
            fileCreationInformation.Content = fileStream;
            File addedFile = docFolder.Files.Add(fileCreationInformation);

            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success");
            
        }

        #endregion

        #region Exercises Optional

        // 1 Create View “Folders” in List “Document Test” which only show folder structure,
        // and set this view as default
        static async Task CreateListViewFoldersSetDefault(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("Document Test");

            ViewCollection viewCollection = list.Views;
            ctx.Load(viewCollection);

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = "Folders";
            viewCreationInformation.ViewTypeKind = ViewType.None;
            viewCreationInformation.SetAsDefaultView = true;
            viewCreationInformation.Query = @"
                                                <Where>
                                                    <Eq>
                                                            <FieldRef Name='FSObjType' />
                                                            <Value Type='Number'>1</Value>
                                                    </Eq>
                                                </Where>
                                            
                                              ";
            View listView = viewCollection.Add(viewCreationInformation);
            await ctx.ExecuteQueryAsync();
            Console.WriteLine("Success");
        }

        // 2 Write code to load User from user email or name
        static async Task LoadUserFromEmailOrName(ClientContext ctx, string searctText)
        {
            try
            {
                var user = Utility.ResolvePrincipal(ctx, ctx.Web, searctText, PrincipalType.User, PrincipalSource.All, ctx.Web.SiteUsers, true);

                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"{user.Value.DisplayName} - {user.Value.Email}");
            }
            catch (Exception)
            {
                Console.WriteLine("Not found");
            }
            

        }

        #endregion

        #region RefCode

        static async Task<string> GetCityFieldGuid(ClientContext ctx)
        {
            FieldCollection fields = ctx.Web.Fields;
            string guid = "";
            ctx.Load(fields);

            await ctx.ExecuteQueryAsync();
            foreach (Field field in fields)
            {
                if (field.Description == "City Taxonomy Field")
                {
                    guid = field.Id.ToString();
                }

            }
            return guid;
        }

        static async Task<TermCollection> GetTerms(string termSetName,ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermSet termSet = taxonomySession.GetTermSetsByName(termSetName, 1033).GetByName(termSetName);
            TermCollection terms = termSet.Terms;

            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
            return terms;
        }

        static async Task<ListItemCollection> GetTermsUsingTaxonomyHiddenList(string idForTermSet,ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("TaxonomyHiddenList");
            CamlQuery query = new CamlQuery();
            query.ViewXml = @$"<View><Query>
                                <Where>
                                    <Eq>
                                         <FieldRef Name='IdForTermSet' />
                                         <Value Type='Text'>{idForTermSet}</Value>
                                    </Eq>
                                </Where>
                            </Query></View>";
            ListItemCollection terms = list.GetItems(query);
            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
            return terms;

        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task GetFieldTermValue(ClientContext Ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            Ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await Ctx.ExecuteQueryAsync();
        }

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        {
            var field = ctx.Web.Fields.GetByTitle("fieldname");

            ctx.Load(field);
            await ctx.ExecuteQueryAsync();

            var taxField = ctx.CastTo<TaxonomyField>(field);

            taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "correct label here",
                TermGuid = "term id"
            });
            item.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");

            var terms = termSet.GetAllTerms();

            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }

        #endregion


    }
}
