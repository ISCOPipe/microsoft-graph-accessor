import msal
import os
from office365.graph_client import GraphClient
# Example using async credentials and application access.
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient #, GraphRequestAdapter
from msgraph_core import PageIterator

class MicrosoftGraphAccessor(object):
    def __init__(self,credentials):
        self.client = GraphClient(
            lambda:self.__acquire_token_func(
                **credentials))
        tenant_id, client_id, client_secret = [credentials[k] for k in ['tenant_name_or_id','client_id','client_secret']]
        self.msgraph_credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        scopes = ['https://graph.microsoft.com/.default']
        # self.msgraph_request_adapter = GraphRequestAdapter(self.msgraph_credential)
        self.msgraph_client = GraphServiceClient(credentials=self.msgraph_credential, scopes=scopes)

    def __acquire_token_func(self,tenant_name_or_id, client_id, client_secret):
        authority_url = f'https://login.microsoftonline.com/{tenant_name_or_id}'
        app = msal.ConfidentialClientApplication(
            authority=authority_url,
            client_id=client_id, 
            client_credential=client_secret
        )

        token = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])
        return token

    def _get_execute(func):
        def wrapper(self, *args, **kwargs):
            resource = func(self, *args, **kwargs)
            return resource.get().execute_query()
        return wrapper

    
    @_get_execute
    def get_root(self):
        return self.client.sites.root
    
    @_get_execute
    def get_root_sites(self):
        return self.client.root.sites
    
    def display_list_info(self,list_):
        print(list_.list)

    @_get_execute
    def list_sites_from_root(self,relative_root):
        return relative_root.sites

    def get_site_from_root(self,relative_root,site_name):
        sites = self.list_sites_from_root(relative_root)
        for site in sites:
            if site.name == site_name:
                return site
    
    @_get_execute
    def list_lists_from_site(self,site):
        return site.lists
    
    @_get_execute
    def get_list_from_site(self,site,list_name):
        #lists = self.list_lists_from_site(site)
        return site.lists.get().execute_query().get_by_name(list_name)

    @_get_execute
    def get_library_items(self,list_):
        return list_.drive.root.children
    
    # @_get_execute 
    # def get_list_items(self,list_):
    #     return list_.items
    
    async def get_all_list_items(self,site_id,list_id):
        items = []
        response = await self.msgraph_client.sites.by_site_id(site_id).lists.by_list_id(list_id).items.get()
        if response:
            items.extend(response.value)
            while response.odata_next_link:
                response = await self.msgraph_client.sites.by_site_id(site_id).lists.by_list_id(list_id).items.with_url(response.odata_next_link).get()#response.next_page()
                items.extend(response.value)
        return items
    
    async def get_list_item_fields(self,site_id,list_id,item_id,semaphore,request_configuration=None):
        async with semaphore:
            response = await self.msgraph_client.sites.by_site_id(site_id).lists.by_list_id(list_id).items.by_list_item_id(item_id).get(request_configuration=request_configuration)
        return response
    
    async def get_list_item_attachments(self, site_id, list_id, item_id, semaphore):
        """Fetches attachments for a specific list item asynchronously."""
        return "NOT IMPLEMENTED"
        async with semaphore:
            try:
                # Fetch the attachments using the MS Graph client
                response = await self.msgraph_client.sites.by_site_id(site_id).lists.by_list_id(list_id).items.by_list_item_id(item_id).attachments.get()
                if response:
                    return response.value  # Return the list of attachments
                else:
                    print(f"No attachments found for item {item_id} in site {site_id} and list {list_id}.")
                    return []
            except Exception as e:
                print(f"Error fetching attachments for item {item_id} in site {site_id} and list {list_id}: {e}")
                return e
        
    @_get_execute
    def get_list_operations(self,list_):
        return list_.operations

    
    @_get_execute
    def get_drive(self,user_id_or_principal_name):
        return self.client.users[f"{user_id_or_principal_name}"].drive
    
    @_get_execute
    def get_drive_list(self):
        return self.client.drives
    
    @_get_execute
    def get_drive_list_from_site(self,relative_root):
        return relative_root.drives
    
    def get_drive_items(self,drive):
        try:
            drive_items = drive.root.children.get().execute_query()
        except:
            drive_items = drive.root.get_files().execute_query()
        return drive_items
    
    @_get_execute
    def get_driveitem_children(self,drive_item):
        return drive_item.children
    
    def download_item(self,item,to_path,to_s3=False):
        if not to_s3:
            path = os.path.join(to_path,item.name)
            with open(path,'wb') as local_file:
                item.download(local_file).execute_query()
                return os.path.exists(path)
    
    def get_users(self):
        #Need permissions for this to work
        raise NotImplementedError("This method is not implemented")
        users = self.client.users.get().execute_query()
        return users