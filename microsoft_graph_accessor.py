import msal
import os
from office365.graph_client import GraphClient

class MicrosoftGraphAccessor(object):
    def __init__(self,credentials):
        self.client = GraphClient(
            lambda:self.__acquire_token_func(
                **credentials))

    def __acquire_token_func(self,tenant_name_or_id, client_id, client_secret):
        authority_url = f'https://login.microsoftonline.com/{tenant_name_or_id}'
        app = msal.ConfidentialClientApplication(
            authority=authority_url,
            client_id=client_id, 
            client_credential=client_secret
        )

        token = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])
        return token
    
    def get_root(self):
        return self.client.sites.root.get().execute_query()
    
    def get_root_sites(self):
        return self.client.root.sites.get().execute_query()

    def list_sites_from_root(self,relative_root):
        return relative_root.sites.get().execute_query()
    
    def list_lists_from_site(self,site):
        lists = site.lists.get().execute_query()
        return lists
    
    def get_list_items(self,list_):
        list_items = list_.drive.root.children.get().execute_query()
        return list_items
    
    def get_drive(self,user_id_or_principal_name):
        drive = self.client.users[f"{user_id_or_principal_name}"].drive.get().execute_query()
        return drive
    
    def get_drive_list(self):
        drives = self.client.drives.get().execute_query()
        return drives
    
    def get_drive_items(self,drive):
        try:
            drive_items = drive.root.children.get().execute_query()
        except:
            drive_items = drive.root.get_files().execute_query()
        return drive_items
    
    def get_driveitem_children(self,drive_item):
        return drive_item.children.get().execute_query()
    
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
    