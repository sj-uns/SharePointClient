# Import required libraries
if __name__ == '__main__':
    import os
    import json
    import requests
    from requests.auth import HTTPBasicAuth


class SharePointClient:
    """
    SharePoint Client Class.

    Description:
    Allows users to connect to a SharePoint Site using the API Client Secret and:
    - Obtain lists of files recursively from a SharePoint site directory
    - Download / copy files from SharePoint site (preserving or flattening hierarchy)
    """
    def __init__(self, tenant:str, tenant_id:str, client_id:str, client_secret:str, sp_site:str):
        """
        Setup connection to SharePoint Site.

        Args:
            tenant (str): The SharePoint site tenant  (e.g. https://<tenant>.sharepoint.com)
            tenant_id (str): The Azure AD tenant ID.
            client_id (str): The client ID from Azure AD app registration.
            client_secret (str): The client secret from Azure AD app registration.
            sp_site (str): The SharePoint site identifier or name. This is the part of the SharePoint URL after '/sites/'.
                            For example, if your SharePoint site URL is 'https://<tenant>.sharepoint.com/sites/MySite',
                            then 'MySite' is the value to use for this parameter.
        """
        self.tenant = tenant
        self.tenant_id = tenant_id
        self.client_id = f'{client_id}@{tenant_id}'
        self.client_secret = client_secret
        self.sp_site = sp_site
      
        self.site_url = f'https://{tenant}.sharepoint.com' # Create the SharePoint Site URL
        self.access_token = self.get_access_token() # Initialise and store the access token


    def get_access_token(self):
        """
        Retrieve an access token for authenticating with SharePoint.

        Returns:
            access_token (str): The access token.
        
        Raises:
            Exception: If the token request fails.
        """

        # Get Access Token
        auth_body = {
            'grant_type':'client_credentials',
            'resource': f'00000003-0000-0ff1-ce00-000000000000/{self.tenant}.sharepoint.com@{self.tenant_id}', 
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        }

        auth_headers = {
            'Content-Type':'application/x-www-form-urlencoded'
        }

        auth_url = f"https://accounts.accesscontrol.windows.net/{self.tenant_id}/tokens/OAuth/2"
        auth_response = requests.post(auth_url, data=auth_body, headers=auth_headers, verify=False)

        auth_response.raise_for_status()
        token_result = auth_response.json()
        access_token = token_result.get('access_token')

        return access_token


    def get_sp_folder_contents(self, sp_folder_url:str, max_depth:int = -1, _depth:int = 0, _base_path:str = None):
        """
        Recursively retrieve the files from a SharePoint directory.
        If max_depth is set to 0, the entire directory structure will be scanned.

        Args:
            sp_folder_url (str): The server-relative URL of the file in SharePoint. It should include the path from
                                    the base SharePoint site URL (e.g., '/sites/MySite/Shared Documents/folder_name')
            max_depth (int, optional): The maximum depth to recurse into subfolders.
                                        (default is -1, meaning scan entire directory)
            _depth (int, internal): The current depth of recursion.
                                    This parameter is used internally and should not be set by the user. (default is 0)

        Returns:
            file_list (list): A list of dictionaries, each containing:
                - 'name' (str): Name of the file.
                - 'rel_path' (str): The relative path to the file, starting from the sp_folder_url argument.
                - 'server_relative_url' (str): The server-relative URL of the file.
                - 'sp_url_path' (str): The full SharePoint URL of the file (including the SharePoint site URL).
                - 'file_depth' (int): The depth of the file from the root folder.
            
            Raises:
                requests.HTTPError: If the folder path cannot be found in the SharePoint site.
        """

        # Get the base path from the sp_folder_path argument on first iteration
        if not _base_path:
            _base_path = sp_folder_url

        # Construct API request URL
        api_headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept':'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }

        api_url = rf"{self.site_url}/sites/{self.sp_site}/_api/Web/GetFolderByServerRelativeUrl('{sp_folder_url}')?$expand=Folders,Files"

        # Make the request to SharePoint API, to return the contents of the folder
        api_response = requests.get(api_url, headers=api_headers, verify=False)
        api_response.raise_for_status()  # Check response is valid (e.g. 200 == OK)
        api_result = api_response.json()

        # Extract Files and Folders
        files = api_result['d']['Files']['results']
        folders = api_result['d']['Folders']['results']

        # Prepare the list of files
        file_list = []
        
        # Add files to the file_list
        for file in files:
            file_list.append({
                'name': file['Name'],
                'sp_site': self.sp_site,
                'rel_path': os.path.relpath(file['ServerRelativeUrl'], base_path).replace('\\', '/'),
                'server_relative_url': file['ServerRelativeUrl'],
                'sp_url_path': self.site_url + file['ServerRelativeUrl'],
                'file_depth': _depth  # Depth of the file from the root folder as int
            })

        # Recursively retrieve contents of subfolders, only if within the max_depth 
        # or if max_depth is -1 scan entire directory
        if max_depth == -1 or _depth < max_depth:
            for folder in folders:
                subfolder_url = folder['ServerRelativeUrl']
                subfolder_contents = self.get_sp_folder_contents(sp_folder_url=subfolder_url, max_depth=max_depth, _depth=_depth + 1)
                
                # Add the subfolder files to the file_list
                file_list.extend(subfolder_contents)

        return file_list
    

    def download_sp_file(self, sp_file_url:str, target_dir:str, flatten:bool = False):
        """
        Download a file from SharePoint using its server-relative URL and save it to the target location.

        This method retrieves a file from a SharePoint site and saves it to a specified local directory. The file can
        be saved either with the directory structure preserved or flattened into a single directory, based on the
        `flatten` parameter.

        Args:
            sp_file_url (str): The server-relative URL of the file in SharePoint. It should include the path from
                            the base SharePoint site URL, such as '/sites/MySite/Shared Documents/folder/file_name.ext'.
            target_dir (str): The local directory path where the file will be saved.
            flatten (bool, optional): If True, save the file directly in the target directory, flattening the directory structure.
                                    If False, preserve the directory structure. (default is False)

        Returns:
            file_path (str): The target path where the file is saved.

        Raises:
            requests.HTTPError: If the request to SharePoint fails.
            OSError: If there's an error creating the directory or saving the file.
        """

        # Get the base path
        base_path = f'/sites/{self.sp_site}/Shared Documents'

        # Construct API request URL
        api_url = f'{self.site_url}/sites/{self.sp_site}/_api/web/GetFileByServerRelativeUrl("{sp_file_url}")/$value'
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json;odata=verbose'
        }

        # Make the request to SharePoint API, to return the file binary
        response = requests.get(api_url, headers=headers, verify=False, stream=True)
        response.raise_for_status()  # Check response is valid (e.g. 200 == OK)

        # Determine the file name and relative path
        file_name = os.path.basename(sp_file_url)
        if not flatten:
            # Preserve the directory structure
            rel_path = os.path.relpath(sp_file_url, base_path).replace('\\', '/')
            file_path = os.path.join(target_dir, rel_path)
        else:
            # Flatten the directory structure
            file_path = os.path.join(target_dir, file_name)

        # Ensure the target directory exists
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        # Write the binary to the target path
        with open(file_path, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)  # Write the file in chunks (8Kb / chunk)

        print(f"File downloaded successfully to {file_path}")
        return file_path
