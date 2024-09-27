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
        auth_response = requests.post(auth_url, data=auth_body, headers=auth_headers)

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
        api_response = requests.get(api_url, headers=api_headers)
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
    

    def download_sp_file(self, sp_file_url:str, target_dir:str):
        """
        Download a file from SharePoint using its server-relative URL and save it to the target location.

        Args:
            sp_file_url (str): The server-relative URL of the file in SharePoint. It should include the path from
                            the base SharePoint site URL, such as '/sites/MySite/Shared Documents/folder/file_name.ext'.
            target_dir (str): The local directory path where the file will be saved.

        Returns:
            file_path (str): The target path where the file is saved.

        Raises:
            requests.HTTPError: If the request to SharePoint fails.
            OSError: If there's an error creating the directory or saving the file.
        """

        # Construct API request URL
        api_url = f"{self.site_url}/sites/{self.sp_site}/_api/web/GetFileByServerRelativeUrl('{sp_file_url}')/$value"
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json;odata=verbose'
        }

        # Make the request to SharePoint API, to return the file binary
        response = requests.get(api_url, headers=headers, stream=True)
        response.raise_for_status()  # Check response is valid (e.g. 200 == OK)

        # Determine the file name and path
        file_name = os.path.basename(sp_file_url)
        file_path = os.path.join(target_dir, file_name)

        # Ensure the target directory exists
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        # Write the binary to the target path
        with open(file_path, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)  # Write the file in chunks (8Kb / chunk)

        print(f"File downloaded successfully to {file_path}")
        return file_path
    

    def download_sp_folder(self, sp_folder_url:str, target_dir:str, flatten:bool = False, max_depth:int = -1):
        """
        Download the contents of a folder from SharePoint using its server-relative URL and save them to the target location.

        This method retrieves the files from a folder and subfolders in a SharePoint site and saves it to a specified local directory.
        The file(s) can be saved either with the directory structure preserved or flattened into a single directory, based on the
        `flatten` parameter.

        Args:
            sp_folder_url (str): The server-relative URL of the file in SharePoint. It should include the path from
                                    the base SharePoint site URL (e.g., '/sites/MySite/Shared Documents/folder_name')
            target_dir (str): The local directory path where the file will be saved.
            flatten (bool, optional): If True, save the file directly in the target directory, flattening the directory structure.
                                    If False, preserve the directory structure. (default is False)
            max_depth (int, optional): The maximum depth to recurse into subfolders. (default is -1, meaning scan entire directory)

        Returns:
            files_downloaded (list): A list of dictionaries containing the details of the downloaded files.
        """

        # Initialise files_downloaded list
        files_downloaded = []

        # Get the list of files in the SharePoint folder
        files_to_get = self.get_sp_folder_contents(sp_folder_url=sp_folder_url, max_depth=max_depth)

        # Loop through the list of files and download each file to the target directory, based on the flatten parameter
        for file in files_to_get:
            if not flatten:
                target_path =\
                    self.download_sp_file(sp_file_url=file['server_relative_url'], target_dir=os.path.join(target_dir, os.path.dirname(file['rel_path'])))
            else:
                target_path =\
                    self.download_sp_file(sp_file_url=file['server_relative_url'], target_dir=target_dir)

            files_downloaded.append({
                'file_details': file,
                'target_path': target_path
            })

        print(f"Downloaded {len(files_downloaded)} files to {target_dir}")
        return files_downloaded
    

    def check_sp_folder_exists(self, sp_folder_url:str):
        """
        Check if a folder exists in SharePoint using its server-relative URL.

        Args:
            sp_folder_url (str): The server-relative URL of the folder in SharePoint. It should include the path from
                            the base SharePoint site URL, such as '/sites/MySite/Shared Documents/folder/file_name.ext'.

        Returns:
            bool: True if the folder exists, False otherwise.

        Raises:
            requests.HTTPError: If the request to SharePoint fails.
        """

        # Initialise variable to return if the folder exists
        exists = False

        # Construct the API request URL
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json;odata=verbose',
        }

        api_url = f"{self.site_url}/sites/{self.sp_site}/_api/web/GetFolderByServerRelativeUrl('{sp_folder_url}')"

        # Make the request to SharePoint API, to check if the folder exists
        response = requests.get(api_url, headers=headers)

        if response.status_code == 200:
            exists = True
        elif response.status_code == 404:
            exists = False
            response.raise_for_status()
        else:
            print(f"Error: {response.status_code} - {response.text}")

        return exists
    

    def create_sp_folder(self, new_folder_url:str):
        """
        Create a folder in SharePoint using a server-relative URL.

        Args:
            new_folder_url (str): The server-relative URL of the folder to be created in SharePoint. It should include the path from
                            the base SharePoint site URL, such as '/sites/MySite/Shared Documents/folder/file_name.ext'.

        Returns:
            bool: True if the folder has been created, False otherwise.

        Raises:
            requests.HTTPError: If the request to SharePoint fails.
        """

        # Check if the folder already exists
        if self.check_sp_folder_exists(new_folder_url):
            print(f"Folder '{new_folder_url}' already exists.")
            return False

        created = False

        # Construct the API request URL
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
            }
        
        payload = {
            "__metadata": { "type": "SP.Folder" },
            "ServerRelativeUrl": new_folder_url
            }
        
        api_url = f"{self.site_url}/sites/{self.sp_site}/_api/web/folders"

        # Make the request to SharePoint API, to create the folder
        response = requests.post(api_url, headers=headers, json=payload)

        if response.status_code == 201:
            print(f"Folder '{new_folder_url}' created successfully.")
            created = True
        else:
            print(f"Error: {response.status_code} - {response.text}")
            created = False
            response.raise_for_status()

        return created


    def move_sp_file(self, sp_file_url:str, target_folder_url:str, overwrite_flag:int = 1):
        """
        Move a file in SharePoint using its server-relative URL to the target folder server-relative URL.
        This will create the target folder if it does not exist.

        Args:
            sp_file_url (str): The server-relative URL of the file in SharePoint. It should include the path from
                            the base SharePoint site URL, such as '/sites/MySite/Shared Documents/folder/file_name.ext'.
            target_folder_url (str): The server-relative URL folder path where the file will be saved. It should include the path from
                            the base SharePoint site URL, such as '/sites/MySite/Shared Documents/folder'.
            overwrite_flag (int, optional): The overwrite flag to use when moving the file. Can be 0 (Do Not Overwrite) or 1 (Overwrite).

        Returns:
            file_path (str): The target path where the file is saved.

        Raises:
            requests.HTTPError: If the move request to SharePoint fails.
        """

        # Check if the target folder exists, if not then create the folder
        if not self.check_sp_folder_exists(target_folder_url):
            self.create_sp_folder(target_folder_url)

        # Construct the new file path
        target_file_url = os.path.join(target_folder_url, os.path.basename(sp_file_url))

        # Construct API request URL
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }

        payload = {
            'newUrl': target_file_url,
            'flags': overwrite_flag
        }

        api_url = f"{self.site_url}/sites/{self.sp_site}/_api/web/GetFileByServerRelativeUrl('{sp_file_url}')/moveTo"

        # Make the request to SharePoint API, to move the file
        response = requests.post(api_url, headers=headers, json=payload)

        # Check response is valid (e.g. 200 == OK)
        if response.status_code != 200:
            print(response.content)
            response.raise_for_status()

        print(f"File moved successfully to {target_file_url}")
        return target_file_url