from email import header
import os
from re import A
from msgraph import generate_access_token, GRAPH_API_ENDPOINT
import requests
import json
import base64
from typing import List
import sys
import os
sys.path.append(os.environ['USERPROFILE'] + '/OneDrive - The University of Nottingham/Documents/Programming/DLO/')
from addresses import CREDENTIALS_DIR


"""
Uses a rest API to make calls to Microscoft Azure to set permissions on onedrive folders

The Microsoft API:
Docs - https://learn.microsoft.com/en-us/graph/api/permission-grant?view=graph-rest-1.0&tabs=http

CREDENTIALS_DIR is a path containing the folders to be shared
GRAPH_API_ENDPOINT - web address for requesting security tokens. This is handled in msgraph.py
"""
    


def create_share_link(folder_path, emails : List, headers, link_type='view', permission="read"):
    """ Creates a share link for a particular folder path and adds read permission
    for the users indicated by emails.
    
    emails : comma separated list of email addresses

    **kwarg
    permission = "read", "write", "owner"
    """
    folder_id = get_folder_id(folder_path, headers)   
    
    request = {
        "type": link_type,
        "scope": "organization",
        }
    
    response=requests.post(
        GRAPH_API_ENDPOINT + '/me/drive/items/' + folder_id + '/createLink',
        headers=headers,
        json=request
        )
    
    share_link = response.json()['link']['webUrl']
    share_id = response.json()['shareId']
    add_permission(share_id, emails, headers, permission=permission)

    return share_link

def list_permission(folder_path, headers):   
    """Lists current permissions for a particular folder path.

    Parameters
    ----------
    folder_path : str
        should look like 'root:/Documents/DLO/Campus/modules/test'

    Returns
    -------
    _type_
        _description_
    """
    response=requests.get(
        GRAPH_API_ENDPOINT + '/me/drive/items/' + get_folder_id(folder_path) + ':/permissions',
        headers=headers,
    )
    return response
    
def remove_permission(folder_path, emails : List[str], headers):
    """Remove permission to user to use link

    Docs - https://learn.microsoft.com/en-us/graph/api/permission-grant?view=graph-rest-1.0&tabs=http

    Parameters
    ----------
    folder_path : str
        should look like 'root:/Documents/DLO/Campus/modules/test'
    user_email : List of emails to remove permission from

    """
    recipients = [{"email":email} for email in emails]

    folder_id = get_folder_id(folder_path, headers)

    response=requests.delete(
        GRAPH_API_ENDPOINT + '/me/drive/items/' + folder_id + '/permissions/grant',
        headers=headers,
        )

    return response 

def add_permission(share_id, emails : List, headers, permission="read"):
    """Grant permission to user to use link

    Docs - https://learn.microsoft.com/en-us/graph/api/permission-grant?view=graph-rest-1.0&tabs=http

    Parameters
    ----------
    share_id : An id linked to the share link
    emails : List of emails to grant permission to
    expirationDateTime : Needs implementing in request
    
    **kwarg
    permission = "read", "write", "owner"

    """
    recipients = [{"email":email} for email in emails]

    request = {
        "recipients": recipients,
        "roles": [permission],
        }

    

    response=requests.post(
        GRAPH_API_ENDPOINT + '/shares/' + share_id + '/permission/grant',
        headers=headers,
        json=request
        )

    return response

def get_folder_id(onedrive_filepath, headers):
    """Get the id of a folder in onedrive from its path
    
    Docs for api : https://learn.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_get?view=odsp-graph-online
    
    Parameters
    ----------
    onedrive_filepath : filepath should look like root:/Documents/DLO
        _description_
    """
    response=requests.get(
        GRAPH_API_ENDPOINT + '/me/drive/'+onedrive_filepath,
        headers=headers,
        )
   
    return response.json()['id']


def session_login():
    with open(CREDENTIALS_DIR + 'onedrive_api_key.json') as f:
        APPLICATION_ID = json.load(f)['API_KEY']
    
    #Scopes must match app permissions on https://portal.azure.com/
    #App is called Python Graph API
    SCOPES = ['Files.ReadWrite']

    #Verify identiry
    access_token = generate_access_token(APPLICATION_ID, SCOPES)

    headers = {
        'Authorization':'Bearer ' + access_token['access_token']
        }
    
    return headers

    
if __name__ == '__main__':
    headers = session_login()
    print(headers)
    folder_path = 'root:/Documents/DLO/Campus/modules/PHYS4007'
    #get_folder_id(folder_path)
    #response=list_permission(folder_path)
    #print(response.json())
    #print(response.json()['value'][0]['grantedToIdentitiesV2'][0]['user']['email'])
    
    link = create_share_link(folder_path,['ppzmrs@exmail.nottingham.ac.uk'], headers)
    print(link)

