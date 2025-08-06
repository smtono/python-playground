"""
Decodes contents of Google Docs to reveal a secret message
"""

from pathlib import Path
import os
import googleapiclient.discovery as discovery
from httplib2 import Http
from oauth2client import client
from oauth2client import file
from oauth2client import tools

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Grab google doc (From Google's API Documentation)
SCOPES = "https://www.googleapis.com/auth/documents.readonly"
DISCOVERY_DOC = "https://docs.googleapis.com/$discovery/rest?version=v1"
DOCUMENT_ID = "2PACX-1vRMx5YQlZNa3ra8dYYxmv-QIQ3YJe8tbI3kqcuC7lQiZm-CSEznKfN_HYNSpoXcZIV3Y_O3YoUB1ecq"


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth 2.0 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    # store = file.Storage("token.json")
    # credentials = store.get()

    # if not credentials or credentials.invalid:
    #     file_path = Path(os.getcwd(), 'credentials.json')
    #     flow = client.flow_from_clientsecrets(file_path, SCOPES)
    #     credentials = tools.run_flow(flow, store)
    # return credentials

    creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
        token.write(creds.to_json())


def add_current_and_child_tabs(tab, all_tabs):
    """Adds the provided tab to the list of all tabs, and recurses through and
    adds all child tabs.

    Args:
        tab: a Tab from a Google Doc.
        all_tabs: a list of all tabs in the document.
    """
    all_tabs.append(tab)
    for tab in tab.get('childTabs'):
        add_current_and_child_tabs(tab, all_tabs)

def get_all_tabs(doc):
    """Returns a flat list of all tabs in the document in the order they would
    appear in the UI (top-down ordering). Includes all child tabs.

    Args:
        doc: a document.
    """
    all_tabs = []
    # Iterate over all tabs and recursively add any child tabs to generate a
    # flat list of Tabs.
    for tab in doc.get('tabs'):
        add_current_and_child_tabs(tab, all_tabs)
    return all_tabs

def read_table(element):
    """Recurses through a list of Structural Elements to read a document's text
    where text may be in nested elements.

    Args:
        elements: a list of Structural Elements.
    """
    text = ""
    # The text in table cells are in nested Structural Elements and tables may
    # be nested.
    table = element.value.get("table")
    for row in table.get("tableRows"):
        cells = row.get("tableCells")
        for cell in cells:
            text += read_table(cell.get("content"))
    return text


# Parse data

# Sort coords

# Find how large grid is

# Fill in grid from sorted coords


if __name__ == "__main__":
    credentials = get_credentials()
    http = credentials.authorize(Http())
    docs_service = discovery.build(
        'docs', 'v1', http=http, discoveryServiceUrl=DISCOVERY_DOC
    )
    # Fetch the document with all of the tabs populated, including any nested
    # child tabs.
    doc = (
        docs_service.documents()
        .get(documentId=DOCUMENT_ID, include_tabs_content=True)
        .execute()
    )

    all_tabs = get_all_tabs(doc)

    # Print the text from each tab in the document.
    for tab in all_tabs:
        # Get the DocumentTab from the generic Tab.
        document_tab = tab.get('documentTab')
        doc_content = document_tab.get('body').get('content')
        print(read_table(doc_content))
