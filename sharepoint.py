from shareplum import Site, Office365
from shareplum.site import Version

import json
import os

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
config_path = '\\'.join([ROOT_DIR, 'config.json'])

# read json config file
with open(config_path) as config_file:
    config = json.load(config_file)
    config = config['share_point']

USERNAME = config['user']
PASSWORD = config['password']
SHAREPOINT_URL = config['url']
SHAREPOINT_SITE = config['site']


class SharePoint:
    def auth(self):
        self.authcookie = Office365(
            SHAREPOINT_URL,
            username=USERNAME,
            password=PASSWORD,
        ).GetCookies()
        self.site = Site(
            SHAREPOINT_SITE,
            version=Version.v365,
            authcookie=self.authcookie,
        )
        return self.site

    def connect_to_list(self, ls_name):
        self.auth_site = self.auth()

        list_data = self.auth_site.List(list_name=ls_name).GetListItems()

        return list_data
