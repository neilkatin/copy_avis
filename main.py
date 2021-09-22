#!/usr/bin/env python

import os
import re
import sys
import logging
import argparse
import datetime
import json
import io
import zipfile

import openpyxl
import openpyxl.utils
import openpyxl.styles
import openpyxl.styles.colors
import openpyxl.writer.excel
from openpyxl.comments import Comment

import neil_tools
from neil_tools import spreadsheet_tools

import config as config_static

import arc_o365
import O365



NOW = datetime.datetime.now().astimezone()
DATESTAMP = NOW.strftime("%Y-%m-%d")
TIMESTAMP = NOW.strftime("%Y-%m-%d %H:%M:%S %Z")
FILESTAMP = NOW.strftime("%Y-%m-%d %H-%M-%S %Z")
EMAILSTAMP = NOW.strftime("%Y-%m-%d %H-%M")


def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    account_avis = init_o365(config, config.TOKEN_FILENAME_AVIS)
    log.debug(f"after initializing account")

    sharepoint = account_avis.sharepoint()

    source_site = sharepoint.get_site(config.SHAREPOINT_SITE, config.SOURCE_SITE)
    source_drive = source_site.get_default_document_library(request_drive=True)

    log.debug(f"source_drive { source_drive }")

    source_folder = get_folder(source_drive, config.SOURCE_PATH)

    log.debug(f"source_folder { source_folder }")


    dest_site = sharepoint.get_site(config.SHAREPOINT_SITE, config.DEST_SITE)
    dest_drive = dest_site.get_default_document_library(request_drive=True)
    dest_folder = get_folder(dest_drive, config.DEST_PATH)

    process_source(source_folder, dest_folder)

def process_source(folder, dest_folder):

    items = folder.get_items()

    avis_regex = re.compile(r'^FY\d\d Avis Reports?')

    for item in items:
        name = item.name
        size = item.size

        is_avis_folder = (avis_regex.match(name) is not None)

        #log.debug(f"item name { name } size { size } is_avis_folder { is_avis_folder }")

        if is_avis_folder:
            process_avis_folder(item, dest_folder)

def open_child(folder, child_name):

    drive = folder.drive
    parent_path = folder.parent_path
    name = folder.name

    new_path = "/".join( [ parent_path, name, child_name ] )

    sub_folder = drive.get_item_by_path(new_path)

    return sub_folder

def process_avis_folder(source_folder, dest_parent):
    name = source_folder.name


    query = dest_parent.new_query().on_attribute('name').equals(name)
    
    dest_children = dest_parent.get_items(query=query)
    #log.debug(f"dest_children { dest_children }")
    dest_folder = None
    for f in dest_children:
        dest_folder = f

    if dest_folder is None:
        dest_folder = dest_parent.create_child_folder(name)

    log.debug(f"dest_folder { dest_folder }")

    # cache all the file names
    dest_cache = {}
    for c in dest_folder.get_items():
        name = c.name
        dest_cache[name] = c

    items = source_folder.get_items()

    for item in items:
        name = item.name
        size = item.size

        #log.debug(f"checking item { name }")

        # see if it exists in the destination
        if name not in dest_cache:
            log.debug(f"copying '{ name }'")
            item.copy(target=dest_folder, name=name)



def get_folder(drive, path):

    folder = drive.get_item_by_path(path)
    log.debug(f"folder { folder } is_folder { folder.is_folder }")

    return folder


def init_o365(config, token_filename=None):
    """ do initial setup to get a handle on office 365 graph api """

    if token_filename != None:
        o365 = arc_o365.arc_o365.arc_o365(config, token_filename=token_filename)
    else:
        o365 = arc_o365.arc_o365.arc_o365(config)

    account = o365.get_account()
    if account is None:
        raise Exception("could not access office 365 graph api")

    return account

    

def parse_args():
    parser = argparse.ArgumentParser(
            description="tools to support Disaster Transportation Tools reporting",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")
    parser.add_argument("-s", "--store", help="Store file on server", action="store_true")

    args = parser.parse_args()

    return args


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)

