#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 12 17:01:39 2017

@author: jloranca
"""
from os.path import expanduser
from googleapiclient import discovery
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import pandas as pd


class Sheets(object):
    def __init__(self):
        self.scopes = 'https://www.googleapis.com/auth/spreadsheets'
        self.store = file.Storage(expanduser('~/storage.json'))

    def start(self):
        """
        :return: Initialize G-Sheet authorization using client secret file
        """
        creds = self.store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(expanduser('~/client_secrets.json'), self.scopes)
            creds = tools.run_flow(flow, self.store)
        http = creds.authorize(Http())
        sheets = discovery.build('sheets', 'v4', http=http)
        return sheets

    def append(self, data, sheet_id, target='A2'):
        sheet = self.start()
        update = data.values.tolist()
        values = {'values': [row for row in update]}
        sheet.spreadsheets().values().append(spreadsheetId=sheet_id, range=target, body=values,
                                             valueInputOption='USER_ENTERED').execute()

    def read(self, sheet_id, read, col=None):
        sheet = self.start()
        read_result = sheet.spreadsheets().values().get(spreadsheetId=sheet_id, range=read).execute()
        df = pd.DataFrame(read_result['values'])
        if col is not None:
            df.columns = col
        return df

    def update(self, data, sheet_id, target='A2'):
        sheet = self.start()
        update = data.values.tolist()
        values = {'values': [row for row in update]}
        sheet.spreadsheets().values().update(spreadsheetId=sheet_id, range=target, body=values,
                                             valueInputOption='USER_ENTERED').execute()

    def clear(self, sheet_id, target):
        sheet = self.start()
        sheet.spreadsheets().values().clear(spreadsheetId=sheet_id, range=target, body={}).execute()

    def bold(self, sheet_id, srow=0, erow=1, scol=0, ecol=1, bold=True):
        bolds = {'requests': [
            {'repeatCell': {
                'range': {'startRowIndex': srow,
                          'endRowIndex': erow,
                          'startColumnIndex': scol,
                          'endColumnIndex': ecol},
                'cell': {'userEnteredFormat': {'textFormat': {'bold': bold}}},
                'fields': 'userEnteredFormat.textFormat.bold',
            }}
        ]}
        sheet = self.start()
        sheet.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=bolds).execute()

    def borders(self, sheet_id, id='682809531', srow=1, erow=2, scol=1, ecol=2, style='SOLID_MEDIUM', inner='NONE'):
        body = {'requests': [{
                    "updateBorders": {
                      "range": {
                        "sheetId": id,
                        "startRowIndex": srow - 1,
                        "endRowIndex": erow - 1,
                        "startColumnIndex": scol - 1,
                        "endColumnIndex": ecol - 1},
                      "top": {
                        "style": style,
                        "color": {"blue": 0}},
                      "bottom": {
                        "style": style,
                        "color": {"blue": 0}},
                      "left": {
                        "style": style,
                        "color": {"blue": 0}},
                      "right": {
                        "style": style,
                        "color": {"blue": 0}},
                      "innerHorizontal": {
                        "style": inner,
                        "color": {"blue": 0}},
                      "innerVertical": {
                        "style": inner,
                        "color": {"blue": 0}}}}]}
        borders = {'requests': body['requests']}
        sheet = self.start()
        sheet.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=borders).execute()

    def insert(self, rows, sheet_id, target):
        sheet = self.start()
        l = []
        r = ['', '']
        for i in range(0, rows):
            l.append(r)
        data = pd.DataFrame(l)
        update = data.values.tolist()
        values = {'values': [row for row in update]}
        sheet.spreadsheets().values().append(spreadsheetId=sheet_id, range=target, body=values,
                                             valueInputOption='USER_ENTERED',
                                             insertDataOption='INSERT_ROWS').execute()

    def delete(self, sheet_id, id='682809531', srow=1, erow=2):
        sheet = self.start()
        delete = {
            "requests": [{
                "deleteDimension": {
                    "range": {
                        "sheetId": id,
                        "dimension": "ROWS",
                        "startIndex": srow - 1,
                        "endIndex": erow - 1}}}]}
        sheet.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=delete).execute()


class Drive(object):
    def __init__(self):
        self.scopes = 'https://www.googleapis.com/auth/drive'
        self.store = file.Storage(expanduser('~/credentials.json'))

    def start(self):
        """
        :return: Initialize G-Drive authorization using client secret file
        """
        creds = self.store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(expanduser('~/client_secrets.json'), self.scopes)
            creds = tools.run_flow(flow, self.store)
        drive = build('drive', 'v3', http=creds.authorize(Http()))
        return drive

    def download(self, file_id, filename, form='pdf'):
        drive = self.start
        request = drive().files().export(fileId=file_id, mimeType='application/{}'.format(form)).execute()
        with open(filename, 'wb') as fh:
            fh.write(request)

    def upload(self, filename, form='pdf', folder=None):
        drive = self.start
        metadata = {'name': filename, 'mimeType': 'application/{}'.format(form), 'parents': [folder]}
        result = drive().files().create(body=metadata, media_body=filename,
                                        media_mime_type=metadata['mimeType']).execute()
        if result:
            print('Uploaded {}'.format(filename))


if __name__ == '__main__':
    Sheets.start()
