# -*- coding: utf-8 -*-

import calendar
from datetime import datetime
import httplib2
import apiclient
from config import CREDENTIALS_FILE, SHEETS_FILE, CELL_LIST, DAY_TIME, USER
from oauth2client.service_account import ServiceAccountCredentials


class SheetWork:
    """

    """
    def __init__(self, spreadsheet_id=None, date=datetime.now(), function='update'):
        self.spreadsheetId = spreadsheet_id
        self.function = function
        credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])
        self.date = date
        self.httpAuth = credentials.authorize(httplib2.Http())  # api_auth
        self.service = apiclient.discovery.build('sheets',
                                                 'v4',
                                                 http=self.httpAuth)

    def week_day(self):
        cal = calendar.Calendar(firstweekday=0)
        mouth = cal.monthdayscalendar(self.date.year, self.date.month)
        day_list = []
        for week in mouth:
            for day in week:
                if day != 0:
                    week_day = datetime(self.date.year, self.date.month, day)
                    day_list.append(f'{day} - {week_day.strftime("%A")}')
        return day_list

    def create(self):
        spreadsheet = self.service.spreadsheets().create(
            body={'properties': {'title': 'Еimetable OrcMaster',
                                 'locale': 'ru_RU'},
                  'sheets': [{'properties': {'sheetType': 'GRID',
                                             'sheetId': 0,
                                             'title': self.date.strftime("%B"),
                                             'gridProperties': {'rowCount': 100,
                                                                'columnCount': 15}}}]
                  }).execute()
        return spreadsheet

    def add_sheets(self, name):
        sheet_id = 0
        result = self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheetId,
            body={"requests": [{"addSheet": {
                "properties": {"title": name,
                               "gridProperties": {
                                   "rowCount": 20,
                                   "columnCount": (len(self.week_day()) * 24) + 1}}}}
            ]}).execute()
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        sheet_list = spreadsheet.get('sheets')
        for mouth in range(len(sheet_list)):
            if name == sheet_list[mouth]['properties']['title']:
                sheet_id = sheet_list[mouth]['properties']['sheetId']
        return sheet_id

    def add_title(self, sheet_id):
        result = {"repeatCell": {"cell": {"userEnteredFormat": {
            "horizontalAlignment": 'CENTER',
            "backgroundColor": {
                "red": 0.8,
                "green": 0.8,
                "blue": 0.8,
                "alpha": 1},
            "textFormat": {
                "bold": True,
                "fontSize": 28}}},
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 1,
                "endColumnIndex": 4},
            "fields": "userEnteredFormat"}}
        return result

    def add_month(self, sheet_id):
        result = [{"repeatCell": {
                "cell": {"userEnteredFormat": {
                    "horizontalAlignment": 'CENTER',
                    "backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8, "alpha": 1},
                    "textFormat": {
                        "bold": True,
                        "fontSize": 14}}},
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 3,
                    "startColumnIndex": 1,
                    "endColumnIndex": (len(self.week_day()) * 24) + 1},
                "fields": "userEnteredFormat"}}]
        for day in range(len(self.week_day())):
            if self.week_day()[day].find('Sunday') != -1 or self.week_day()[day].find('Saturday') != -1:
                res = {"repeatCell": {
                    "cell": {"userEnteredFormat": {
                        "horizontalAlignment": 'CENTER',
                        "backgroundColor": {"red": 1.0, "green": 0.8, "blue": 0.8, "alpha": 1},
                        "textFormat": {
                            "bold": True,
                            "fontSize": 14}}},
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": 3,
                        "startColumnIndex": ((day + 1) * 24) - 23,
                        "endColumnIndex": ((day + 1) * 24) + 1},
                    "fields": "userEnteredFormat"}}
                result.append(res)
        return result

    def add_day_event(self, sheet_id, counts, day):
        start, end = 4, 5
        result = []
        for _ in range(counts):
            result = [{'mergeCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': start,
                          'endRowIndex': end,
                          'startColumnIndex': ((day + 1) * 24) - 23,
                          'endColumnIndex': ((day + 1) * 24) - 3},
                'mergeType': 'MERGE_ALL'}},
                {'mergeCells': {
                    'range': {'sheetId': sheet_id,
                              'startRowIndex': start,
                              'endRowIndex': end,
                              'startColumnIndex': ((day + 1) * 24) - 3,
                              'endColumnIndex': ((day + 1) * 24) - 1},
                    'mergeType': 'MERGE_ALL'}},
                {'mergeCells': {
                    'range': {'sheetId': sheet_id, 'startRowIndex': start,
                              'endRowIndex': end,
                              'startColumnIndex': ((day + 1) * 24) - 1,
                              'endColumnIndex': ((day + 1) * 24) + 1},
                    'mergeType': 'MERGE_ALL'}}]
            start = start + 1
            end = end + 1
        return result

    def balnce_day_night(self, sheet_id, day=0, day_time=6, night_time=22):
        result = {"repeatCell": {"cell": {"userEnteredFormat": {
            "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0, "alpha": 1}}},
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 3,
                "endRowIndex": 4,
                "startColumnIndex": ((day + 1) * 24) - 23 + (day_time - 1),
                "endColumnIndex": ((day + 1) * 24 - (24 - night_time))},
            "fields": "userEnteredFormat"}}, {
            "repeatCell": {"cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.5, "green": 0.0, "blue": 1.5, "alpha": 1}}},
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 3,
                    "endRowIndex": 4,
                    "startColumnIndex": ((day + 1) * 24) - 23,
                    "endColumnIndex": ((day + 1) * 24) - 23 + (day_time - 1)},
                "fields": "userEnteredFormat"}}, {
            "repeatCell": {"cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.5, "green": 0.0, "blue": 1.5, "alpha": 1}}},
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 3,
                    "endRowIndex": 4,
                    "startColumnIndex": ((day + 1) * 24 - (24 - night_time)),
                    "endColumnIndex": ((day + 1) * 24) + 1},
                "fields": "userEnteredFormat"}}
        return result

    def marge_calls(self, sheet_id):
        print((len(self.week_day()) * 24) + 2, 'CALLS')
        result = [{'mergeCells': {'range': {'sheetId': sheet_id,
                                            'startRowIndex': 0,
                                            'endRowIndex': 2,
                                            'startColumnIndex': 1,
                                            'endColumnIndex': (len(self.week_day()) * 24) + 1},
                                  'mergeType': 'MERGE_ALL'}},
                  {"updateDimensionProperties": {"range": {"sheetId": sheet_id,
                                                           "dimension": "COLUMNS",
                                                           "startIndex": 1,
                                                           "endIndex": (len(self.week_day()) * 24) + 1},
                                                 "properties": {"pixelSize": 20},
                                                 "fields": "pixelSize"}}]
        start, end = 1, 25
        for _ in range(len(self.week_day())):
            result.append({'mergeCells': {'range': {'sheetId': sheet_id,
                                                    'startRowIndex': 2,
                                                    'endRowIndex': 3,
                                                    'startColumnIndex': start,
                                                    'endColumnIndex': end},
                                          'mergeType': 'MERGE_ALL'}})
            start = end
            end = end + 24
        return result

    def rights(self, sheet_id, role='writer', email=USER):
        DriveService = apiclient.discovery.build('drive', 'v3', http=self.httpAuth)
        access = DriveService.permissions().create(
            fileId=sheet_id,
            body={'type': 'user', 'role': role, 'emailAddress': email},
            fields='id'
        ).execute()

    def event_new_sheet(self, name):
        result = []
        sheet_id = self.add_sheets(name=name)
        result.append(self.add_title(sheet_id))
        result = result + self.marge_calls(sheet_id) + self.add_month(sheet_id)
        for day in range(len(self.week_day())):
            result.append(self.balnce_day_night(
                sheet_id,
                day=day,
                day_time=DAY_TIME['end_night'],
                night_time=DAY_TIME['start_night']
            ))
        result = result + self.add_day_event(sheet_id, 4, 13)
        results = self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheetId,
            body={"requests": result}).execute()
        title = f'Schedule for {self.date.strftime("%B %Y")}'
        results = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body={
            "valueInputOption": "USER_ENTERED", "data": [
                {"range": f"{name}!B1",
                 "majorDimension": "ROWS",
                 "values": [[title]]}
            ]}).execute()
        for day in range(len(self.week_day())):
            results = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body={
                "valueInputOption": "USER_ENTERED", "data": [
                    {"range": f"{name}!{CELL_LIST[day]}3",
                     "majorDimension": "ROWS",
                     "values": [[self.week_day()[day]]]}
                ]}).execute()

    def _add(self, sheet_id):
        pass

    def _update(self, sheet_id, name):
        print(sheet_id, name)

    def _delete(self, sheet_id):
        pass

    def run(self):
        if self.spreadsheetId:
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
            sheet_list = spreadsheet.get('sheets')
            identy = (None, None)
            for mouth in range(len(sheet_list)):
                if self.date.strftime('%B') == sheet_list[mouth]['properties']['title']:
                    identy = (mouth, sheet_list[mouth]['properties']['sheetId'])
            if identy[1]:
                if self.function == 'add':
                    print(f'ADD EVENT: https://docs.google.com/spreadsheets/d/{self.spreadsheetId}')
                elif self.function == 'update':
                    self._update(sheet_id=sheet_list[identy[0]]['properties']['sheetId'],
                                      name=self.date.strftime('%B'))
                    print(f'UPDATE EVENT: https://docs.google.com/spreadsheets/d/{self.spreadsheetId}')
                elif self.function == 'delete':
                    print(f'DELETE EVENT: https://docs.google.com/spreadsheets/d/{self.spreadsheetId}')
                else:
                    print(f'NEED DEBUG BY SheetWork class!')
            else:
                self.event_new_sheet(name=self.date.strftime('%B'))
                print(f'CREATE SHEET {self.date.strftime("%B")}')
                self.run()
        else:
            self.spreadsheetId = self.create()['spreadsheetId']
            self.rights(self.spreadsheetId)
            print(f'CREATE FILE: https://docs.google.com/spreadsheets/d/{self.spreadsheetId}')
            self.run()


if __name__ == '__main__':
    try:
        daily_schedule = SheetWork(spreadsheet_id=SHEETS_FILE,
                                   date=datetime(2022, 12, 18),
                                   function='update')
        daily_schedule.run()
    except ValueError as exc:
        exit(f'ERROR by: {exc}')
