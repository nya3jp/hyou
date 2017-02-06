# Copyright 2015 Google Inc. All rights reserved
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

from __future__ import (
    absolute_import, division, print_function, unicode_literals)
from builtins import (  # type: ignore  # noqa: F401
    ascii, bytes, chr, dict, filter, hex, input, int, list, map, next,
    object, oct, open, pow, range, round, str, super, zip)

import datetime
from typing import Any, AnyStr, Dict, Generator, Optional, Sequence, Tuple

import future.utils
import googleapiclient.discovery
import httplib2

from . import util


SHEETS_API_DISCOVERY_URL = (
    'https://sheets.googleapis.com/$discovery/rest?version=v4')

# For compatibility.
GOOGLE_SPREADSHEET_SCOPES = util.SCOPES


def to_native_str(s: str) -> str:
    return future.utils.text_to_native_str(s, encoding='utf-8')


class API(object):

    def __init__(self, http: httplib2.Http) -> None:
        self.sheets = googleapiclient.discovery.build(
            'sheets', 'v4', http=http,
            discoveryServiceUrl=SHEETS_API_DISCOVERY_URL)
        self.drive = googleapiclient.discovery.build(
            'drive', 'v2', http=http)


class Collection(util.LazyOrderedDictionary[str, 'Spreadsheet']):

    def __init__(self, api: API) -> None:
        super(Collection, self).__init__(
            self._spreadsheet_enumerator,
            self._spreadsheet_constructor)
        self._api = api

    @classmethod
    def login(cls,
              json_path: Optional[AnyStr]=None,
              json_text: Optional[str]=None) -> 'Collection':
        if json_path is not None:
            assert json_text is None
            with open(json_path, 'r') as f:
                json_text = f.read()
        assert json_text is not None
        credentials = util.parse_credentials(json_text)
        http = credentials.authorize(httplib2.Http())
        return cls(API(http))

    def create_spreadsheet(
            self, title: str, rows: int=1000, cols: int=26) -> 'Spreadsheet':
        body = {
            'title': title,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
        }
        response = self._api.drive.files().insert(body=body).execute()
        key = response['id']
        self.refresh()
        spreadsheet = self[key]
        spreadsheet[0].set_size(rows, cols)
        return spreadsheet

    def _spreadsheet_enumerator(
            self) -> Generator[Tuple[str, 'Spreadsheet'], None, None]:
        response = self._api.drive.files().list(
            maxResults=1000,
            q=('mimeType="application/vnd.google-apps.spreadsheet" and '
               'trashed = false'),
            fields='items/id').execute()
        for item in response['items']:
            key = item['id']
            yield (key, Spreadsheet(self._api, key, None))

    def _spreadsheet_constructor(self, key: str) -> 'Spreadsheet':
        entry = self._api.sheets.spreadsheets().get(
            spreadsheetId=key, includeGridData=False).execute()
        return Spreadsheet(self._api, entry['spreadsheetId'], entry)


class Spreadsheet(util.LazyOrderedDictionary[str, 'Worksheet']):

    def __init__(self, api: API, key: str, entry: Optional[Dict]) -> None:
        super(Spreadsheet, self).__init__(self._worksheet_enumerator, None)
        self._api = api
        self._key = key
        self._entry = entry
        self._updated = None  # type: Optional[datetime.datetime]

    def __repr__(self) -> str:
        return to_native_str(
            '<%s key="%s">' % (self.__class__.__name__, self.key))

    def refresh(self, entry: Dict=None) -> None:
        if entry is not None:
            self._entry = entry
        else:
            self._entry = self._api.sheets.spreadsheets().get(
                spreadsheetId=self.key, includeGridData=False).execute()
        self._updated = None
        super(Spreadsheet, self).refresh()

    def add_worksheet(
            self, title: str, rows: int=1000, cols: int=26) -> 'Worksheet':
        new_entry = self._make_single_batch_request(
            'addSheet',
            {
                'properties': {
                    'title': title,
                    'gridProperties': {
                        'rowCount': rows,
                        'columnCount': cols,
                    },
                },
            })
        self.refresh(new_entry)
        return self[title]

    def delete_worksheet(self, title: str) -> None:
        worksheet = self[title]
        new_entry = self._make_single_batch_request(
            'deleteSheet',
            {'sheetId': worksheet.key})
        self.refresh(new_entry)

    @property
    def key(self) -> str:
        return self._key

    @property
    def url(self) -> str:
        return 'https://docs.google.com/spreadsheets/d/%s/edit' % self.key

    @property
    def title(self) -> str:
        self._ensure_entry()
        assert self._entry is not None
        return self._entry['properties']['title']

    @title.setter
    def title(self, new_title: str) -> None:
        new_entry = self._make_single_batch_request(
            'updateSpreadsheetProperties',
            {
                'properties': {
                    'title': new_title,
                },
                'fields': 'title',
            })
        self.refresh(new_entry)

    @property
    def updated(self) -> datetime.datetime:
        if not self._updated:
            response = self._api.drive.files().get(fileId=self.key).execute()
            self._updated = datetime.datetime.strptime(
                response['modifiedDate'], '%Y-%m-%dT%H:%M:%S.%fZ')
        return self._updated

    def _ensure_entry(self) -> None:
        if self._entry is None:
            self.refresh()

    def _worksheet_enumerator(
            self) -> Generator[Tuple[str, 'Worksheet'], None, None]:
        self._ensure_entry()
        assert self._entry is not None
        for sheet_entry in self._entry['sheets']:
            worksheet = Worksheet(self, self._api, sheet_entry)
            yield (worksheet.title, worksheet)

    def _make_single_batch_request(self, method: str, params: Dict) -> Dict:
        request = {
            'requests': [{method: params}],
            'include_spreadsheet_in_response': True,
        }
        response = self._api.sheets.spreadsheets().batchUpdate(
            spreadsheetId=self.key, body=request).execute()
        return response['updatedSpreadsheet']


class WorksheetView(Sequence['WorksheetViewRow']):

    def __init__(self, worksheet: 'Worksheet', api: API,
                 start_row: int, end_row: int,
                 start_col: int, end_col: int) -> None:
        self._worksheet = worksheet
        self._api = api
        self._reset_size(start_row, end_row, start_col, end_col)
        self._input_value_map = {}  # type: Dict[Tuple[int, int], str]
        self._cells_fetched = False
        self._queued_updates = []  # type: List[Tuple[int, int, str]]

    def refresh(self) -> None:
        self._input_value_map.clear()
        self._cells_fetched = False
        del self._queued_updates[:]

    def _reset_size(
            self, start_row: int, end_row: int,
            start_col: int, end_col: int) -> None:
        self.start_row = start_row
        self.end_row = end_row
        self.start_col = start_col
        self.end_col = end_col
        self._view_rows = [
            WorksheetViewRow(self, row, start_col, end_col)
            for row in range(start_row, end_row)]

    def _ensure_cells_fetched(self) -> None:
        if self._cells_fetched:
            return
        range_str = util.format_range_a1_notation(
            self._worksheet.title, self.start_row, self.end_row,
            self.start_col, self.end_col)
        response = self._api.sheets.spreadsheets().values().get(
            spreadsheetId=self._worksheet._spreadsheet.key,
            range=to_native_str(range_str),
            majorDimension='ROWS',
            valueRenderOption='FORMATTED_VALUE',
            dateTimeRenderOption='FORMATTED_STRING').execute()
        self._input_value_map = {}
        for i, row in enumerate(response.get('values', [])):
            index_row = self.start_row + i
            for j, value in enumerate(row):
                index_col = self.start_col + j
                self._input_value_map.setdefault((index_row, index_col), value)
        self._cells_fetched = True

    def commit(self) -> None:
        if not self._queued_updates:
            return
        request = {
            'data': [
                {
                    'range': util.format_range_a1_notation(
                        self._worksheet.title, row, row + 1, col, col + 1),
                    'majorDimension': 'ROWS',
                    'values': [[value]],
                }
                for row, col, value in self._queued_updates
            ],
            'valueInputOption': 'USER_ENTERED',
            'includeValuesInResponse': False,
        }
        self._api.sheets.spreadsheets().values().batchUpdate(
            spreadsheetId=self._worksheet._spreadsheet.key,
            body=request).execute()
        del self._queued_updates[:]

    def __nonzero__(self) -> bool:
        return len(self) > 0

    def __getitem__(self, index: Any) -> Any:
        return self._view_rows[index]

    def __len__(self) -> int:
        return self.rows

    def __iter__(self) -> Generator['WorksheetViewRow', None, None]:
        for row in self._view_rows:
            yield row

    def __repr__(self) -> str:
        return repr(self._view_rows)

    def __enter__(self) -> 'WorksheetView':
        return self

    def __exit__(self, exc_type: Any, exc_value: Any, traceback: Any) -> None:
        self.commit()

    @property
    def rows(self) -> int:
        return self.end_row - self.start_row

    @property
    def cols(self) -> int:
        return self.end_col - self.start_col


class WorksheetViewRow(util.CustomMutableFixedList[str]):

    def __init__(self, view: WorksheetView, row: int,
                 start_col: int, end_col: int) -> None:
        self._view = view
        self._row = row
        self._start_col = start_col
        self._end_col = end_col

    def __nonzero__(self) -> bool:
        return len(self) > 0

    def __getitem__(self, index: Any) -> Any:
        if isinstance(index, slice):
            start, stop, step = index.indices(len(self))  # type: ignore
            assert step == 1, 'slicing with step is not supported'
            if stop < start:
                stop = start
            return WorksheetViewRow(
                self._view, self._row,
                self._start_col + start, self._start_col + stop)
        assert isinstance(index, int)
        if index < 0:
            col = self._end_col + index
        else:
            col = self._start_col + index
        if not (self._start_col <= col < self._end_col):
            raise IndexError()
        if (self._row, col) not in self._view._input_value_map:
            self._view._ensure_cells_fetched()
        return self._view._input_value_map.get((self._row, col), '')

    def __setitem__(self, index: Any, new_value: Any) -> None:
        if isinstance(index, slice):
            start, stop, step = index.indices(len(self))  # type: ignore
            assert step == 1, 'slicing with step is not supported'
            if stop < start:
                stop = start
            if len(new_value) != stop - start:
                raise ValueError(
                    'Tried to assign %d values to %d element slice' %
                    (len(new_value), stop - start))
            for i, new_value_one in zip(range(start, stop), new_value):
                self[i] = new_value_one
            return
        assert isinstance(index, int)
        if index < 0:
            col = self._end_col + index
        else:
            col = self._start_col + index
        if not (self._start_col <= col < self._end_col):
            raise IndexError()
        if new_value is None:
            new_value = ''
        elif isinstance(new_value, int):
            new_value = '%d' % new_value
        elif isinstance(new_value, float):
            # Do best not to lose precision...
            new_value = '%.20e' % new_value
        elif isinstance(new_value, bytes):
            # May raise UnicodeDecodeError.
            new_value = new_value.decode('ascii')
        elif not isinstance(new_value, str):
            new_value = str(new_value)
        self._view._input_value_map[(self._row, col)] = new_value
        self._view._queued_updates.append((self._row, col, new_value))

    def __len__(self) -> int:
        return self._end_col - self._start_col

    def __iter__(self) -> Generator[str, None, None]:
        self._view._ensure_cells_fetched()
        for col in range(self._start_col, self._end_col):
            yield self._view._input_value_map.get((self._row, col), '')

    def __repr__(self) -> str:
        return repr([self[i] for i in range(len(self))])


class Worksheet(WorksheetView):

    def __init__(
            self, spreadsheet: Spreadsheet, api: API, entry: Dict) -> None:
        self._spreadsheet = spreadsheet
        self._api = api
        self._entry = entry
        super(Worksheet, self).__init__(self, api, 0, self.rows, 0, self.cols)

    def refresh(self, entry: Optional[Dict]=None) -> None:
        if entry is not None:
            self._entry = entry
        else:
            spreadsheet_entry = self._api.sheets.spreadsheets().get(
                spreadsheetId=self._spreadsheet.key,
                includeGridData=False).execute()
            for new_entry in spreadsheet_entry['sheets']:
                if new_entry['properties']['sheetId'] == self.key:
                    self._entry = new_entry
                    break
            else:
                raise KeyError('Sheet has been removed')
        self._reset_size(0, self.rows, 0, self.cols)
        super(Worksheet, self).refresh()

    def view(self,
             start_row: Optional[int]=None,
             end_row: Optional[int]=None,
             start_col: Optional[int]=None,
             end_col: Optional[int]=None) -> WorksheetView:
        if start_row is None:
            start_row = 0
        if end_row is None:
            end_row = self.rows
        if start_col is None:
            start_col = 0
        if end_col is None:
            end_col = self.cols
        if not (0 <= start_row <= end_row <= self.rows):
            raise IndexError()
        if not (0 <= start_col <= end_col <= self.cols):
            raise IndexError()
        return WorksheetView(
            self, self._api,
            start_row=start_row, end_row=end_row,
            start_col=start_col, end_col=end_col)

    def set_size(self, rows: int, cols: int) -> None:
        assert isinstance(rows, int) and rows > 0
        assert isinstance(cols, int) and cols > 0
        new_entry = self._make_single_batch_request(
            'updateSheetProperties',
            {
                'properties': {
                    'sheetId': self.key,
                    'gridProperties': {
                        'rowCount': rows,
                        'columnCount': cols,
                    },
                },
                'fields': 'gridProperties(rowCount,columnCount)',
            })
        self.refresh(new_entry)

    @property
    def key(self) -> str:
        return self._entry['properties']['sheetId']

    @property
    def title(self) -> str:
        return self._entry['properties']['title']

    @title.setter
    def title(self, new_title: str) -> None:
        new_entry = self._make_single_batch_request(
            'updateSheetProperties',
            {
                'properties': {
                    'sheetId': self.key,
                    'title': new_title,
                },
                'fields': 'title',
            })
        self.refresh(new_entry)

    @property
    def rows(self) -> int:
        return self._entry['properties']['gridProperties']['rowCount']

    @rows.setter
    def rows(self, rows: int) -> None:
        self.set_size(rows, self.cols)

    @property
    def cols(self) -> int:
        return self._entry['properties']['gridProperties']['columnCount']

    @cols.setter
    def cols(self, cols: int) -> None:
        self.set_size(self.rows, cols)

    def _make_single_batch_request(self, method: str, params: Dict) -> Dict:
        spreadsheet_entry = self._spreadsheet._make_single_batch_request(
            method, params)
        for entry in spreadsheet_entry['sheets']:
            if entry['properties']['sheetId'] == self.key:
                return entry
        raise KeyError('Sheet has been removed')
