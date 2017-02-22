# Copyright 2017 Google Inc. All rights reserved
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

import six

from . import cell
from . import py3
from . import util


class View(util.CustomMutableFixedList):

    def __init__(self, worksheet, api, start_row, end_row, start_col, end_col):
        self._worksheet = worksheet
        self._api = api
        self._start_row = start_row
        self._end_row = end_row
        self._start_col = start_col
        self._end_col = end_col
        self._view_rows = [
            ViewRow(self, row, start_col, end_col)
            for row in py3.range(start_row, end_row)]
        self._cell_map = {}
        self._default_format = None
        self._fetched = False

    def refresh(self):
        self._default_format = None
        self._fetched = False

    def _get_cell(self, row, col):
        if not (self._start_row <= row < self._end_row):
            raise IndexError('Row %d is out of this view' % row)
        if not (self._start_col <= col < self._end_col):
            raise IndexError('Column %d is out of this view' % col)
        self._ensure_fetched()
        location = (row, col)
        if location not in self._cell_map:
            self._cell_map[location] = cell.Cell(
                self, row, col, {}, self._default_format)
        return self._cell_map[location]

    def _ensure_fetched(self):
        if self._fetched:
            return
        response = self._fetch()
        # Since this function is recursively called inside the constructor
        # of Cell, we need to set self._fetched first to avoid infinite
        # recursion.
        self._fetched = True
        self._default_format = response['properties']['defaultFormat']
        cleared_locations = set(self._cell_map.keys())
        range_data = response['sheets'][0]['data'][0]
        for i, row_data in enumerate(range_data['rowData']):
            row = self._start_row + i
            for j, cell_data in enumerate(row_data.get('values', [])):
                col = self._start_col + j
                location = (row, col)
                cleared_locations.discard(location)
                if location not in self._cell_map:
                    self._cell_map[location] = cell.Cell(
                        self, row, col, cell_data, self._default_format)
                else:
                    self._cell_map[location]._reset_data(
                        cell_data, self._default_format)
        for location in cleared_locations:
            self._cell_map[location]._reset_data({}, self._default_format)

    def _fetch(self):
        range_str = util.format_range_a1_notation(
            self._worksheet.title, self._start_row, self._end_row,
            self._start_col, self._end_col)
        response = self._api.sheets.spreadsheets().get(
            spreadsheetId=self._worksheet._spreadsheet.key,
            ranges=py3.str_to_native_str(range_str),
            includeGridData=True).execute()
        return response

    def commit(self):
        update_requests = []
        for c in self._cell_map.values():
            update_request = c._build_update_request()
            if update_request:
                update_requests.append(update_request)
        if not update_requests:
            return
        request = {
            'requests': update_requests,
            'include_spreadsheet_in_response': False,
        }
        self._api.sheets.spreadsheets().batchUpdate(
            spreadsheetId=self._worksheet._spreadsheet.key,
            body=request).execute()
        self.refresh()

    def __getitem__(self, index):
        return self._view_rows[index]

    def __setitem__(self, index, new_value):
        if isinstance(index, slice):
            start, stop, step = index.indices(len(self))
            assert step == 1, 'slicing with step is not supported'
            if stop < start:
                stop = start
            if len(new_value) != stop - start:
                raise ValueError(
                    'Tried to assign %d values to %d element slice' %
                    (len(new_value), stop - start))
            for i, new_value_one in py3.zip(py3.range(start, stop), new_value):
                self[i] = new_value_one
            return
        self._view_rows[index][:] = new_value

    def __len__(self):
        return self.rows

    def __iter__(self):
        return iter(self._view_rows)

    def __repr__(self):
        return str('View(%r)') % (self._view_rows,)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.commit()

    @property
    def rows(self):
        return self._end_row - self._start_row

    @property
    def cols(self):
        return self._end_col - self._start_col

    @property
    def start_row(self):
        return self._start_row

    @property
    def end_row(self):
        return self._end_row

    @property
    def start_col(self):
        return self._start_col

    @property
    def end_col(self):
        return self._end_col


class ViewRow(util.CustomMutableFixedList):

    def __init__(self, view, row, start_col, end_col):
        self._view = view
        self._row = row
        self._start_col = start_col
        self._end_col = end_col

    def __getitem__(self, index):
        if isinstance(index, slice):
            start, stop, step = index.indices(len(self))
            assert step == 1, 'slicing with step is not supported'
            if stop < start:
                stop = start
            return ViewRow(
                self._view, self._row,
                self._start_col + start, self._start_col + stop)
        assert isinstance(index, six.integer_types)
        if index < 0:
            col = self._end_col + index
        else:
            col = self._start_col + index
        return self._view._get_cell(self._row, col)

    def __setitem__(self, index, new_value):
        if isinstance(index, slice):
            start, stop, step = index.indices(len(self))
            assert step == 1, 'slicing with step is not supported'
            if stop < start:
                stop = start
            if len(new_value) != stop - start:
                raise ValueError(
                    'Tried to assign %d values to %d element slice' %
                    (len(new_value), stop - start))
            for i, new_value_one in py3.zip(py3.range(start, stop), new_value):
                self[i] = new_value_one
            return
        assert isinstance(index, six.integer_types)
        cell = self[index]
        cell._assign_value(new_value)

    def __len__(self):
        return self._end_col - self._start_col

    def __iter__(self):
        for i in py3.range(len(self)):
            yield self[i]

    def __repr__(self):
        return repr(list(self))
