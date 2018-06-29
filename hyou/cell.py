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

import datetime

import six

from . import cell_format
from . import exception
from . import py3
from . import util


class Cell(object):

    __slots__ = [
        '_col',
        '_effective_format',
        '_effective_value',
        '_format_dirty',
        '_formatted_value',
        '_row',
        '_user_entered_format',
        '_user_entered_value',
        '_value_dirty',
        '_view',
    ]

    def __init__(self, view, row, col, cell_data, default_format):
        super(Cell, self).__init__()
        self._view = view
        self._row = row
        self._col = col
        self._user_entered_format = cell_format.CellFormat(
            self, cell_data.get('userEnteredFormat', {}), default_format,
            user_entered=True)
        self._effective_format = cell_format.CellFormat(
            self, cell_data.get('effectiveFormat', {}), default_format,
            user_entered=False)
        self._user_entered_value = self._parse_extended_value(
            cell_data.get('userEnteredValue'), self._user_entered_format)
        self._effective_value = self._parse_extended_value(
            cell_data.get('effectiveValue'), self._effective_format)
        self._formatted_value = cell_data.get('formattedValue', '')
        self._value_dirty = False
        self._format_dirty = False

    def _reset_data(self, cell_data, default_format):
        self._user_entered_format._reset_data(
            cell_data.get('userEnteredFormat', {}), default_format)
        self._effective_format._reset_data(
            cell_data.get('effectiveFormat', {}), default_format)
        self._user_entered_value = self._parse_extended_value(
            cell_data.get('userEnteredValue'), self._user_entered_format)
        self._effective_value = self._parse_extended_value(
            cell_data.get('effectiveValue'), self._effective_format)
        self._formatted_value = cell_data.get('formattedValue', '')
        self._value_dirty = False
        self._format_dirty = False

    def _ensure_fetched_and_valid(self):
        self._view._ensure_fetched()
        if not (self._view._start_row <= self._row < self._view._end_row and
                self._view._start_col <= self._col < self._view._end_col):
            raise exception.HyouRuntimeError('This cell no longer exists.')

    def _assign_value(self, new_value):
        self._ensure_fetched_and_valid()
        _, number_format_type = self._unparse_extended_value(new_value)
        if number_format_type:
            self._user_entered_format.number_format_type = number_format_type
        self._user_entered_value = new_value
        self._value_dirty = True

    def _build_update_request(self):
        if not (self._value_dirty or self._format_dirty):
            return None
        value_data = {}
        fields = []
        if self._value_dirty:
            extended_value, _ = self._unparse_extended_value(
                self._user_entered_value)
            value_data['userEnteredValue'] = extended_value
            fields.append('userEnteredValue')
        if self._format_dirty:
            pending_format, format_fields = (
                self._user_entered_format._get_pending_update())
            value_data['userEnteredFormat'] = pending_format
            fields.extend('userEnteredFormat.%s' % f for f in format_fields)
        request = {
            'updateCells': {
                'start': {
                    'sheetId': self._view._worksheet.key,
                    'rowIndex': self._row,
                    'columnIndex': self._col,
                },
                'rows': [{'values': [value_data]}],
                'fields': ','.join(fields)
            },
        }
        return request

    def __repr__(self):
        return str('Cell(%r)') % (self.user_entered_value,)

    def __str__(self):
        return str(self.user_entered_value)

    def __unicode__(self):
        return py3.str(self.user_entered_value)

    def clear(self):
        self.user_entered_value = None
        self.user_entered_format.clear()

    @property
    def row(self):
        # Do not call self._ensure_fetched_and_valid() here.
        return self._row

    @property
    def col(self):
        # Do not call self._ensure_fetched_and_valid() here.
        return self._col

    @property
    def user_entered_value(self):
        self._ensure_fetched_and_valid()
        return self._user_entered_value

    @user_entered_value.setter
    def user_entered_value(self, new_value):
        self._assign_value(new_value)

    @property
    def effective_value(self):
        self._ensure_fetched_and_valid()
        if self._value_dirty or self._format_dirty:
            raise exception.UncommittedCellPropertyError('effective_value')
        return self._effective_value

    @property
    def formatted_value(self):
        self._ensure_fetched_and_valid()
        if self._value_dirty or self._format_dirty:
            raise exception.UncommittedCellPropertyError('formatted_value')
        return self._formatted_value

    @property
    def user_entered_format(self):
        self._ensure_fetched_and_valid()
        return self._user_entered_format

    @property
    def effective_format(self):
        self._ensure_fetched_and_valid()
        if self._value_dirty or self._format_dirty:
            raise exception.UncommittedCellPropertyError('effective_format')
        return self._effective_format

    @classmethod
    def _parse_extended_value(cls, extended_value, format):
        if not extended_value:
            return None
        if 'stringValue' in extended_value:
            return extended_value['stringValue']
        if 'boolValue' in extended_value:
            return extended_value['boolValue']
        if 'formulaValue' in extended_value:
            return FormulaValue(extended_value['formulaValue'])
        if 'errorValue' in extended_value:
            return ErrorValue(extended_value['errorValue'])
        number = extended_value['numberValue']
        if format.number_format_type == cell_format.NumberFormatType.DATE_TIME:
            return util.serial_to_datetime(number)
        if format.number_format_type == cell_format.NumberFormatType.DATE:
            return util.serial_to_datetime(number).date()
        if format.number_format_type == cell_format.NumberFormatType.TIME:
            return util.serial_to_datetime(number).time()
        return number

    @classmethod
    def _unparse_extended_value(cls, value):
        value = py3.promote_to_str(value)
        if value is None:
            return ({}, None)
        if isinstance(value, py3.str):
            return ({'stringValue': value}, None)
        if isinstance(value, bool):
            return ({'boolValue': value}, None)
        if isinstance(value, FormulaValue):
            return ({'formulaValue': py3.str(value)}, None)
        if isinstance(value, ErrorValue):
            raise ValueError('ErrorValue can not be unparsed.')
        if isinstance(value, six.integer_types):
            return ({'numberValue': value}, None)
        if isinstance(value, datetime.datetime):
            serial = util.datetime_to_serial(value)
            return (
                {'numberValue': serial},
                cell_format.NumberFormatType.DATE_TIME)
        if isinstance(value, datetime.date):
            serial = util.datetime_to_serial(datetime.datetime(
                value.year, value.month, value.day))
            return ({'numberValue': serial}, cell_format.NumberFormatType.DATE)
        if isinstance(value, datetime.time):
            serial = util.datetime_to_serial(
                util.SERIAL_EPOCH +
                datetime.timedelta(
                    seconds=(
                        (value.hour * 60 + value.minute) * 60 + value.second),
                    microseconds=value.microsecond))
            return ({'numberValue': serial}, cell_format.NumberFormatType.TIME)
        raise ValueError('Can not unparse %r' % value)


class ErrorValue(object):

    __slots__ = ['_error']

    def __init__(self, error):
        error = py3.promote_to_str(error)
        if not isinstance(error, py3.str):
            raise TypeError('Error message must be a string.')
        self._error = error

    @property
    def error(self):
        return self._error

    def __repr__(self):
        return str('ErrorValue(error=%r)') % (self._error,)


class FormulaValue(object):

    __slots__ = ['_formula']

    def __init__(self, formula):
        formula = py3.promote_to_str(formula)
        if not isinstance(formula, py3.str):
            raise TypeError('Formula value must be a string.')
        self._formula = formula

    @property
    def formula(self):
        return self._formula

    def __repr__(self):
        return str('FormulaValue(formula=%r)') % (self._formula,)
