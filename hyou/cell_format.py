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

import copy
import enum

import six

from . import py3
from . import util


class NumberFormatType(enum.Enum):

    UNSPECIFIED = 'NUMBER_FORMAT_TYPE_UNSPECIFIED'
    TEXT = 'TEXT'
    NUMBER = 'NUMBER'
    PERCENT = 'PERCENT'
    CURRENCY = 'CURRENCY'
    DATE = 'DATE'
    TIME = 'TIME'
    DATE_TIME = 'DATE_TIME'
    SCIENTIFIC = 'SCIENTIFIC'


# TODO(v3): Inherit from collections.namedtuple
class Color(object):

    __slots__ = ['_r', '_g', '_b', '_a']

    # Set later.
    BLACK = None
    WHITE = None

    def __init__(self, r, g, b, a=1.0):
        if not (0 <= r <= 1 and 0 <= g <= 1 and 0 <= b <= 1 and 0 <= a <= 1):
            raise ValueError(
                'Invalid Color: r=%r, g=%r, b=%r, a=%r' % (r, g, b, a))
        self._r = float(r)
        self._g = float(g)
        self._b = float(b)
        self._a = float(a)

    @classmethod
    def _from_data(cls, data):
        return cls(
            data.get('red', 0),
            data.get('green', 0),
            data.get('blue', 0),
            data.get('alpha', 1))

    def _to_data(self):
        return {
            'red': self._r,
            'green': self._g,
            'blue': self._b,
            'alpha': self._a,
        }

    def replace(self, r=None, g=None, b=None, a=None):
        if r is None:
            r = self.r
        if g is None:
            g = self.g
        if b is None:
            b = self.b
        if a is None:
            a = self.a
        return Color(r, g, b, a)

    @property
    def r(self):
        return self._r

    @property
    def g(self):
        return self._g

    @property
    def b(self):
        return self._b

    @property
    def a(self):
        return self._a

    def __repr__(self):
        return str('Color(r=%r, g=%r, b=%r, a=%r)') % (
            self._r, self._g, self._b, self._a)


Color.BLACK = Color(0, 0, 0)
Color.WHITE = Color(1, 1, 1)


class CellFormatProperty(object):

    _next_index = 0

    def __init__(
            self, data_path, type, default_value=None, allow_null=False,
            decoder=None, encoder=None):
        self._data_path = data_path
        self._type = type
        if default_value is not None:
            assert not allow_null
            self._default_value = default_value
            self._explicit_default_value = True
        elif allow_null:
            self._default_value = None
            self._explicit_default_value = True
        else:
            self._default_value = type()
            self._explicit_default_value = False
        self._allow_null = allow_null
        self._decoder = decoder or (lambda x: x)
        self._encoder = encoder or (lambda x: x)
        self._index = self.__class__._next_index
        self.__class__._next_index += 1

    def _get_value_from_data(self, format_data):
        for data_name in self._data_path:
            if format_data is not None:
                format_data = format_data.get(data_name)
        return format_data

    def _set_value_to_data(self, format_data, new_value):
        for data_name in self._data_path[:-1]:
            format_data = format_data.setdefault(data_name, {})
        format_data[self._data_path[-1]] = new_value

    def __get__(self, format, format_class):
        if format is None:
            return self
        format._cell._ensure_fetched_and_valid()
        format_data = self._get_value_from_data(format._current_format)
        if format_data is None:
            return self._default_value
        return self._decoder(format_data)

    def __set__(self, format, new_value):
        if not format._user_entered:
            raise AttributeError('This property is read-only.')
        if new_value is None:
            if not self._explicit_default_value:
                raise ValueError('None is invalid for this property.')
            new_value = self._default_value
        else:
            if self._type is py3.str:
                new_value = py3.promote_str(new_value)
                util.check_type(new_value, py3.str)
            elif self._type in six.integer_types:
                util.check_type(new_value, six.integer_types)
            elif self._type is float:
                if isinstance(new_value, six.integer_types):
                    new_value = float(new_value)
                util.check_type(new_value, float)
            elif issubclass(self._type, enum.Enum):
                if isinstance(new_value, six.string_types):
                    new_value = self._type(py3.promote_str(new_value))
                util.check_type(new_value, self._type)
            elif self._type is Color:
                if isinstance(new_value, tuple) and len(new_value) in (3, 4):
                    new_value = Color(*new_value)
                util.check_type(new_value, self._type)
            else:
                util.check_type(new_value, self._type)
        format._cell._ensure_fetched_and_valid()
        for format_data in (format._current_format, format._pending_format):
            self._set_value_to_data(
                format_data,
                None if new_value is None else self._encoder(new_value))
        format._cell._format_dirty = True


class CellFormat(object):

    __slots__ = [
        '_cell',
        '_current_format',
        '_default_format',
        '_local_format',
        '_pending_format',
        '_user_entered',
    ]

    # Set in _initialize_properties().
    _property_names = None

    def __init__(self, cell, local_format, default_format, user_entered):
        self._cell = cell
        self._user_entered = user_entered
        self._reset_data(local_format, default_format)

    def _reset_data(self, local_format, default_format):
        self._local_format = local_format
        self._default_format = default_format
        self._current_format = self._merge_formats(
            local_format, default_format)
        self._pending_format = {}

    def clear(self):
        # TODO(v3): Make sure this is correct
        self._current_format = copy.deepcopy(self._default_format)
        self._pending_format = copy.deepcopy(self._default_format)
        self._cell._format_dirty = True

    def _get_pending_update(self):
        fields = []
        for property_name in self.__class__._property_names:
            property = getattr(self.__class__, property_name)
            format_data = property._get_value_from_data(self._pending_format)
            if format_data is not None:
                fields.append('.'.join(property._data_path))
        pending_format = copy.deepcopy(self._pending_format)
        # numberFormat must contain both type and pattern, if it ever exists.
        pending_number_format = pending_format.get('numberFormat')
        if pending_number_format:
            pending_number_format.setdefault(
                'type', NumberFormatType.UNSPECIFIED.value)
            pending_number_format.setdefault('pattern', '')
            fields = [f for f in fields if not f.startswith('numberFormat.')]
            fields.append('numberFormat')
        fields.sort()
        return (pending_format, fields)

    @classmethod
    def _initialize_properties(cls):
        property_names = []
        for name in cls.__dict__:
            value = getattr(cls, name)
            if isinstance(value, CellFormatProperty):
                property_names.append(name)
        property_names.sort(key=lambda name: getattr(cls, name)._index)
        cls._property_names = tuple(property_names)

    number_format_type = CellFormatProperty(
        ('numberFormat', 'type'),
        type=NumberFormatType,
        default_value=NumberFormatType.UNSPECIFIED,
        decoder=NumberFormatType,
        encoder=lambda value: value.value)
    number_format_pattern = CellFormatProperty(
        ('numberFormat', 'pattern'),
        type=py3.str,
        allow_null=True)
    background_color = CellFormatProperty(
        ('backgroundColor',),
        type=Color,
        default_value=Color.WHITE,
        decoder=Color._from_data,
        encoder=Color._to_data)
    foreground_color = CellFormatProperty(
        ('textFormat', 'foregroundColor'),
        type=Color,
        default_value=Color.BLACK,
        decoder=Color._from_data,
        encoder=Color._to_data)
    font_family = CellFormatProperty(
        ('textFormat', 'fontFamily'),
        type=py3.str)
    font_size = CellFormatProperty(
        ('textFormat', 'fontSize'),
        type=int)
    bold = CellFormatProperty(
        ('textFormat', 'bold'),
        type=bool)
    italic = CellFormatProperty(
        ('textFormat', 'italic'),
        type=bool)
    strikethrough = CellFormatProperty(
        ('textFormat', 'strikethrough'),
        type=bool)
    underline = CellFormatProperty(
        ('textFormat', 'underline'),
        type=bool)

    def __repr__(self):
        return str('CellFormat(%s)') % str(', ').join(
            str('%s=%r') % (name, getattr(self, name))
            for name in self._property_names)

    @classmethod
    def _merge_formats(cls, local_format, default_format):
        merged_format = {}
        for property_name in cls._property_names:
            property = getattr(cls, property_name)
            local_data = local_format
            default_data = default_format
            merged_data = merged_format
            for data_name in property._data_path[:-1]:
                local_data = local_data.get(data_name, {})
                default_data = default_data.get(data_name, {})
                merged_data = merged_data.setdefault(data_name, {})
            last_path = property._data_path[-1]
            local_data = local_data.get(last_path)
            default_data = default_data.get(last_path)
            if local_data is not None:
                merged_data[last_path] = local_data
            elif default_data is not None:
                merged_data[last_path] = default_data
        return merged_format


CellFormat._initialize_properties()
