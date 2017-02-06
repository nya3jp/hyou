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

import json
from typing import (
    Any, Callable, Generic, Iterable, Iterator, List, Optional, Sequence,
    Tuple, TypeVar, Union)

import oauth2client.client
import oauth2client.service_account


K = TypeVar('K')
V = TypeVar('V')
T = TypeVar('T')

SCOPES = (
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
)


def format_column_address(index_column: int) -> str:
    k = index_column
    p = 1
    while k >= 26 ** p:
        k -= 26 ** p
        p += 1
    s = ''
    for i in range(p):
        s = chr(ord('A') + k % 26) + s
    return s


def format_range_a1_notation(
        worksheet_title: str, start_row: int, end_row: int,
        start_col: int, end_col: int) -> str:
    return '\'%s\'!%s%d:%s%d' % (
        worksheet_title.replace('\'', '\'\''),
        format_column_address(start_col),
        start_row + 1,
        format_column_address(end_col - 1),
        end_row)


def parse_credentials(json_text: str) -> oauth2client.client.Credentials:
    json_data = json.loads(json_text)
    if '_module' in json_data:
        return oauth2client.client.Credentials.new_from_json(
            json_text)
    elif 'private_key' in json_data:
        return (
            oauth2client.service_account.ServiceAccountCredentials
            .from_json_keyfile_dict(
                json_data,
                scopes=SCOPES))
    raise ValueError('unrecognized credential format')


class LazyOrderedDictionary(Generic[K, V]):

    def __init__(self,
                 enumerator: Callable[[], Iterable[Tuple[K, V]]],
                 constructor: Optional[Callable[[K], Optional[V]]]) -> None:
        self._enumerator = enumerator
        self._constructor = constructor
        self._cache_list = []  # type: List[Tuple[K, V]]
        self._cache_index = {}  # type: Dict[K, int]
        self._enumerated = False

    def refresh(self) -> None:
        del self._cache_list[:]
        self._cache_index.clear()
        self._enumerated = False

    def __len__(self) -> int:
        self._ensure_enumerated()
        return len(self._cache_list)

    def __iter__(self) -> Iterable[K]:
        return self.iterkeys()

    def iterkeys(self) -> Iterable[K]:
        self._ensure_enumerated()
        for key, _ in self._cache_list:
            yield key

    def itervalues(self) -> Iterable[V]:
        for _, value in self.iteritems():
            yield value

    def iteritems(self) -> Iterable[Tuple[K, V]]:
        self._ensure_enumerated()
        for key, value in self._cache_list:
            yield (key, value)

    def keys(self) -> List[K]:
        return list(self.iterkeys())

    def values(self) -> List[V]:
        return list(self.itervalues())

    def items(self) -> List[Tuple[K, V]]:
        return list(self.iteritems())

    def __getitem__(self, key: Union[K, int]) -> V:
        if isinstance(key, int):
            self._ensure_enumerated()
            return self._cache_list[key][1]
        index = self._cache_index.get(key)
        if index is not None:
            return self._cache_list[index][1]
        if self._constructor:
            value = self._constructor(key)
            if value is None:
                raise KeyError(key)
            index = len(self._cache_list)
            self._cache_index[key] = index
            self._cache_list.append((key, value))
            return value
        self._ensure_enumerated()
        index = self._cache_index.get(key)
        if index is None:
            raise KeyError(key)
        return self._cache_list[index][1]

    def get(self, key: Union[K, int], default: Any) -> Any:
        try:
            return self[key]
        except KeyError:
            return default

    def _ensure_enumerated(self) -> None:
        if self._enumerated:
            return
        # Save partially constructed entries.
        saves = self._cache_list[:]
        # Initialize cache with the enumerator.
        del self._cache_list[:]
        self._cache_index.clear()
        for key, value in self._enumerator():
            self._cache_index[key] = len(self._cache_list)
            self._cache_list.append((key, value))
        # Restore saved entries.
        for key, value in saves:
            index = self._cache_index.get(key)
            if index is None:
                index = len(self._cache_list)
                self._cache_list.append((None, None))
            self._cache_list[index] = (key, value)
        self._enumerated = True


class CustomMutableFixedList(Generic[T], Sequence[T]):
    """Provides methods to mimic a mutable fixed-size Python list.

    Subclasses need to provide implementation of at least following methods:
    - __getitem__
    - __setitem__
    - __iter__
    - __len__
    """

    def __getitem__(self, index: Any) -> Any:
        raise NotImplementedError()

    def __setitem__(self, index: Any, value: Any) -> None:
        raise NotImplementedError()

    def __iter__(self) -> Iterator[T]:
        raise NotImplementedError()

    def __len__(self) -> int:
        raise NotImplementedError()

    def __eq__(self, other: Sequence[T]) -> bool:  # type: ignore
        if len(self) != len(other):
            return False
        for a, b in zip(self, other):
            if a != b:
                return False
        return True

    def __ne__(self, other: Sequence[T]) -> bool:  # type: ignore
        return not (self == other)

    def __lt__(self, other: Sequence[T]) -> bool:
        for a, b in zip(self, other):
            if a != b:
                return a < b  # type: ignore
        return len(self) < len(other)

    def __le__(self, other: Sequence[T]) -> bool:
        for a, b in zip(self, other):
            if a != b:
                return a < b  # type: ignore
        return len(self) <= len(other)

    def __gt__(self, other: Sequence[T]) -> bool:
        return not (self <= other)

    def __ge__(self, other: Sequence[T]) -> bool:
        return not (self < other)

    def __contains__(self, find_value: T) -> bool:  # type: ignore
        for value in self:
            if value == find_value:
                return True
        return False

    def count(self, find_value: T) -> int:
        result = 0
        for value in self:
            if value == find_value:
                result += 1
        return result

    def index(self, find_value: T) -> int:  # type: ignore
        for i, value in enumerate(self):
            if value == find_value:
                return i
        raise ValueError('%r is not in list' % find_value)

    def reverse(self) -> None:
        for i, new_value in enumerate(list(reversed(self))):
            self[i] = new_value

    def sort(self,
             key: Optional[Callable[[T], Any]]=None,
             reverse: bool=False) -> None:
        for i, new_value in enumerate(sorted(
                self, key=key, reverse=reverse)):
            self[i] = new_value

    def __delitem__(self, key: Any) -> None:
        raise NotImplementedError(
            'Methods changing the list size are unavailable')

    def append(self, x: T) -> None:
        raise NotImplementedError(
            'Methods changing the list size are unavailable')

    def extend(self, x: Iterable[T]) -> None:
        raise NotImplementedError(
            'Methods changing the list size are unavailable')

    def insert(self, i: int, x: T) -> None:
        raise NotImplementedError(
            'Methods changing the list size are unavailable')

    def pop(self, i: int=None) -> T:
        raise NotImplementedError(
            'Methods changing the list size are unavailable')

    def remove(self, x: T) -> None:
        raise NotImplementedError(
            'Methods changing the list size are unavailable')
