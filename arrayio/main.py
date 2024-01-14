from __future__ import annotations

import linecache
from pathlib import Path
from typing import Any

from openpyxl.utils import coordinate_to_tuple
from openpyxl.worksheet.worksheet import Worksheet
from tqdm.autonotebook import tqdm


class Config:
    verbose: bool = True


config = Config()


def set_verbosity(val: bool) -> None:
    config.verbose = val


class Array:
    def __init__(self, values: list[list[Any]] | None = None) -> None:
        self.__values = values
        self._iter_values = self._eager_iter_values

    def get_values(self) -> list[list[Any]]:
        if self.__values is None:
            self.__values = list(
                self._iter_values(
                    prog=config.verbose,
                    desc="Lazy Reading",
                )
            )
        self._iter_values = self._eager_iter_values
        return self.__values

    def write_to_csv(self, fp: Path | str) -> None:
        with Path(fp).open(mode="a", encoding="utf-8") as fo:
            for line in self._iter_values(
                prog=config.verbose,
                desc=("Lazy " if self.__values is None else "") + "Writing",
            ):
                fo.write(", ".join([str(v) for v in line]) + "\n")

    def write_to_excel(self, sheet: Worksheet, cell: str) -> None:
        crow_ini, ccol_ini = coordinate_to_tuple(cell)
        crow = crow_ini
        for row_vals in self._iter_values(
            prog=config.verbose,
            desc=("Lazy " if self.__values is None else "") + "Writing",
        ):
            ccol = ccol_ini
            for val in row_vals:
                sheet.cell(row=crow, column=ccol).value = val
                ccol += 1
            crow += 1

    def _iter_values(self, prog: bool = False, desc: str = "") -> list[Any]:
        raise NotImplementedError

    def _eager_iter_values(
        self, prog: bool = False, desc: str = ""
    ) -> list[Any]:
        for val in tqdm(
            self.__values,
            desc=desc,
            disable=not prog,
        ):
            yield val


def read_csv_array(
    fp: Path | str,
    range_: int | tuple[int],
    lazy: bool = False,
) -> Array:
    fp_ = Path(fp)
    if not fp_.exists():
        raise FileNotFoundError(
            f"No such file or directory: '{fp_.as_posix()}'"
        )

    def _lazy_iter_values(
        self: Array | None = None,
        prog: bool = False,
        desc: str = "",
    ) -> list[Any]:
        linecache.clearcache()

        if isinstance(range_, int):
            from_, to = range_, range_
        else:
            from_, to = range_
        for n in tqdm(
            range(from_, to + 1),
            desc=desc,
            disable=not prog,
        ):
            line = linecache.getline(fp_.as_posix(), lineno=n)
            vstrs = line.replace(",", " ").split()
            vrows = [_try_eval_as_num(v) for v in vstrs]
            yield vrows

    if not lazy:
        values = []
        for val in _lazy_iter_values(
            prog=config.verbose,
            desc="Reading",
        ):
            values.append(val)
        return Array(values=values)
    else:
        array = Array(values=None)
        array._iter_values = _lazy_iter_values
        return array


def read_excel_array(
    sheet: Worksheet,
    range_: str | tuple[str],
    lazy: bool = False,
) -> Array:
    if isinstance(range_, str):
        from_, to = range_, range_
    else:
        from_, to = range_
    crow_ini, ccol_ini = coordinate_to_tuple(from_)
    crow_end, ccol_end = coordinate_to_tuple(to)

    def _lazy_iter_values(
        self: Array | None = None,
        prog: bool = False,
        desc: str = "",
    ) -> list[Any]:
        for row_valls in tqdm(
            sheet.iter_rows(
                min_row=crow_ini,
                max_row=crow_end,
                min_col=ccol_ini,
                max_col=ccol_end,
                values_only=True,
            ),
            desc=desc,
            disable=not prog,
            total=crow_end - crow_ini + 1,
        ):
            yield [_try_eval_as_num(v) for v in row_valls]

    if not lazy:
        values = []
        for val in _lazy_iter_values(
            prog=config.verbose,
            desc="Reading",
        ):
            values.append(val)
        return Array(values=values)
    else:
        array = Array(values=None)
        array._iter_values = _lazy_iter_values
        return array


def _try_eval_as_num(val) -> str | int | float:
    if not isinstance(val, str):
        return val
    try:
        return float(val) if "." in val else int(val)
    except ValueError:
        return val
