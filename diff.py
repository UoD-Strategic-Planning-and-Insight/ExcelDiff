from dataclasses import dataclass
from typing import Any

import openpyxl
from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table

import Utils
import Utils.boilerplate as bp

from xltables import XLTable

@dataclass
class TableReference:
    filepath:   str
    sheet_name: str
    table_name: str


@dataclass
class CellDifference:
    column_name: str
    value1:      Any
    value2:      Any


@dataclass
class RowDifference:
    keys:             dict[str, Any]
    cell_differences: list[CellDifference]


@dataclass
class TableColumnContent:
    column_name: str
    values:      list[Any]


class TableDiff:
    first_table_ref:  TableReference
    second_table_ref: TableReference
    result_filepath:  str
    key_column_names: list[str]

    first_table:  XLTable | None
    second_table: XLTable | None

    row_numbers_for_key_sets_in_first:  dict[str, int]
    row_numbers_for_key_sets_in_second: dict[str, int]

    row_differences:        list[RowDifference]
    rows_only_in_first:     list[dict[str, Cell]]
    rows_only_in_second:    list[dict[str, Cell]]
    columns_only_in_first:  list[TableColumnContent]
    columns_only_in_second: list[TableColumnContent]

    def __init__(self,
                 first:            TableReference,
                 second:           TableReference,
                 result_filepath:  str,
                 key_column_names: list[str]):

        self.first_table_ref  = first
        self.second_table_ref = second
        self.result_filepath  = result_filepath
        self.key_column_names = key_column_names

        self.row_numbers_for_key_sets_in_first  = {}
        self.row_numbers_for_key_sets_in_second = {}

        self.row_differences        = []
        self.rows_only_in_first     = []
        self.rows_only_in_second    = []
        self.columns_only_in_first  = []
        self.columns_only_in_second = []

    def process_and_save(self):

        self.load_tables()
        self.build_table_indices()
        self.read_row_differences()
        self.read_rows_only_in_second()
        self.read_columns_only_in_first()
        self.read_columns_only_in_second()
        self.save_to_file()
        self.discard_loaded_tables()

    def load_tables(self):
        ref1              = self.first_table_ref
        ref2              = self.second_table_ref
        self.first_table  = bp.load_table(ref1.filepath, ref1.sheet_name, ref1.table_name)
        self.second_table = bp.load_table(ref2.filepath, ref2.sheet_name, ref2.table_name)

    def discard_loaded_tables(self):
        self.first_table.source_workbook.close()
        self.second_table.source_workbook.close()
        self.first_table  = None
        self.second_table = None

    def build_table_indices(self):
        self._build_row_index(self.first_table,  self.row_numbers_for_key_sets_in_first)
        self._build_row_index(self.second_table, self.row_numbers_for_key_sets_in_second)

    def read_row_differences(self):
        self.row_differences    = []
        self.rows_only_in_first = []

        for row in self.first_table.row_iterator:
            keys: dict[str, Any] = {}

            for key_col_name in self.key_column_names:
                keys[key_col_name] = row[key_col_name].value

            matching_row_in_second: dict[str, Cell] | None \
                = self._get_row_with_keys(self.second_table, keys, self.row_numbers_for_key_sets_in_second)

            if(matching_row_in_second is None):
                self.rows_only_in_first.append(row)
                continue

            cell_diffs: list[CellDifference] = self._get_differences_between_rows(row, matching_row_in_second)

            if(len(cell_diffs) != 0):
                self.row_differences.append(RowDifference(keys, cell_diffs))

    def read_rows_only_in_second(self):
        self.rows_only_in_second = []

        for row in self.second_table.row_iterator:
            keys: dict[str, Any] = {}

            for key_col_name in self.key_column_names:
                keys[key_col_name] = row[key_col_name].value

            matching_row_exists_in_first: bool \
                = self.row_numbers_for_key_sets_in_first.get(Utils.dict_to_str(keys)) is not None

            if(not matching_row_exists_in_first):
                self.rows_only_in_second.append(row)

    def read_columns_only_in_first(self):
        self.columns_only_in_first = self._get_columns_not_in_other(self.first_table, self.second_table)

    def read_columns_only_in_second(self):
        self.columns_only_in_second = self._get_columns_not_in_other(self.second_table, self.first_table)

    def save_to_file(self):
        wb = openpyxl.Workbook()

        self._add_diffs_sheet_to_workbook(wb)
        self._add_rows_only_in_one_sheet_to_workbook(wb, self.rows_only_in_first,
                                                     2, "Rows unique to first", "RowsUniqueToFirst")

        self._add_rows_only_in_one_sheet_to_workbook(wb, self.rows_only_in_second,
                                                     3, "Rows unique to second", "RowsUniqueToSecond")

        key_cols_in_first:  list[TableColumnContent] = self._get_key_columns(self.first_table)
        key_cols_in_second: list[TableColumnContent] = self._get_key_columns(self.second_table)

        self._add_columns_only_in_one_sheet_to_workbook(wb, key_cols_in_first, self.columns_only_in_first,
                                                        4, "Columns unique to first", "ColumnsUniqueToFirst")

        self._add_columns_only_in_one_sheet_to_workbook(wb, key_cols_in_second, self.columns_only_in_second,
                                                        5, "Columns unique to second", "ColumnsUniqueToSecond")

        wb.remove_sheet(wb.get_sheet_by_name("Sheet"))
        wb.save(self.result_filepath)

    def _add_diffs_sheet_to_workbook(self, wb: Workbook):
        if(len(self.row_differences) == 0):
            return

        wb.create_sheet("Differences", 1)
        sheet = wb.get_sheet_by_name("Differences")

        for i in range(len(self.key_column_names)):
            sheet.cell(1, i + 1).value = self.key_column_names[i]

        table = Table(displayName = "DiffTable",
                      ref         = f"A1:{Utils.convert_int_to_alphabetic_number(len(self.key_column_names))}1")

        sheet.add_table(table)
        tbl_diff = XLTable(self.result_filepath, wb, sheet, table)

        for diff in self.row_differences:
            tbl_diff.add_row()
            row = tbl_diff.bottom_row

            for k, v in diff.keys.items():
                row[k].value = v

            for cell_diff in diff.cell_differences:
                col_name_1 = cell_diff.column_name + " * 1"
                col_name_2 = cell_diff.column_name + " * 2"
                columns_added: bool = False

                if(not tbl_diff.has_column(col_name_1)):
                    tbl_diff.add_column(col_name_1)
                    tbl_diff.add_column(col_name_2)
                    columns_added = True

                if(columns_added):
                    row = tbl_diff.bottom_row

                row[col_name_1].value = cell_diff.value1
                row[col_name_2].value = cell_diff.value2

    def _add_rows_only_in_one_sheet_to_workbook(self,
                                                wb:          Workbook,
                                                rows:        list[dict[str, Cell]],
                                                sheet_index: int,
                                                sheet_name:  str,
                                                table_name:  str):
        if(len(rows) == 0):
            return

        Workbook.create_sheet(wb, sheet_name, sheet_index)
        sheet: Worksheet = wb.get_sheet_by_name(sheet_name)
        key_column_count = len(self.key_column_names)

        for i in range(key_column_count):
            sheet.cell(1, i + 1).value = self.key_column_names[i]

        non_key_column_names = [x for x in rows[0].keys() if x not in self.key_column_names]
        table_width = len(rows[0].keys())

        for i in range(len(non_key_column_names)):
            sheet.cell(1, key_column_count + i + 1).value = non_key_column_names[i]

        table = Table(displayName = table_name,
                      ref         = f"A1:{Utils.convert_int_to_alphabetic_number(table_width)}1")

        sheet.add_table(table)
        tbl = XLTable(self.result_filepath, wb, sheet, table)

        for source_row in rows:
            tbl.add_row()
            dest_row = tbl.bottom_row

            for k, v in source_row.items():
                dest_row[k].value = v.value

    def _add_columns_only_in_one_sheet_to_workbook(self,
                                                   wb:          Workbook,
                                                   key_columns: list[TableColumnContent],
                                                   columns:     list[TableColumnContent],
                                                   sheet_index: int,
                                                   sheet_name:  str,
                                                   table_name:  str):

        if(len(columns) == 0):
            return

        wb.create_sheet(sheet_name, sheet_index)
        sheet = wb.get_sheet_by_name(sheet_name)
        row_count = len(columns[0].values)

        for i in range(len(key_columns)):
            col: TableColumnContent = key_columns[i]
            sheet.cell(1, i + 1).value = col.column_name

            for j in range(len(col.values)):
                sheet.cell(j + 2, i + 1).value = col.values[j]

        table = Table(displayName=table_name,
                      ref=f"A1:{Utils.convert_int_to_alphabetic_number(len(key_columns))}{row_count + 1}")

        sheet.add_table(table)
        tbl = XLTable(self.result_filepath, wb, sheet, table)

        for col in columns:
            tbl.add_column(col.column_name, col.values)

    def _build_row_index(self, table: XLTable, cache: dict[str, int]):
        row_no: int = -1

        for row in table.row_iterator:
            keys: dict[str, str] = {}
            row_no += 1

            for key_col_name in self.key_column_names:
                keys[key_col_name] = row[key_col_name].value

            key_str: str = Utils.dict_to_str(keys)
            cache[key_str] = row_no

    def _get_key_columns(self, table: XLTable) -> list[TableColumnContent]:
        result: list[TableColumnContent] = []

        for col_name in self.key_column_names:
            col_vals: list[Any] = []

            for row in table.row_iterator:
                col_vals.append(row[col_name].value)

            result.append(TableColumnContent(col_name, col_vals))

        return result

    def _get_row_with_keys(self, table: XLTable, keys: dict[str, Any], row_number_lookup_dict: dict[str, int])\
            -> dict[str, Cell] | None:

        key_string: str = Utils.dict_to_str(keys)
        row_no: int | None = row_number_lookup_dict.get(key_string)
        return (table.get_row(row_no)) if (row_no is not None) else (None)

    def _get_differences_between_rows(self, first: dict[str, Cell], second: dict[str, Cell]) \
            -> list[CellDifference]:

        result: list[CellDifference] = []

        for k, v1 in first.items():
            v2: Cell | None = second.get(k)

            if(v2 is None):
                continue

            v1val = str(v1.value).strip() if v1.value is not None else ""
            v2val = str(v2.value).strip() if v2.value is not None else ""

            if(v1val != v2val):
                result.append(CellDifference(k, v1val, v2val))

        return result

    def _get_columns_not_in_other(self, table: XLTable, other_table: XLTable) \
            -> list[TableColumnContent]:

        col_names_1 = table.column_names
        col_names_2 = other_table.column_names

        cols_not_in_other: list[TableColumnContent] = []

        for col_name in col_names_1:
            if(col_name not in col_names_2):
                col_vals: list[Any] = []

                for row in table.row_iterator:
                    col_vals.append(row[col_name].value)

                cols_not_in_other.append(TableColumnContent(col_name, col_vals))

        return cols_not_in_other
