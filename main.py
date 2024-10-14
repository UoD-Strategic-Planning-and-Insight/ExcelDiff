from diff import TableDiff, TableReference


def main():

    # The following is an example

    folder_path: str = (r"C:\Users\yourusername\Desktop\Example folder")
    path1: str = f"{folder_path}\\Old file.xlsx"
    path2: str = f"{folder_path}\\New file.xlsx"
    resultPath: str = f"{folder_path}\\Diff.xlsx"

    diff: TableDiff = TableDiff(TableReference(path1, "Sheet1", "Table1"),
                                TableReference(path2, "Sheet1", "Table1"),
                                result_filepath  = resultPath,
                                key_column_names = ["Key1" ,"Key2", "Key3"])

    diff.process_and_save()


if __name__ == '__main__':
    main()
