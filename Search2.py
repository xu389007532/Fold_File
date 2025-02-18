from pathlib import Path


def list_files(directory):
    for path in Path(directory).rglob('*'):  # rglob 用于递归地搜索所有文件和目录
        print(path)

        if path.is_file():  # 确认是文件而不是目录
            pass


# path = '//hmpfs01/DG3 Public Library/ITProgram/PDFtypesetting'
# 指定你要搜索的目录路径
directory_path = Path('//hmpfs01/DG3 Public Library/ITProgram/PDFtypesetting')
list_files(directory_path)